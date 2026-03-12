using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.ComponentModel;
using App.Core.Runtime;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace App.Inference;

public sealed class KittenTtsOnnxBackend : ITtsBackend, IDisposable
{
    private const int OutputSampleRate = 24000;
    private const int StyleDim = 256;
    private static readonly object CudaBootstrapSync = new();
    private static bool _cudaBootstrapDone;
    private static readonly Regex BasicTokenRegex = new(@"\w+|[^\w\s]", RegexOptions.Compiled);
    private static readonly Regex MultiWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex LeadingDecimalRegex = new(@"(?<!\d)\.(\d)", RegexOptions.Compiled);

    private readonly LocalInferenceOptions _options;
    private InferenceSession? _session;
    private string _loadedKey = string.Empty;
    private KittenModelConfig? _config;
    private Dictionary<string, float[]> _voices = new(StringComparer.OrdinalIgnoreCase);

    public KittenTtsOnnxBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "kitten-tts-mini-onnx-native";

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
        => Task.Run(() => SynthesizeInternal(request, ct), ct);

    public void Dispose()
    {
        _session?.Dispose();
        _session = null;
        _loadedKey = string.Empty;
        _config = null;
        _voices = new Dictionary<string, float[]>(StringComparer.OrdinalIgnoreCase);
        GC.SuppressFinalize(this);
    }

    public static IReadOnlyList<(string Alias, string InternalId)> GetDefaultVoices()
        => new (string Alias, string InternalId)[]
        {
            ("Bella", "expr-voice-2-f"),
            ("Jasper", "expr-voice-2-m"),
            ("Luna", "expr-voice-3-f"),
            ("Bruno", "expr-voice-3-m"),
            ("Rosie", "expr-voice-4-f"),
            ("Hugo", "expr-voice-4-m"),
            ("Kiki", "expr-voice-5-f"),
            ("Leo", "expr-voice-5-m")
        };

    private void SynthesizeInternal(TtsRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            throw new ArgumentException("Input text is empty.");
        }
        if (string.IsNullOrWhiteSpace(request.OutputPath))
        {
            throw new ArgumentException("Output path is required.");
        }

        var outDir = Path.GetDirectoryName(request.OutputPath);
        if (!string.IsNullOrWhiteSpace(outDir))
        {
            Directory.CreateDirectory(outDir);
        }

        var repoId = string.IsNullOrWhiteSpace(_options.ModelRepoId)
            ? "KittenML/kitten-tts-mini-0.8"
            : _options.ModelRepoId.Trim();
        var cacheDir = ModelCachePath.ResolveAbsolute(_options.ModelCacheDir, RuntimePathResolver.AppRoot);
        EnsureLoaded(cacheDir, repoId);

        if (_session is null || _config is null)
        {
            throw new InvalidOperationException("Kitten TTS backend failed to initialize.");
        }

        var voiceKey = ResolveVoiceKey(request.VoicePath, _config);
        if (!_voices.TryGetValue(voiceKey, out var styleRows) || styleRows.Length < StyleDim)
        {
            throw new InvalidOperationException($"Kitten built-in voice '{voiceKey}' not found in voices.npz.");
        }

        var text = NormalizeInputText(request.Text);
        var tokenIds = BuildTokenIds(text);
        if (tokenIds.Length < 3)
        {
            throw new InvalidOperationException("Kitten tokenizer produced too few tokens.");
        }

        var refRow = Math.Clamp(text.Length, 0, Math.Max(0, (styleRows.Length / StyleDim) - 1));
        var style = new float[StyleDim];
        Array.Copy(styleRows, refRow * StyleDim, style, 0, StyleDim);

        var speed = Math.Clamp(request.Speed, 0.5f, 1.5f);
        if (_config.SpeedPriors.TryGetValue(voiceKey, out var prior) && prior > 0)
        {
            speed *= prior;
        }
        speed = Math.Clamp(speed, 0.35f, 2.5f);

        ct.ThrowIfCancellationRequested();
        var waveform = RunModel(_session, tokenIds, style, speed, ct);
        WriteWav16Mono(request.OutputPath, waveform, OutputSampleRate);
    }

    private void EnsureLoaded(string cacheDir, string repoId)
    {
        var parts = repoId.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            throw new InvalidOperationException($"Invalid model repo id: {repoId}");
        }

        var repoRoot = Path.Combine(cacheDir, "hf-cache", $"models--{parts[0]}--{parts[1]}");
        var key = $"{repoRoot.ToLowerInvariant()}|{NormalizePreferDevice(_options.PreferDevice)}";
        if (_loadedKey == key && _session is not null && _config is not null && _voices.Count > 0)
        {
            return;
        }

        Dispose();

        if (!Directory.Exists(repoRoot))
        {
            throw new DirectoryNotFoundException($"Kitten model repo not found in cache: {repoRoot}");
        }

        var configPath = Path.Combine(repoRoot, "config.json");
        var onnxPath = Path.Combine(repoRoot, "kitten_tts_mini_v0_8.onnx");
        var voicesPath = Path.Combine(repoRoot, "voices.npz");
        if (!File.Exists(configPath) || !File.Exists(onnxPath) || !File.Exists(voicesPath))
        {
            throw new FileNotFoundException(
                "Kitten model files are incomplete. Required: config.json, kitten_tts_mini_v0_8.onnx, voices.npz.");
        }

        _config = LoadConfig(configPath);
        _voices = LoadVoicesNpz(voicesPath);
        if (_voices.Count == 0)
        {
            throw new InvalidOperationException("voices.npz contains no voice embeddings.");
        }

        var sessionOptions = CreateSessionOptions(_options.PreferDevice);
        _session = new InferenceSession(onnxPath, sessionOptions);
        _loadedKey = key;
    }

    private static KittenModelConfig LoadConfig(string configPath)
    {
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var speedPriors = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
            var root = doc.RootElement;

            if (root.TryGetProperty("voice_aliases", out var voiceAliases) && voiceAliases.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in voiceAliases.EnumerateObject())
                {
                    var value = prop.Value.GetString();
                    if (!string.IsNullOrWhiteSpace(prop.Name) && !string.IsNullOrWhiteSpace(value))
                    {
                        aliases[prop.Name.Trim()] = value.Trim();
                    }
                }
            }

            if (root.TryGetProperty("speed_priors", out var priors) && priors.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in priors.EnumerateObject())
                {
                    if (prop.Value.TryGetSingle(out var f) && f > 0)
                    {
                        speedPriors[prop.Name.Trim()] = f;
                    }
                    else if (prop.Value.ValueKind == JsonValueKind.Number)
                    {
                        speedPriors[prop.Name.Trim()] = (float)Math.Clamp(prop.Value.GetDouble(), 0.1, 4.0);
                    }
                }
            }
        }
        catch
        {
            // Fall back to defaults below.
        }

        foreach (var (alias, internalId) in GetDefaultVoices())
        {
            if (!aliases.ContainsKey(alias))
            {
                aliases[alias] = internalId;
            }
        }

        return new KittenModelConfig(aliases, speedPriors);
    }

    private static Dictionary<string, float[]> LoadVoicesNpz(string npzPath)
    {
        var map = new Dictionary<string, float[]>(StringComparer.OrdinalIgnoreCase);
        using var fs = File.OpenRead(npzPath);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.EndsWith(".npy", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var stream = entry.Open();
            var (data, shape) = ReadNpyFloat32(stream);
            if (shape.Length != 2 || shape[1] != StyleDim || data.Length < StyleDim)
            {
                continue;
            }

            var key = Path.GetFileNameWithoutExtension(entry.Name);
            if (!string.IsNullOrWhiteSpace(key))
            {
                map[key] = data;
            }
        }

        return map;
    }

    private static (float[] Data, int[] Shape) ReadNpyFloat32(Stream stream)
    {
        using var br = new BinaryReader(stream, Encoding.ASCII, leaveOpen: true);
        var magic = br.ReadBytes(6);
        if (magic.Length != 6 || magic[0] != 0x93 || Encoding.ASCII.GetString(magic, 1, 5) != "NUMPY")
        {
            throw new InvalidOperationException("Invalid NPY header.");
        }

        var major = br.ReadByte();
        var minor = br.ReadByte();
        var headerLength = major switch
        {
            1 => br.ReadUInt16(),
            2 or 3 => (int)br.ReadUInt32(),
            _ => throw new InvalidOperationException($"Unsupported NPY version {major}.{minor}.")
        };

        var headerBytes = br.ReadBytes(headerLength);
        var header = Encoding.ASCII.GetString(headerBytes);
        if (!header.Contains("'descr': '<f4'") && !header.Contains("\"descr\": \"<f4\""))
        {
            throw new InvalidOperationException("NPY array is not float32 little-endian.");
        }
        if (header.Contains("True", StringComparison.Ordinal) && header.Contains("fortran_order", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Fortran-ordered NPY arrays are not supported.");
        }

        var shapeMatch = Regex.Match(header, @"shape'\s*:\s*\(([^)]*)\)");
        if (!shapeMatch.Success)
        {
            shapeMatch = Regex.Match(header, "\"shape\"\\s*:\\s*\\(([^)]*)\\)");
        }
        if (!shapeMatch.Success)
        {
            throw new InvalidOperationException("NPY shape metadata missing.");
        }

        var dims = shapeMatch.Groups[1].Value
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => int.Parse(x.Trim(), System.Globalization.CultureInfo.InvariantCulture))
            .ToArray();
        if (dims.Length == 0)
        {
            throw new InvalidOperationException("NPY shape is empty.");
        }

        var count = 1;
        foreach (var d in dims)
        {
            count = checked(count * Math.Max(1, d));
        }

        var bytes = br.ReadBytes(count * sizeof(float));
        if (bytes.Length < count * sizeof(float))
        {
            throw new InvalidOperationException("NPY data is truncated.");
        }

        var data = new float[count];
        Buffer.BlockCopy(bytes, 0, data, 0, bytes.Length);
        return (data, dims);
    }

    private static float[] RunModel(InferenceSession session, long[] tokenIds, float[] style, float speed, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        var inputTensor = new DenseTensor<long>(new[] { 1, tokenIds.Length });
        tokenIds.CopyTo(inputTensor.Buffer.Span);

        var styleTensor = new DenseTensor<float>(new[] { 1, StyleDim });
        style.CopyTo(styleTensor.Buffer.Span);

        var speedTensor = new DenseTensor<float>(new[] { 1 });
        speedTensor[0] = speed;

        var inputs = new List<NamedOnnxValue>(3)
        {
            NamedOnnxValue.CreateFromTensor("input_ids", inputTensor),
            NamedOnnxValue.CreateFromTensor("style", styleTensor),
            NamedOnnxValue.CreateFromTensor("speed", speedTensor)
        };

        using var results = session.Run(inputs);
        float[]? waveform = null;
        foreach (var result in results)
        {
            if (result.Name.Equals("waveform", StringComparison.OrdinalIgnoreCase))
            {
                waveform = TensorTo1DFloat(result.AsTensor<float>());
                break;
            }

            if (waveform is null && result.Value is Tensor<float> tf)
            {
                waveform = TensorTo1DFloat(tf);
            }
        }

        if (waveform is null || waveform.Length == 0)
        {
            throw new InvalidOperationException("Kitten ONNX inference returned no waveform.");
        }

        // Match the reference wrapper which trims trailing decoder tail.
        if (waveform.Length > 6000)
        {
            Array.Resize(ref waveform, waveform.Length - 5000);
        }

        return waveform;
    }

    private static float[] TensorTo1DFloat(Tensor<float> tensor)
    {
        if (tensor.Length == 0)
        {
            return Array.Empty<float>();
        }

        var data = tensor.ToArray();
        return data;
    }

    private static long[] BuildTokenIds(string text)
    {
        var normalized = EnsureEndingPunctuation(text);
        var phonemeText = PhonemizeWithEspeakRequired(normalized);
        var tokenized = NormalizePhonemeTokenStream(phonemeText);
        if (string.IsNullOrWhiteSpace(tokenized))
        {
            throw new InvalidOperationException("Kitten phonemizer returned no tokens. Check eSpeak-NG installation.");
        }

        var core = KittenTextCleaner.Encode(tokenized);
        // Match Kitten source pipeline: [0] + tokens + [10] + [0]
        var ids = new long[core.Count + 3];
        ids[0] = 0;
        for (var i = 0; i < core.Count; i++)
        {
            ids[i + 1] = core[i];
        }
        ids[^2] = 10;
        ids[^1] = 0;
        return ids;
    }

    private static string NormalizeInputText(string text)
    {
        var s = (text ?? string.Empty)
            .Normalize(NormalizationForm.FormC)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace('\u201C', '"')
            .Replace('\u201D', '"')
            .Replace('\u2018', '\'')
            .Replace('\u2019', '\'')
            .Replace('\u2014', '-')
            .Replace('\u2013', '-')
            .Replace('\u00A0', ' ');

        s = ExpandCommonContractions(s);
        s = LeadingDecimalRegex.Replace(s, "0.$1");
        s = MultiWhitespaceRegex.Replace(s, " ").Trim();
        return s;
    }

    private static string ExpandCommonContractions(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var s = text;
        s = Regex.Replace(s, @"\bcan't\b", "cannot", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\bwon't\b", "will not", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\bshan't\b", "shall not", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\blet's\b", "let us", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\bit's\b", "it is", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)n't\b", "$1 not", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)'re\b", "$1 are", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)'ve\b", "$1 have", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)'ll\b", "$1 will", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)'d\b", "$1 would", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\b(\w+)'m\b", "$1 am", RegexOptions.IgnoreCase);
        return s;
    }

    private static string EnsureEndingPunctuation(string text)
    {
        var s = (text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(s))
        {
            return string.Empty;
        }

        var last = s[^1];
        return last is '.' or '!' or '?' or ',' or ';' or ':'
            ? s
            : s + ",";
    }

    private static string BasicEnglishTokenize(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var tokens = BasicTokenRegex.Matches(text).Select(m => m.Value).ToArray();
        return string.Join(" ", tokens);
    }

    // Preserve IPA/stress sequences from eSpeak output. We only normalize whitespace and punctuation spacing,
    // instead of re-tokenizing with a generic regex (which can split IPA combining/stress patterns unnaturally).
    private static string NormalizePhonemeTokenStream(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var s = text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
        s = MultiWhitespaceRegex.Replace(s, " ").Trim();
        if (string.IsNullOrWhiteSpace(s))
        {
            return string.Empty;
        }

        // Normalize spacing around punctuation tokens we intentionally preserve from the phonemizer path.
        var sb = new StringBuilder(s.Length + 16);
        for (var i = 0; i < s.Length; i++)
        {
            var ch = s[i];
            if (IsKittenPreservedPunctuation(ch))
            {
                if (sb.Length > 0 && sb[^1] != ' ')
                {
                    sb.Append(' ');
                }
                sb.Append(ch);
                if (i < s.Length - 1 && s[i + 1] != ' ')
                {
                    sb.Append(' ');
                }
                continue;
            }

            sb.Append(ch);
        }

        return MultiWhitespaceRegex.Replace(sb.ToString(), " ").Trim();
    }

    private static string PhonemizeWithEspeakRequired(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        // Kitten source uses an eSpeak-backed phonemizer with punctuation + stress preservation.
        // CLI espeak-ng IPA output can drop punctuation, so we phonemize text segments and re-insert
        // punctuation tokens to stay closer to phonemizer(... preserve_punctuation=True).
        var sb = new StringBuilder();
        var segment = new StringBuilder();
        void FlushSegment()
        {
            var seg = segment.ToString().Trim();
            segment.Clear();
            if (string.IsNullOrWhiteSpace(seg))
            {
                return;
            }

            var phon = PhonemizeEspeakSegment(seg);
            if (string.IsNullOrWhiteSpace(phon))
            {
                return;
            }

            if (sb.Length > 0 && !char.IsWhiteSpace(sb[^1]))
            {
                sb.Append(' ');
            }
            sb.Append(phon.Trim());
        }

        foreach (var ch in text)
        {
            if (IsKittenPreservedPunctuation(ch))
            {
                FlushSegment();
                if (sb.Length > 0 && !char.IsWhiteSpace(sb[^1]))
                {
                    sb.Append(' ');
                }
                sb.Append(ch);
                sb.Append(' ');
                continue;
            }

            if (char.IsWhiteSpace(ch))
            {
                FlushSegment();
                if (sb.Length > 0 && !char.IsWhiteSpace(sb[^1]))
                {
                    sb.Append(' ');
                }
                continue;
            }

            segment.Append(ch);
        }

        FlushSegment();
        var outText = MultiWhitespaceRegex.Replace(sb.ToString(), " ").Trim();
        if (string.IsNullOrWhiteSpace(outText))
        {
            throw new InvalidOperationException("eSpeak-NG returned empty phoneme output.");
        }

        return outText;
    }

    private static bool IsKittenPreservedPunctuation(char ch)
        => ch is ',' or '.' or '!' or '?' or ';' or ':' or '\u2026';

    private static string PhonemizeEspeakSegment(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var exe = ResolveEspeakNgExecutablePath();
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            psi.ArgumentList.Add("-q");
            psi.ArgumentList.Add("--ipa");
            psi.ArgumentList.Add("-v");
            psi.ArgumentList.Add("en-us");
            psi.ArgumentList.Add(text);

            using var process = Process.Start(psi);
            if (process is null)
            {
                throw new InvalidOperationException("Failed to start eSpeak-NG.");
            }

            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit(5000);
            if (!process.HasExited)
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                throw new InvalidOperationException("eSpeak-NG timed out while phonemizing text.");
            }

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    string.IsNullOrWhiteSpace(stderr)
                        ? $"eSpeak-NG exited with code {process.ExitCode}."
                        : $"eSpeak-NG error: {stderr.Trim()}");
            }

            return (stdout ?? string.Empty).Trim();
        }
        catch (Win32Exception)
        {
            throw new InvalidOperationException(
                "Kitten TTS requires eSpeak-NG for phonemization. Install eSpeak-NG and restart the app.");
        }
    }

    private static string ResolveEspeakNgExecutablePath()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "tools", "espeak-ng", "espeak-ng.exe"),
            Path.Combine(AppContext.BaseDirectory, "tools", "espeak-ng", "command_line", "espeak-ng.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "eSpeak NG", "espeak-ng.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "eSpeak NG", "espeak-ng.exe"),
            "espeak-ng"
        };

        foreach (var candidate in candidates)
        {
            if (string.Equals(candidate, "espeak-ng", StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }

            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
            {
                return candidate;
            }
        }

        return "espeak-ng";
    }

    private static string ResolveVoiceKey(string rawSelection, KittenModelConfig config)
    {
        var value = (rawSelection ?? string.Empty).Trim();
        if (value.StartsWith("kitten://", StringComparison.OrdinalIgnoreCase))
        {
            value = value["kitten://".Length..];
        }
        else if (value.StartsWith("kitten:", StringComparison.OrdinalIgnoreCase))
        {
            value = value["kitten:".Length..];
        }

        if (config.Aliases.TryGetValue(value, out var aliasTarget) && !string.IsNullOrWhiteSpace(aliasTarget))
        {
            return aliasTarget;
        }

        return string.IsNullOrWhiteSpace(value) ? "expr-voice-5-m" : value;
    }

    private static SessionOptions CreateSessionOptions(string? preferDevice)
    {
        var device = NormalizePreferDevice(preferDevice);
        var options = new SessionOptions
        {
            GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_EXTENDED
        };

        if (device == "cpu")
        {
            return options;
        }

        TryBootstrapCudaRuntimePaths();
        Exception? cudaError = null;
        try
        {
            options.AppendExecutionProvider_CUDA(0);
            return options;
        }
        catch (Exception ex)
        {
            cudaError = ex;
        }

        if (OperatingSystem.IsWindows())
        {
            try
            {
                options.AppendExecutionProvider_DML(0);
                return options;
            }
            catch
            {
                // fall through
            }
        }

        if (device == "gpu")
        {
            throw new InvalidOperationException(
                $"GPU was selected but CUDA/DML execution providers are unavailable. CUDA load error: {cudaError?.Message}");
        }

        return options;
    }

    private static string NormalizePreferDevice(string? value)
        => (value ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "gpu" => "gpu",
            "cuda" => "gpu",
            "cpu" => "cpu",
            _ => "auto"
        };

    private static void TryBootstrapCudaRuntimePaths()
    {
        lock (CudaBootstrapSync)
        {
            if (_cudaBootstrapDone)
            {
                return;
            }

            _cudaBootstrapDone = true;

            var candidates = new List<string>();
            static void AddCandidate(ICollection<string> list, string? dir)
            {
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    list.Add(dir);
                }
            }

            var cudaPath = Environment.GetEnvironmentVariable("CUDA_PATH");
            if (!string.IsNullOrWhiteSpace(cudaPath))
            {
                candidates.Add(Path.Combine(cudaPath, "bin"));
            }

            var appRoot = RuntimePathResolver.AppRoot;
            AddCandidate(candidates, Path.Combine(appRoot, "python_qwen", "Lib", "site-packages", "torch", "lib"));
            AddCandidate(candidates, Path.Combine(appRoot, "tools", "python_qwen", "Lib", "site-packages", "torch", "lib"));
            AddCandidate(candidates, Path.Combine(appRoot, ".venv", "Lib", "site-packages", "torch", "lib"));

            // Stability-first default: do not scan arbitrary system Python torch DLLs.
            // Opt in only if explicitly requested.
            var allowSystemTorch = string.Equals(
                Environment.GetEnvironmentVariable("AUDIOBOOK_ALLOW_SYSTEM_TORCH_CUDA"),
                "1",
                StringComparison.OrdinalIgnoreCase);
            if (allowSystemTorch)
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var pythonRoot = Path.Combine(localAppData, "Programs", "Python");
                if (Directory.Exists(pythonRoot))
                {
                    foreach (var pyDir in Directory.GetDirectories(pythonRoot))
                    {
                        AddCandidate(candidates, Path.Combine(pyDir, "Lib", "site-packages", "torch", "lib"));
                    }
                }
            }

            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var cudaToolkitRoot = Path.Combine(programFiles, "NVIDIA GPU Computing Toolkit", "CUDA");
            if (Directory.Exists(cudaToolkitRoot))
            {
                foreach (var verDir in Directory.GetDirectories(cudaToolkitRoot, "v12.*"))
                {
                    candidates.Add(Path.Combine(verDir, "bin"));
                }
            }

            var existingPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            var existingParts = new HashSet<string>(
                existingPath.Split(';', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);

            var additions = candidates
                .Where(Directory.Exists)
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(dir => File.Exists(Path.Combine(dir, "cublasLt64_12.dll")))
                .Where(dir => !existingParts.Contains(dir))
                .ToList();

            if (additions.Count == 0)
            {
                return;
            }

            Environment.SetEnvironmentVariable("PATH", string.Join(";", additions) + ";" + existingPath);
        }
    }

    private static void WriteWav16Mono(string path, float[] samples, int sampleRate)
    {
        using var fs = File.Create(path);
        using var bw = new BinaryWriter(fs);

        const short channels = 1;
        const short bitsPerSample = 16;
        var byteRate = sampleRate * channels * bitsPerSample / 8;
        var blockAlign = (short)(channels * bitsPerSample / 8);
        var dataSize = samples.Length * blockAlign;
        var riffSize = 36 + dataSize;

        bw.Write(Encoding.ASCII.GetBytes("RIFF"));
        bw.Write(riffSize);
        bw.Write(Encoding.ASCII.GetBytes("WAVE"));
        bw.Write(Encoding.ASCII.GetBytes("fmt "));
        bw.Write(16);
        bw.Write((short)1);
        bw.Write(channels);
        bw.Write(sampleRate);
        bw.Write(byteRate);
        bw.Write(blockAlign);
        bw.Write(bitsPerSample);
        bw.Write(Encoding.ASCII.GetBytes("data"));
        bw.Write(dataSize);

        foreach (var sample in samples)
        {
            var clamped = Math.Clamp(sample, -1.0f, 1.0f);
            bw.Write((short)Math.Round(clamped * 32767.0f));
        }
    }

    private sealed record KittenModelConfig(
        Dictionary<string, string> Aliases,
        Dictionary<string, float> SpeedPriors);

    /// <summary>
    /// Minimal English grapheme-to-IPA converter used when espeak-ng is not installed.
    /// Covers ~380 common/irregular words via lexicon and applies letter-to-sound rules
    /// for everything else, producing espeak-ng-compatible IPA for en-us.
    /// </summary>
    private static class BuiltInG2p
    {
        private static readonly Dictionary<string, string> Lex = new(StringComparer.OrdinalIgnoreCase)
        {
            // Articles / determiners
            ["the"] = "ðə", ["a"] = "ə", ["an"] = "æn",
            // Conjunctions
            ["and"] = "ænd", ["or"] = "ɔːɹ", ["but"] = "bʌt", ["nor"] = "nɔːɹ",
            ["yet"] = "jɛt", ["so"] = "soʊ", ["although"] = "ɔːlˈðoʊ",
            ["though"] = "ðoʊ", ["because"] = "bɪˈkɒz", ["while"] = "waɪl",
            ["unless"] = "ʌnˈlɛs", ["until"] = "ʌnˈtɪl", ["since"] = "sɪns",
            ["whether"] = "ˈwɛðɚ",
            // Prepositions
            ["of"] = "əv", ["in"] = "ɪn", ["on"] = "ɒn", ["at"] = "æt",
            ["to"] = "tə", ["for"] = "fɔːɹ", ["with"] = "wɪð", ["by"] = "baɪ",
            ["from"] = "fɹɒm", ["up"] = "ʌp", ["into"] = "ˈɪntə", ["as"] = "æz",
            ["through"] = "θɹuː", ["about"] = "əˈbaʊt", ["between"] = "bɪˈtwiːn",
            ["under"] = "ˈʌndɚ", ["above"] = "əˈbʌv", ["over"] = "ˈoʊvɚ",
            ["after"] = "ˈæftɚ", ["before"] = "bɪˈfɔːɹ", ["behind"] = "bɪˈhaɪnd",
            ["beside"] = "bɪˈsaɪd", ["beyond"] = "bɪˈɒnd", ["among"] = "əˈmʌŋ",
            ["against"] = "əˈɡɛnst", ["during"] = "ˈdjʊɹɪŋ", ["without"] = "wɪˈðaʊt",
            ["within"] = "wɪˈðɪn", ["upon"] = "əˈpɒn", ["toward"] = "tɔːɹd",
            ["towards"] = "tɔːɹdz",
            // Pronouns
            ["i"] = "aɪ", ["me"] = "miː", ["my"] = "maɪ", ["mine"] = "maɪn",
            ["myself"] = "maɪˈsɛlf", ["we"] = "wiː", ["our"] = "aʊɚ",
            ["us"] = "ʌs", ["ours"] = "aʊɚz", ["you"] = "juː", ["your"] = "jɔːɹ",
            ["yours"] = "jɔːɹz", ["he"] = "hiː", ["him"] = "hɪm", ["his"] = "hɪz",
            ["she"] = "ʃiː", ["her"] = "hɜːɹ", ["hers"] = "hɜːɹz",
            ["it"] = "ɪt", ["its"] = "ɪts", ["itself"] = "ɪtˈsɛlf",
            ["they"] = "ðeɪ", ["them"] = "ðɛm", ["their"] = "ðɛɹ", ["theirs"] = "ðɛɹz",
            ["this"] = "ðɪs", ["that"] = "ðæt", ["these"] = "ðiːz", ["those"] = "ðoʊz",
            ["who"] = "huː", ["whom"] = "huːm", ["whose"] = "huːz",
            ["what"] = "wɒt", ["which"] = "wɪtʃ", ["when"] = "wɛn",
            ["where"] = "wɛɹ", ["how"] = "haʊ", ["why"] = "waɪ",
            ["each"] = "iːtʃ", ["every"] = "ˈɛvɹiː", ["any"] = "ˈɛniː",
            ["some"] = "sʌm", ["all"] = "ɔːl", ["none"] = "nʌn", ["both"] = "boʊθ",
            ["either"] = "ˈiːðɚ", ["neither"] = "ˈniːðɚ",
            ["anyone"] = "ˈɛnɪwʌn", ["someone"] = "ˈsʌmwʌn", ["everyone"] = "ˈɛvɹɪwʌn",
            ["anything"] = "ˈɛnɪθɪŋ", ["something"] = "ˈsʌmθɪŋ",
            ["everything"] = "ˈɛvɹɪθɪŋ", ["nothing"] = "ˈnʌθɪŋ",
            ["nowhere"] = "ˈnoʊwɛɹ", ["somewhere"] = "ˈsʌmwɛɹ",
            ["everywhere"] = "ˈɛvɹɪwɛɹ", ["another"] = "əˈnʌðɚ",
            // Auxiliaries and copula
            ["is"] = "ɪz", ["are"] = "ɑːɹ", ["was"] = "wɒz", ["were"] = "wɜːɹ",
            ["be"] = "biː", ["been"] = "biːn", ["being"] = "ˈbiːɪŋ",
            ["have"] = "hæv", ["has"] = "hæz", ["had"] = "hæd", ["having"] = "ˈhævɪŋ",
            ["do"] = "duː", ["does"] = "dʌz", ["did"] = "dɪd", ["done"] = "dʌn",
            ["doing"] = "ˈduːɪŋ", ["will"] = "wɪl", ["would"] = "wʊd",
            ["shall"] = "ʃæl", ["should"] = "ʃʊd", ["may"] = "meɪ",
            ["might"] = "maɪt", ["can"] = "kæn", ["could"] = "kʊd", ["must"] = "mʌst",
            // Contractions
            ["don't"] = "doʊnt", ["doesn't"] = "ˈdʌzənt", ["didn't"] = "ˈdɪdənt",
            ["won't"] = "woʊnt", ["wouldn't"] = "ˈwʊdənt", ["couldn't"] = "ˈkʊdənt",
            ["shouldn't"] = "ˈʃʊdənt", ["can't"] = "kænt", ["isn't"] = "ˈɪzənt",
            ["aren't"] = "ɑːɹnt", ["wasn't"] = "ˈwɒzənt", ["weren't"] = "ˈwɜːɹənt",
            ["haven't"] = "ˈhævənt", ["hasn't"] = "ˈhæzənt", ["hadn't"] = "ˈhædənt",
            ["i'm"] = "aɪm", ["i've"] = "aɪv", ["i'll"] = "aɪl", ["i'd"] = "aɪd",
            ["he's"] = "hiːz", ["she's"] = "ʃiːz", ["it's"] = "ɪts",
            ["we're"] = "wɪɹ", ["we've"] = "wiːv", ["we'll"] = "wiːl",
            ["you're"] = "jɔːɹ", ["you've"] = "juːv", ["you'll"] = "juːl",
            ["they're"] = "ðɛɹ", ["they've"] = "ðeɪv", ["they'll"] = "ðeɪl",
            ["could've"] = "ˈkʊdəv", ["would've"] = "ˈwʊdəv", ["should've"] = "ˈʃʊdəv",
            ["that's"] = "ðæts", ["there's"] = "ðɛɹz", ["here's"] = "hɪɹz",
            // Common irregular verbs
            ["go"] = "ɡoʊ", ["goes"] = "ɡoʊz", ["went"] = "wɛnt", ["gone"] = "ɡɒn",
            ["going"] = "ˈɡoʊɪŋ",
            ["get"] = "ɡɛt", ["gets"] = "ɡɛts", ["got"] = "ɡɒt", ["getting"] = "ˈɡɛtɪŋ",
            ["come"] = "kʌm", ["comes"] = "kʌmz", ["came"] = "keɪm", ["coming"] = "ˈkʌmɪŋ",
            ["take"] = "teɪk", ["took"] = "tʊk", ["taken"] = "ˈteɪkən", ["taking"] = "ˈteɪkɪŋ",
            ["give"] = "ɡɪv", ["gave"] = "ɡeɪv", ["given"] = "ˈɡɪvən", ["giving"] = "ˈɡɪvɪŋ",
            ["know"] = "noʊ", ["knew"] = "njuː", ["known"] = "noʊn", ["knowing"] = "ˈnoʊɪŋ",
            ["think"] = "θɪŋk", ["thinks"] = "θɪŋks", ["thought"] = "θɔːt",
            ["thinking"] = "ˈθɪŋkɪŋ",
            ["see"] = "siː", ["saw"] = "sɔː", ["seen"] = "siːn", ["seeing"] = "ˈsiːɪŋ",
            ["say"] = "seɪ", ["says"] = "sɛz", ["said"] = "sɛd", ["saying"] = "ˈseɪɪŋ",
            ["tell"] = "tɛl", ["told"] = "toʊld",
            ["make"] = "meɪk", ["made"] = "meɪd", ["making"] = "ˈmeɪkɪŋ",
            ["put"] = "pʊt", ["putting"] = "ˈpʊtɪŋ",
            ["keep"] = "kiːp", ["kept"] = "kɛpt",
            ["let"] = "lɛt", ["letting"] = "ˈlɛtɪŋ",
            ["run"] = "ɹʌn", ["ran"] = "ɹæn", ["running"] = "ˈɹʌnɪŋ",
            ["set"] = "sɛt", ["setting"] = "ˈsɛtɪŋ",
            ["turn"] = "tɜːɹn", ["turned"] = "tɜːɹnd", ["turning"] = "ˈtɜːɹnɪŋ",
            ["ask"] = "æsk", ["asked"] = "æskt",
            ["seem"] = "siːm", ["seemed"] = "siːmd",
            ["look"] = "lʊk", ["looked"] = "lʊkt", ["looking"] = "ˈlʊkɪŋ",
            ["use"] = "juːz", ["used"] = "juːzd", ["using"] = "ˈjuːzɪŋ",
            ["want"] = "wɒnt", ["wanted"] = "ˈwɒntɪd", ["wanting"] = "ˈwɒntɪŋ",
            ["feel"] = "fiːl", ["felt"] = "fɛlt",
            ["begin"] = "bɪˈɡɪn", ["began"] = "bɪˈɡæn", ["begun"] = "bɪˈɡʌn",
            ["show"] = "ʃoʊ", ["showed"] = "ʃoʊd", ["shown"] = "ʃoʊn",
            ["move"] = "muːv", ["moved"] = "muːvd", ["moving"] = "ˈmuːvɪŋ",
            ["live"] = "lɪv", ["lived"] = "lɪvd", ["living"] = "ˈlɪvɪŋ",
            ["die"] = "daɪ", ["died"] = "daɪd", ["dying"] = "ˈdaɪɪŋ",
            ["stand"] = "stænd", ["stood"] = "stʊd",
            ["fall"] = "fɔːl", ["fell"] = "fɛl", ["fallen"] = "ˈfɔːlən",
            ["hold"] = "hoʊld", ["held"] = "hɛld",
            ["bring"] = "bɹɪŋ", ["brought"] = "bɹɔːt",
            ["write"] = "ɹaɪt", ["wrote"] = "ɹoʊt", ["written"] = "ˈɹɪtən",
            ["read"] = "ɹiːd",
            ["speak"] = "spiːk", ["spoke"] = "spoʊk", ["spoken"] = "ˈspoʊkən",
            ["hear"] = "hɪɹ", ["heard"] = "hɜːɹd",
            ["meet"] = "miːt", ["met"] = "mɛt",
            ["eat"] = "iːt", ["ate"] = "eɪt", ["eaten"] = "ˈiːtən",
            ["sit"] = "sɪt", ["sat"] = "sæt", ["sitting"] = "ˈsɪtɪŋ",
            ["leave"] = "liːv", ["left"] = "lɛft", ["leaving"] = "ˈliːvɪŋ",
            ["find"] = "faɪnd", ["found"] = "faʊnd",
            ["lose"] = "luːz", ["lost"] = "lɒst",
            ["choose"] = "tʃuːz", ["chose"] = "tʃoʊz", ["chosen"] = "ˈtʃoʊzən",
            ["break"] = "bɹeɪk", ["broke"] = "bɹoʊk", ["broken"] = "ˈbɹoʊkən",
            ["open"] = "ˈoʊpən", ["opened"] = "ˈoʊpənd", ["opening"] = "ˈoʊpənɪŋ",
            ["close"] = "kloʊz", ["closed"] = "kloʊzd",
            ["build"] = "bɪld", ["built"] = "bɪlt",
            ["buy"] = "baɪ", ["bought"] = "bɔːt",
            ["fight"] = "faɪt", ["fought"] = "fɔːt",
            ["teach"] = "tiːtʃ", ["taught"] = "tɔːt",
            ["catch"] = "kætʃ", ["caught"] = "kɔːt",
            ["draw"] = "dɹɔː", ["drew"] = "dɹuː", ["drawn"] = "dɹɔːn",
            ["grow"] = "ɡɹoʊ", ["grew"] = "ɡɹuː", ["grown"] = "ɡɹoʊn",
            ["throw"] = "θɹoʊ", ["threw"] = "θɹuː", ["thrown"] = "θɹoʊn",
            ["fly"] = "flaɪ", ["flew"] = "fluː", ["flown"] = "floʊn",
            ["rise"] = "ɹaɪz", ["rose"] = "ɹoʊz", ["risen"] = "ˈɹɪzən",
            ["drive"] = "dɹaɪv", ["drove"] = "dɹoʊv", ["driven"] = "ˈdɹɪvən",
            ["ride"] = "ɹaɪd", ["rode"] = "ɹoʊd", ["ridden"] = "ˈɹɪdən",
            ["bite"] = "baɪt", ["bit"] = "bɪt", ["bitten"] = "ˈbɪtən",
            ["hide"] = "haɪd", ["hid"] = "hɪd", ["hidden"] = "ˈhɪdən",
            ["wake"] = "weɪk", ["woke"] = "woʊk",
            ["wear"] = "wɛɹ", ["wore"] = "wɔːɹ", ["worn"] = "wɔːɹn",
            ["pay"] = "peɪ", ["paid"] = "peɪd",
            ["lay"] = "leɪ", ["laid"] = "leɪd",
            ["sing"] = "sɪŋ", ["sang"] = "sæŋ", ["sung"] = "sʌŋ",
            ["ring"] = "ɹɪŋ", ["rang"] = "ɹæŋ", ["rung"] = "ɹʌŋ",
            ["spring"] = "spɹɪŋ", ["sprang"] = "spɹæŋ",
            ["swim"] = "swɪm", ["swam"] = "swæm",
            ["win"] = "wɪn", ["won"] = "wʌn",
            ["hit"] = "hɪt", ["cut"] = "kʌt", ["shut"] = "ʃʌt",
            ["hurt"] = "hɜːɹt", ["cost"] = "kɒst",
            // Common nouns
            ["time"] = "taɪm", ["year"] = "jɪɹ", ["day"] = "deɪ",
            ["way"] = "weɪ", ["man"] = "mæn", ["men"] = "mɛn",
            ["woman"] = "ˈwʊmən", ["women"] = "ˈwɪmɪn",
            ["child"] = "tʃaɪld", ["children"] = "ˈtʃɪldɹən",
            ["world"] = "wɜːɹld", ["life"] = "laɪf", ["hand"] = "hænd",
            ["part"] = "pɑːɹt", ["place"] = "pleɪs", ["case"] = "keɪs",
            ["home"] = "hoʊm", ["word"] = "wɜːɹd", ["eye"] = "aɪ",
            ["house"] = "haʊs", ["night"] = "naɪt", ["friend"] = "fɹɛnd",
            ["love"] = "lʌv", ["head"] = "hɛd", ["heart"] = "hɑːɹt",
            ["body"] = "ˈbɒdiː", ["voice"] = "vɔɪs", ["door"] = "dɔːɹ",
            ["blood"] = "blʌd", ["face"] = "feɪs", ["book"] = "bʊk",
            ["room"] = "ɹuːm", ["money"] = "ˈmʌniː", ["water"] = "ˈwɔːtɚ",
            ["fire"] = "ˈfaɪɚ", ["air"] = "ɛɹ", ["ground"] = "ɡɹaʊnd",
            ["sun"] = "sʌn", ["moon"] = "muːn", ["sky"] = "skaɪ",
            ["light"] = "laɪt", ["dark"] = "dɑːɹk", ["people"] = "ˈpiːpəl",
            ["matter"] = "ˈmætɚ", ["name"] = "neɪm", ["side"] = "saɪd",
            ["end"] = "ɛnd", ["floor"] = "flɔːɹ", ["food"] = "fuːd",
            ["power"] = "ˈpaʊɚ", ["town"] = "taʊn", ["road"] = "ɹoʊd",
            ["city"] = "ˈsɪtiː", ["earth"] = "ɜːɹθ", ["death"] = "dɛθ",
            ["hour"] = "aʊɚ", ["brother"] = "ˈbɹʌðɚ", ["sister"] = "ˈsɪstɚ",
            ["mother"] = "ˈmʌðɚ", ["father"] = "ˈfɑːðɚ", ["daughter"] = "ˈdɔːtɚ",
            ["son"] = "sʌn", ["husband"] = "ˈhʌzbənd", ["wife"] = "waɪf",
            ["king"] = "kɪŋ", ["queen"] = "kwiːn", ["lord"] = "lɔːɹd",
            ["knight"] = "naɪt", ["sword"] = "sɔːɹd",
            ["answer"] = "ˈænsɚ", ["question"] = "ˈkwɛstʃən",
            ["laughter"] = "ˈlæftɚ", ["laugh"] = "læf",
            ["half"] = "hæf", ["calf"] = "kæf",
            ["back"] = "bæk", ["front"] = "fɹʌnt",
            // Adjectives
            ["new"] = "njuː", ["old"] = "oʊld", ["good"] = "ɡʊd",
            ["great"] = "ɡɹeɪt", ["little"] = "ˈlɪtəl", ["own"] = "oʊn",
            ["other"] = "ˈʌðɚ", ["right"] = "ɹaɪt", ["large"] = "lɑːɹdʒ",
            ["big"] = "bɪɡ", ["small"] = "smɔːl", ["long"] = "lɒŋ",
            ["high"] = "haɪ", ["low"] = "loʊ", ["next"] = "nɛkst",
            ["early"] = "ˈɜːɹliː", ["young"] = "jʌŋ", ["real"] = "ɹɪəl",
            ["only"] = "ˈoʊnliː", ["same"] = "seɪm", ["last"] = "læst",
            ["first"] = "fɜːɹst", ["much"] = "mʌtʃ", ["more"] = "mɔːɹ",
            ["most"] = "moʊst", ["such"] = "sʌtʃ", ["few"] = "fjuː",
            ["free"] = "fɹiː", ["full"] = "fʊl", ["sure"] = "ʃɔːɹ",
            ["true"] = "tɹuː", ["whole"] = "hoʊl", ["wide"] = "waɪd",
            ["deep"] = "diːp", ["white"] = "waɪt", ["black"] = "blæk",
            ["red"] = "ɹɛd", ["blue"] = "bluː", ["green"] = "ɡɹiːn",
            ["cold"] = "koʊld", ["hot"] = "hɒt", ["hard"] = "hɑːɹd",
            ["soft"] = "sɒft", ["near"] = "nɪɹ", ["far"] = "fɑːɹ",
            ["dead"] = "dɛd", ["ready"] = "ˈɹɛdiː", ["bright"] = "bɹaɪt",
            ["heavy"] = "ˈhɛviː", ["strange"] = "stɹeɪndʒ",
            ["beautiful"] = "ˈbjuːtɪfəl", ["important"] = "ɪmˈpɔːɹtənt",
            ["strong"] = "stɹɒŋ", ["weak"] = "wiːk", ["tall"] = "tɔːl",
            ["short"] = "ʃɔːɹt", ["thin"] = "θɪn", ["thick"] = "θɪk",
            ["quick"] = "kwɪk", ["slow"] = "sloʊ", ["quiet"] = "ˈkwaɪɪt",
            ["loud"] = "laʊd", ["happy"] = "ˈhæpiː", ["sad"] = "sæd",
            ["afraid"] = "əˈfɹeɪd", ["angry"] = "ˈæŋɡɹiː",
            ["alone"] = "əˈloʊn", ["together"] = "təˈɡɛðɚ",
            // Adverbs
            ["not"] = "nɒt", ["also"] = "ˈɔːlsoʊ", ["very"] = "ˈvɛɹiː",
            ["now"] = "naʊ", ["still"] = "stɪl", ["just"] = "dʒʌst",
            ["even"] = "ˈiːvən", ["well"] = "wɛl", ["then"] = "ðɛn",
            ["here"] = "hɪɹ", ["there"] = "ðɛɹ", ["out"] = "aʊt",
            ["again"] = "əˈɡɛn", ["already"] = "ɔːlˈɹɛdiː",
            ["never"] = "ˈnɛvɚ", ["always"] = "ˈɔːlweɪz", ["often"] = "ˈɔːfən",
            ["ever"] = "ˈɛvɚ", ["once"] = "wʌns", ["twice"] = "twaɪs",
            ["perhaps"] = "pɚˈhæps", ["maybe"] = "ˈmeɪbiː",
            ["almost"] = "ˈɔːlmoʊst", ["quite"] = "kwaɪt",
            ["rather"] = "ˈɹæðɚ", ["soon"] = "suːn", ["away"] = "əˈweɪ",
            ["around"] = "əˈɹaʊnd", ["across"] = "əˈkɹɒs", ["ahead"] = "əˈhɛd",
            ["below"] = "bɪˈloʊ", ["down"] = "daʊn", ["off"] = "ɒf",
            // Numbers
            ["one"] = "wʌn", ["two"] = "tuː", ["three"] = "θɹiː",
            ["four"] = "fɔːɹ", ["five"] = "faɪv", ["six"] = "sɪks",
            ["seven"] = "ˈsɛvən", ["eight"] = "eɪt", ["nine"] = "naɪn",
            ["ten"] = "tɛn", ["eleven"] = "ɪˈlɛvən", ["twelve"] = "twɛlv",
            ["thirteen"] = "θɜːɹˈtiːn", ["fourteen"] = "fɔːɹˈtiːn",
            ["fifteen"] = "fɪfˈtiːn", ["sixteen"] = "sɪksˈtiːn",
            ["seventeen"] = "sɛvənˈtiːn", ["eighteen"] = "eɪˈtiːn",
            ["nineteen"] = "naɪnˈtiːn", ["twenty"] = "ˈtwɛntiː",
            ["hundred"] = "ˈhʌndɹəd", ["thousand"] = "ˈθaʊzənd",
            ["million"] = "ˈmɪljən",
            // Irregular spellings / high-frequency fiction words
            ["no"] = "noʊ", ["yes"] = "jɛs",
            ["please"] = "pliːz", ["thank"] = "θæŋk", ["thanks"] = "θæŋks",
            ["walk"] = "wɔːk", ["talk"] = "tɔːk", ["calm"] = "kɑːm",
            ["palm"] = "pɑːm", ["rough"] = "ɹʌf", ["tough"] = "tʌf",
            ["enough"] = "ɪˈnʌf", ["cough"] = "kɒf",
            // ea-words where rules give wrong vowel (iː vs ɛ)
            ["threat"] = "θɹɛt", ["threats"] = "θɹɛts", ["threatening"] = "ˈθɹɛtənɪŋ",
            ["breath"] = "bɹɛθ", ["breathe"] = "bɹiːð", ["breathing"] = "ˈbɹiːðɪŋ",
            ["health"] = "hɛlθ", ["healthy"] = "ˈhɛlθiː",
            ["wealth"] = "wɛlθ", ["wealthy"] = "ˈwɛlθiː",
            ["sweat"] = "swɛt", ["bread"] = "bɹɛd", ["spread"] = "spɹɛd",
            ["dread"] = "dɹɛd", ["thread"] = "θɹɛd", ["tread"] = "tɹɛd",
            ["instead"] = "ɪnˈstɛd", ["weapon"] = "ˈwɛpən", ["weapons"] = "ˈwɛpənz",
            ["heaven"] = "ˈhɛvən", ["breakfast"] = "ˈbɹɛkfəst",
            ["bear"] = "bɛɹ", ["bears"] = "bɛɹz", ["swear"] = "swɛɹ", ["pear"] = "pɛɹ",
            ["thing"] = "θɪŋ", ["things"] = "θɪŋz",
            ["knife"] = "naɪf", ["knee"] = "niː", ["kneel"] = "niːl",
            ["knock"] = "nɒk", ["wrap"] = "ɹæp", ["wrong"] = "ɹɒŋ",
            ["wrist"] = "ɹɪst", ["forward"] = "ˈfɔːɹwɚd",
            ["reward"] = "ɹɪˈwɔːɹd", ["indeed"] = "ɪnˈdiːd",
            ["moment"] = "ˈmoʊmənt", ["second"] = "ˈsɛkənd",
        };

        private static readonly string[] DigitWords =
            { "zɪɹoʊ", "wʌn", "tuː", "θɹiː", "fɔːɹ", "faɪv", "sɪks", "sɛvən", "eɪt", "naɪn" };

        public static string Convert(string text)
        {
            var sb = new StringBuilder(text.Length * 2);
            var i = 0;
            while (i < text.Length)
            {
                var c = text[i];
                if (char.IsLetter(c))
                {
                    var j = i;
                    while (j < text.Length && (char.IsLetter(text[j]) ||
                           (text[j] == '\'' && j > i && j + 1 < text.Length && char.IsLetter(text[j + 1]))))
                    {
                        j++;
                    }
                    sb.Append(WordToIpa(text.Substring(i, j - i)));
                    i = j;
                }
                else if (char.IsDigit(c))
                {
                    sb.Append(DigitWords[c - '0']);
                    i++;
                }
                else
                {
                    sb.Append(c);
                    i++;
                }
            }
            return sb.ToString();
        }

        private static string WordToIpa(string word)
        {
            if (string.IsNullOrEmpty(word))
            {
                return word;
            }

            var lower = word.ToLowerInvariant();
            if (Lex.TryGetValue(lower, out var ipa))
            {
                return ipa;
            }

            return ApplyRules(lower);
        }

        private static string ApplyRules(string w)
        {
            if (string.IsNullOrEmpty(w))
            {
                return w;
            }

            var sb = new StringBuilder(w.Length * 2);
            var len = w.Length;
            for (var i = 0; i < len;)
            {
                var c = w[i];
                var n1 = i + 1 < len ? w[i + 1] : '\0';
                var n2 = i + 2 < len ? w[i + 2] : '\0';
                var n3 = i + 3 < len ? w[i + 3] : '\0';
                var n4 = i + 4 < len ? w[i + 4] : '\0';
                var prev = i > 0 ? w[i - 1] : '\0';

                // 5-char patterns
                if (c == 'o' && n1 == 'u' && n2 == 'g' && n3 == 'h' && n4 == 't')
                { sb.Append("ɔːt"); i += 5; continue; }

                // 4-char patterns
                if (c == 't' && n1 == 'i' && n2 == 'o' && n3 == 'n')
                { sb.Append("ʃən"); i += 4; continue; }
                if (c == 't' && n1 == 'u' && n2 == 'r' && n3 == 'e')
                { sb.Append("tʃɚ"); i += 4; continue; }
                if (c == 'i' && n1 == 'g' && n2 == 'h' && n3 == 't')
                { sb.Append("aɪt"); i += 4; continue; }
                if (c == 's' && n1 == 'i' && n2 == 'o' && n3 == 'n' && IsVwl(prev))
                { sb.Append("ʒən"); i += 4; continue; }
                if (c == 'o' && n1 == 'u' && n2 == 'g' && n3 == 'h')
                { sb.Append("ʌf"); i += 4; continue; }  // rough, tough fallback

                // 3-char patterns
                if (c == 'i' && n1 == 'g' && n2 == 'h')
                { sb.Append("aɪ"); i += 3; continue; }
                if (c == 'd' && n1 == 'g' && n2 == 'e')
                { sb.Append("dʒ"); i += 3; continue; }
                if (c == 't' && n1 == 'c' && n2 == 'h')
                { sb.Append("tʃ"); i += 3; continue; }

                // 2-char consonant clusters
                if (c == 'c' && n1 == 'h') { sb.Append("tʃ"); i += 2; continue; }
                if (c == 's' && n1 == 'h') { sb.Append("ʃ"); i += 2; continue; }
                if (c == 't' && n1 == 'h') { sb.Append("θ"); i += 2; continue; }
                if (c == 'p' && n1 == 'h') { sb.Append("f"); i += 2; continue; }
                if (c == 'w' && n1 == 'h') { sb.Append("w"); i += 2; continue; }
                if (c == 'c' && n1 == 'k') { sb.Append("k"); i += 2; continue; }
                if (c == 'n' && n1 == 'g') { sb.Append("ŋ"); i += 2; continue; }
                if (c == 'n' && n1 == 'k') { sb.Append("ŋk"); i += 2; continue; }
                if (c == 'g' && n1 == 'h') { i += 2; continue; }  // silent gh
                if (c == 'k' && n1 == 'n') { sb.Append("n"); i += 2; continue; }  // knee, knife
                if (c == 'w' && n1 == 'r') { sb.Append("ɹ"); i += 2; continue; }  // write, wrap
                if (c == 'q' && n1 == 'u') { sb.Append("kw"); i += 2; continue; }

                // Vowel digraphs
                if (c == 'a' && n1 == 'i') { sb.Append("eɪ"); i += 2; continue; }
                if (c == 'a' && n1 == 'y') { sb.Append("eɪ"); i += 2; continue; }
                if (c == 'a' && n1 == 'u') { sb.Append("ɔː"); i += 2; continue; }
                if (c == 'a' && n1 == 'w') { sb.Append("ɔː"); i += 2; continue; }
                // ea before certain consonants (or end-of-word) → ɛ; otherwise → iː
                // Words like threat/breath/health/dread/spread use ɛ; beam/seal/read use iː
                if (c == 'e' && n1 == 'a')
                {
                    bool eaShort = (n2 == 'd' || n2 == 't' || n2 == 'l' || n2 == '\0') ||
                                   (n2 == 't' && n3 == 'h');
                    sb.Append(eaShort ? "ɛ" : "iː"); i += 2; continue;
                }
                if (c == 'e' && n1 == 'e') { sb.Append("iː"); i += 2; continue; }
                if (c == 'e' && n1 == 'w') { sb.Append("juː"); i += 2; continue; }
                if (c == 'i' && n1 == 'e') { sb.Append(i + 2 == len ? "iː" : "aɪ"); i += 2; continue; }
                if (c == 'o' && n1 == 'a') { sb.Append("oʊ"); i += 2; continue; }
                if (c == 'o' && n1 == 'o') { sb.Append(n2 == 'k' ? "ʊ" : "uː"); i += 2; continue; }
                if (c == 'o' && n1 == 'i') { sb.Append("ɔɪ"); i += 2; continue; }
                if (c == 'o' && n1 == 'y') { sb.Append("ɔɪ"); i += 2; continue; }
                if (c == 'o' && n1 == 'u') { sb.Append("aʊ"); i += 2; continue; }
                if (c == 'o' && n1 == 'w') { sb.Append("oʊ"); i += 2; continue; }
                if (c == 'u' && n1 == 'e') { sb.Append("juː"); i += 2; continue; }
                if (c == 'u' && n1 == 'i') { sb.Append("uː"); i += 2; continue; }

                // r-coloured vowels
                // ar+e at word end = magic-e pattern → ɛɹ (care, dare, share, aware)
                if (c == 'a' && n1 == 'r' && n2 == 'e' && i + 3 == len) { sb.Append("ɛɹ"); i += 3; continue; }
                if (c == 'a' && n1 == 'r') { sb.Append("ɑːɹ"); i += 2; continue; }
                if (c == 'e' && n1 == 'r') { sb.Append(i + 2 >= len ? "ɚ" : "ɜːɹ"); i += 2; continue; }
                if (c == 'i' && n1 == 'r') { sb.Append("ɜːɹ"); i += 2; continue; }
                if (c == 'o' && n1 == 'r') { sb.Append("ɔːɹ"); i += 2; continue; }
                if (c == 'u' && n1 == 'r') { sb.Append("ɜːɹ"); i += 2; continue; }

                // Magic-e: vowel + single consonant + 'e' at word end
                if (IsVwl(c) && n1 != '\0' && !IsVwl(n1) && n2 == 'e' && i + 3 == len)
                {
                    sb.Append(c switch
                    {
                        'a' => "eɪ", 'e' => "iː", 'i' => "aɪ",
                        'o' => "oʊ", 'u' => "juː", _ => c.ToString()
                    });
                    i++;
                    continue;
                }

                // Single characters
                switch (c)
                {
                    case 'b': sb.Append("b"); break;
                    case 'c': sb.Append(n1 == 'e' || n1 == 'i' || n1 == 'y' ? "s" : "k"); break;
                    case 'd': sb.Append("d"); break;
                    case 'e': sb.Append(i == len - 1 ? "" : "ɛ"); break;  // silent final e
                    case 'f': sb.Append("f"); break;
                    case 'g': sb.Append(n1 == 'e' || n1 == 'i' || n1 == 'y' ? "dʒ" : "ɡ"); break;
                    case 'h': sb.Append("h"); break;
                    case 'i': sb.Append("ɪ"); break;
                    case 'j': sb.Append("dʒ"); break;
                    case 'k': sb.Append("k"); break;
                    case 'l': sb.Append("l"); break;
                    case 'm': sb.Append("m"); break;
                    case 'n': sb.Append("n"); break;
                    case 'o': sb.Append("ɒ"); break;
                    case 'p': sb.Append("p"); break;
                    case 'r': sb.Append("ɹ"); break;
                    case 's': sb.Append(IsVwl(prev) && IsVwl(n1) ? "z" : "s"); break;
                    case 't': sb.Append("t"); break;
                    case 'u': sb.Append("ʌ"); break;
                    case 'v': sb.Append("v"); break;
                    case 'w': sb.Append("w"); break;
                    case 'x': sb.Append("ks"); break;
                    case 'y': sb.Append(i == 0 ? "j" : (i == len - 1 || IsVwl(n1)) ? "iː" : "ɪ"); break;
                    case 'z': sb.Append("z"); break;
                    case 'a': sb.Append("æ"); break;
                    default: sb.Append(c); break;
                }
                i++;
            }
            return sb.ToString();
        }

        private static bool IsVwl(char c)
            => c == 'a' || c == 'e' || c == 'i' || c == 'o' || c == 'u';
    }

    private static class KittenTextCleaner
    {
        private static readonly Dictionary<char, int> SymbolToId = BuildDictionary();
        private const string Punctuation =
            ";:,.!?\u00A1\u00BF\u2014\u2026\"\u00AB\u00BB\"\" ";
        private const string Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
        private const string LettersIpa =
            "\u0251\u0250\u0252\u00E6\u0253\u0299\u03B2\u0254\u0255\u00E7\u0257\u0256\u00F0\u02A4\u0259\u0258\u025A\u025B\u025C\u025D\u025E\u025F\u0284\u0261\u0260\u0262\u029B\u0266\u0267\u0127\u0265\u029C\u0268\u026A\u029D\u026D\u026C\u026B\u026E\u029F\u0271\u026F\u0270\u014B\u0273\u0272\u0274\u00F8\u0275\u0278\u03B8\u0153\u0276\u0298\u0279\u027A\u027E\u027B\u0280\u0281\u027D\u0282\u0283\u0288\u02A7\u0289\u028A\u028B\u2C71\u028C\u0263\u0264\u028D\u03C7\u028E\u028F\u0291\u0290\u0292\u0294\u02A1\u0295\u02A2\u01C0\u01C1\u01C2\u01C3\u02C8\u02CC\u02D0\u02D1\u02BC\u02B4\u02B0\u02B1\u02B2\u02B7\u02E0\u02E4\u02DE\u2193\u2191\u2192\u2197\u2198'\u0329'\u1D7B";

        public static List<long> Encode(string text)
        {
            var result = new List<long>(Math.Max(8, text.Length));
            foreach (var ch in text)
            {
                if (SymbolToId.TryGetValue(ch, out var id))
                {
                    result.Add(id);
                }
            }
            return result;
        }

        private static Dictionary<char, int> BuildDictionary()
        {
            var symbols = new List<char>(1 + Punctuation.Length + Letters.Length + LettersIpa.Length)
            {
                '$'
            };
            symbols.AddRange(Punctuation);
            symbols.AddRange(Letters);
            symbols.AddRange(LettersIpa);

            var map = new Dictionary<char, int>(symbols.Count);
            for (var i = 0; i < symbols.Count; i++)
            {
                if (!map.ContainsKey(symbols[i]))
                {
                    map[symbols[i]] = i;
                }
            }
            return map;
        }
    }
}
