using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using App.Core.Runtime;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace App.Inference;

public sealed class Qwen3OnnxSplitBackend : ITtsBackend, IDisposable
{
    private static readonly object CudaBootstrapSync = new();
    private static bool _cudaBootstrapDone;

    private readonly LocalInferenceOptions _options;
    private readonly Random _rng = new();

    private InferenceSession? _speakerEncoder;
    private InferenceSession? _speechDecoder;
    private InferenceSession? _textEmbedding;
    private InferenceSession? _codecEmbedding;
    private InferenceSession? _talkerDecode;
    private InferenceSession? _codePredictor;
    private readonly List<InferenceSession?> _codePredictorEmbeds = new();
    private bool _codePredictorHasKv;

    private QwenSplitConfig? _config;
    private string _loadedKey = string.Empty;

    public Qwen3OnnxSplitBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "qwen3-onnx-split-native";
    public string ActiveExecutionProvider { get; private set; } = "unloaded";

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        return Task.Run(() => SynthesizeInternal(request, ct), ct);
    }

    public void Dispose()
    {
        _speakerEncoder?.Dispose();
        _speechDecoder?.Dispose();
        _textEmbedding?.Dispose();
        _codecEmbedding?.Dispose();
        _talkerDecode?.Dispose();
        _codePredictor?.Dispose();
        foreach (var session in _codePredictorEmbeds)
        {
            session?.Dispose();
        }
        _codePredictorEmbeds.Clear();

        _speakerEncoder = null;
        _speechDecoder = null;
        _textEmbedding = null;
        _codecEmbedding = null;
        _talkerDecode = null;
        _codePredictor = null;
        _config = null;
        _loadedKey = string.Empty;
        _codePredictorHasKv = false;
        ActiveExecutionProvider = "unloaded";
    }

    private void SynthesizeInternal(TtsRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            throw new ArgumentException("Input text is empty.");
        }
        if (string.IsNullOrWhiteSpace(request.VoicePath) || !File.Exists(request.VoicePath))
        {
            throw new FileNotFoundException($"Voice file not found: {request.VoicePath}");
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
            ? "xkos/Qwen3-TTS-12Hz-1.7B-ONNX"
            : _options.ModelRepoId.Trim();
        var cacheDir = ModelCachePath.ResolveAbsolute(_options.ModelCacheDir, RuntimePathResolver.AppRoot);
        EnsureLoaded(cacheDir, repoId);

        if (_config is null ||
            _speakerEncoder is null ||
            _speechDecoder is null ||
            _textEmbedding is null ||
            _codecEmbedding is null ||
            _talkerDecode is null)
        {
            throw new InvalidOperationException("Qwen split ONNX backend failed to initialize.");
        }

        var (wav, sr) = ReadWav(request.VoicePath);
        if (sr != _config.SpeakerSampleRate)
        {
            wav = Resample(wav, sr, _config.SpeakerSampleRate);
        }

        var mel = ComputeMel(wav, _config);
        var melFlat = TransposeMelToInput(mel);
        var speakerEmbed = RunSpeakerEncoder(_speakerEncoder, melFlat);

        var tokenizer = CreateTokenizer(_config.VocabPath, _config.MergesPath, _config.TokenizerConfigPath);
        try
        {
            var prompt = BuildAssistantPrompt(request.Text);
            var inputIds = TokenizerEncode(tokenizer, prompt);
            var audio = GenerateAudioWithRetries(
                request,
                inputIds,
                speakerEmbed,
                _config,
                _textEmbedding,
                _codecEmbedding,
                _talkerDecode,
                _codePredictor,
                _codePredictorEmbeds,
                _codePredictorHasKv,
                _speechDecoder,
                ct);
            if (Math.Abs(request.Speed - 1.0f) > 0.01f)
            {
                audio = ApplySpeedResample(audio, Math.Clamp(request.Speed, 0.55f, 1.75f));
            }
            WriteWav16Mono(request.OutputPath, audio, _config.OutputSampleRate);
        }
        finally
        {
            FreeTokenizer(tokenizer);
        }
    }

    private float[] GenerateAudioWithRetries(
        TtsRequest request,
        long[] inputIds,
        float[] speakerEmbed,
        QwenSplitConfig cfg,
        InferenceSession textEmbedding,
        InferenceSession codecEmbedding,
        InferenceSession talkerDecode,
        InferenceSession? codePredictor,
        IReadOnlyList<InferenceSession?> codePredictorEmbeds,
        bool codePredictorHasKv,
        InferenceSession speechDecoder,
        CancellationToken ct)
    {
        // Use one decode config for all chunks so voice/level stay consistent (no quality-guard retry with MakeSafer).
        var decode = QwenDecodeOptions.From(cfg, request);
        var state = BuildInputState(inputIds, speakerEmbed, cfg, textEmbedding, codecEmbedding);
        var codes = GenerateCodes(state, cfg, talkerDecode, codePredictor, codecEmbedding, codePredictorEmbeds, codePredictorHasKv, decode, ct);
        if (codes.GetLength(0) == 0)
        {
            throw new InvalidOperationException("Qwen split ONNX generation produced no audio codes.");
        }
        var audio = DecodeAudioCodes(speechDecoder, codes);
        return audio;
    }

    private void EnsureLoaded(string cacheDir, string repoId)
    {
        var parts = repoId.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            throw new InvalidOperationException($"Invalid model repo id: {repoId}");
        }

        var repoRoot = Path.Combine(cacheDir, "hf-cache", $"models--{parts[0]}--{parts[1]}");
        var key = repoRoot.ToLowerInvariant();
        if (_loadedKey == key)
        {
            return;
        }

        Dispose();

        if (!Directory.Exists(repoRoot))
        {
            throw new DirectoryNotFoundException($"Qwen split ONNX repo not found: {repoRoot}");
        }

        var onnxSharedDir = Path.Combine(repoRoot, "onnx", "shared");
        var onnxVoiceCloneDir = Path.Combine(repoRoot, "onnx", "voice_clone");
        if (!Directory.Exists(onnxSharedDir) || !Directory.Exists(onnxVoiceCloneDir))
        {
            throw new InvalidOperationException(
                "Qwen split ONNX layout missing. Expected folders: onnx/shared and onnx/voice_clone.");
        }

        var cfgPath = Path.Combine(repoRoot, "voice_clone_config.json");
        if (!File.Exists(cfgPath))
        {
            throw new FileNotFoundException("voice_clone_config.json missing in selected Qwen split ONNX repo.", cfgPath);
        }

        var (vocabPath, mergesPath, tokenizerCfgPath) = ResolveTokenizerFiles(repoRoot, cacheDir);
        var dllSource = ResolveRustDllPath(repoRoot, cacheDir);
        var appDll = Path.Combine(RuntimePathResolver.AppRoot, "qwen3_tts_rust.dll");
        if (!File.Exists(appDll) || new FileInfo(appDll).Length != new FileInfo(dllSource).Length)
        {
            File.Copy(dllSource, appDll, true);
        }

        _config = LoadConfig(cfgPath);
        _config.OnnxSharedDir = onnxSharedDir;
        _config.OnnxVoiceCloneDir = onnxVoiceCloneDir;
        _config.VocabPath = vocabPath;
        _config.MergesPath = mergesPath;
        _config.TokenizerConfigPath = tokenizerCfgPath;

        ValidateRequiredFiles(_config);

        var sessionOptions = CreateSessionOptions(_options.PreferDevice, out var resolvedProvider);
        ActiveExecutionProvider = resolvedProvider;
        _speakerEncoder = CreateSessionOrThrow(Path.Combine(onnxSharedDir, "speaker_encoder.onnx"), sessionOptions);
        _speechDecoder = CreateSessionOrThrow(Path.Combine(onnxSharedDir, "speech_tokenizer_decoder.onnx"), sessionOptions);
        _textEmbedding = CreateSessionOrThrow(Path.Combine(onnxVoiceCloneDir, "text_embedding.onnx"), sessionOptions);
        _codecEmbedding = CreateSessionOrThrow(Path.Combine(onnxVoiceCloneDir, "codec_embedding.onnx"), sessionOptions);
        _talkerDecode = CreateSessionOrThrow(Path.Combine(onnxVoiceCloneDir, "talker_decode.onnx"), sessionOptions);

        var codePredictorKvPath = Path.Combine(onnxVoiceCloneDir, "code_predictor_kv.onnx");
        var codePredictorPath = Path.Combine(onnxVoiceCloneDir, "code_predictor.onnx");
        if (File.Exists(codePredictorKvPath))
        {
            _codePredictor = CreateSessionOrThrow(codePredictorKvPath, sessionOptions);
            _codePredictorHasKv = true;
        }
        else if (File.Exists(codePredictorPath))
        {
            _codePredictor = CreateSessionOrThrow(codePredictorPath, sessionOptions);
            _codePredictorHasKv = false;
        }
        else
        {
            _codePredictor = null;
            _codePredictorHasKv = false;
        }

        for (var i = 0; i < _config.NumCodeGroups; i++)
        {
            var path = Path.Combine(onnxVoiceCloneDir, $"code_predictor_embed_g{i}.onnx");
            if (File.Exists(path))
            {
                _codePredictorEmbeds.Add(CreateSessionOrThrow(path, sessionOptions));
            }
            else
            {
                _codePredictorEmbeds.Add(null);
            }
        }

        _loadedKey = key;
    }

    private static void ValidateRequiredFiles(QwenSplitConfig cfg)
    {
        var required = new[]
        {
            Path.Combine(cfg.OnnxSharedDir, "speaker_encoder.onnx"),
            Path.Combine(cfg.OnnxSharedDir, "speech_tokenizer_decoder.onnx"),
            Path.Combine(cfg.OnnxVoiceCloneDir, "talker_decode.onnx"),
            Path.Combine(cfg.OnnxVoiceCloneDir, "text_embedding.onnx"),
            Path.Combine(cfg.OnnxVoiceCloneDir, "codec_embedding.onnx"),
            cfg.VocabPath,
            cfg.MergesPath,
            cfg.TokenizerConfigPath
        };
        foreach (var file in required)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException("Required Qwen split ONNX file missing.", file);
            }
        }
    }

    private static (string VocabPath, string MergesPath, string TokenizerConfigPath) ResolveTokenizerFiles(string repoRoot, string cacheDir)
    {
        var tokenizerDir = Path.Combine(repoRoot, "tokenizer");
        var vocab = Path.Combine(tokenizerDir, "vocab.json");
        var merges = Path.Combine(tokenizerDir, "merges.txt");
        var cfg = Path.Combine(tokenizerDir, "tokenizer_config.json");
        if (File.Exists(vocab) && File.Exists(merges) && File.Exists(cfg))
        {
            return (vocab, merges, cfg);
        }

        var vocabAny = Directory.GetFiles(repoRoot, "vocab.json", SearchOption.AllDirectories).FirstOrDefault();
        var mergesAny = Directory.GetFiles(repoRoot, "merges.txt", SearchOption.AllDirectories).FirstOrDefault();
        var cfgAny = Directory.GetFiles(repoRoot, "tokenizer_config.json", SearchOption.AllDirectories).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(vocabAny) && !string.IsNullOrWhiteSpace(mergesAny) && !string.IsNullOrWhiteSpace(cfgAny))
        {
            return (vocabAny, mergesAny, cfgAny);
        }

        var hfRoot = Path.Combine(cacheDir, "hf-cache");
        if (Directory.Exists(hfRoot))
        {
            var candidates = new List<(string VocabPath, string MergesPath, string TokenizerConfigPath)>();
            foreach (var vocabPath in Directory.GetFiles(hfRoot, "vocab.json", SearchOption.AllDirectories))
            {
                var dir = Path.GetDirectoryName(vocabPath);
                if (string.IsNullOrWhiteSpace(dir))
                {
                    continue;
                }

                var mergesPath = Path.Combine(dir, "merges.txt");
                var tokenizerCfgPath = Path.Combine(dir, "tokenizer_config.json");
                if (File.Exists(mergesPath) && File.Exists(tokenizerCfgPath))
                {
                    candidates.Add((vocabPath, mergesPath, tokenizerCfgPath));
                }
            }

            if (candidates.Count > 0)
            {
                // Prefer 1.7B tokenizer assets for the split 1.7B backend.
                var preferred = candidates.FirstOrDefault(c =>
                    c.VocabPath.IndexOf("1.7B", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.VocabPath.IndexOf("1_7B", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.VocabPath.IndexOf("12Hz-1.7B-Base", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.VocabPath.IndexOf("12Hz-1_7B-Base", StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrWhiteSpace(preferred.VocabPath))
                {
                    return preferred;
                }

                return candidates[0];
            }
        }

        throw new FileNotFoundException(
            "Tokenizer files missing. Expected tokenizer/vocab.json + merges.txt + tokenizer_config.json " +
            "in split repo or any installed HF cache tokenizer folder.");
    }

    private static string ResolveRustDllPath(string repoRoot, string cacheDir)
    {
        var bundled = QwenEmbeddedRuntime.EnsureBundledRustDllAtAppRoot();
        if (File.Exists(bundled))
        {
            return bundled;
        }

        var direct = Path.Combine(repoRoot, "qwen3_tts_rust.dll");
        if (File.Exists(direct))
        {
            return direct;
        }

        var repoScan = Directory.GetFiles(repoRoot, "qwen3_tts_rust.dll", SearchOption.AllDirectories).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(repoScan))
        {
            return repoScan;
        }

        var hfRoot = Path.Combine(cacheDir, "hf-cache");
        if (Directory.Exists(hfRoot))
        {
            var any = Directory.GetFiles(hfRoot, "qwen3_tts_rust.dll", SearchOption.AllDirectories).FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(any))
            {
                return any;
            }
        }

        throw new FileNotFoundException(
            "qwen3_tts_rust.dll missing. Download Qwen tokenizer/DLL support repo (zukky/Qwen3-TTS-ONNX-DLL) first.");
    }

    private static QwenSplitConfig LoadConfig(string cfgPath)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(cfgPath));
        var root = doc.RootElement;
        var talker = root.GetProperty("talker_config");

        var codePredictorCfg = talker.TryGetProperty("code_predictor_config", out var cp) ? cp : default;

        var lang = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        if (talker.TryGetProperty("codec_language_id", out var langObj))
        {
            foreach (var p in langObj.EnumerateObject())
            {
                if (p.Value.TryGetInt64(out var value))
                {
                    lang[p.Name] = value;
                }
            }
        }

        var speakerCfg = root.TryGetProperty("speaker_encoder_config", out var speaker) ? speaker : default;
        var speakerSampleRate = speakerCfg.ValueKind == JsonValueKind.Object &&
                                speakerCfg.TryGetProperty("sample_rate", out var sr)
            ? sr.GetInt32()
            : 24000;

        return new QwenSplitConfig
        {
            TtsBosTokenId = root.GetProperty("tts_bos_token_id").GetInt32(),
            TtsEosTokenId = root.GetProperty("tts_eos_token_id").GetInt32(),
            TtsPadTokenId = root.GetProperty("tts_pad_token_id").GetInt32(),
            HiddenSize = talker.GetProperty("hidden_size").GetInt32(),
            NumLayers = talker.GetProperty("num_hidden_layers").GetInt32(),
            NumKvHeads = talker.GetProperty("num_key_value_heads").GetInt32(),
            HeadDim = talker.GetProperty("head_dim").GetInt32(),
            NumCodeGroups = talker.GetProperty("num_code_groups").GetInt32(),
            TalkerVocabSize = talker.GetProperty("vocab_size").GetInt32(),
            CodecBosId = talker.GetProperty("codec_bos_id").GetInt64(),
            CodecEosTokenId = talker.GetProperty("codec_eos_token_id").GetInt32(),
            CodecPadId = talker.GetProperty("codec_pad_id").GetInt64(),
            CodecThinkId = talker.GetProperty("codec_think_id").GetInt64(),
            CodecNoThinkId = talker.GetProperty("codec_nothink_id").GetInt64(),
            CodecThinkBosId = talker.GetProperty("codec_think_bos_id").GetInt64(),
            CodecThinkEosId = talker.GetProperty("codec_think_eos_id").GetInt64(),
            CodecLanguageId = lang,
            SpeakerSampleRate = speakerSampleRate,
            CpNumLayers = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("num_hidden_layers", out var cpLayers)
                ? cpLayers.GetInt32()
                : 5,
            CpNumKvHeads = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("num_key_value_heads", out var cpHeads)
                ? cpHeads.GetInt32()
                : 8,
            CpHeadDim = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("head_dim", out var cpHeadDim)
                ? cpHeadDim.GetInt32()
                : 128,
            CodePredictorDoSample = codePredictorCfg.ValueKind == JsonValueKind.Object &&
                                    codePredictorCfg.TryGetProperty("do_sample", out var cpDoSample) &&
                                    cpDoSample.ValueKind is JsonValueKind.True or JsonValueKind.False
                ? cpDoSample.GetBoolean()
                : false,
            CodePredictorTopK = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("top_k", out var cpTopK)
                ? cpTopK.GetInt32()
                : 50,
            CodePredictorTopP = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("top_p", out var cpTopP)
                ? (float)cpTopP.GetDouble()
                : 1.0f,
            CodePredictorTemperature = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("temperature", out var cpTemp)
                ? (float)cpTemp.GetDouble()
                : 1.0f,
            CodePredictorRepetitionPenalty = codePredictorCfg.ValueKind == JsonValueKind.Object && codePredictorCfg.TryGetProperty("repetition_penalty", out var cpRep)
                ? (float)cpRep.GetDouble()
                : 1.0f
        };
    }

    private static SessionOptions CreateSessionOptions(string preferDevice, out string resolvedProvider)
    {
        var options = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_EXTENDED };
        var normalized = (preferDevice ?? string.Empty).Trim().ToLowerInvariant();

        if (normalized == "cpu")
        {
            resolvedProvider = "cpu";
            return options;
        }

        if (normalized == "dml")
        {
            if (!OperatingSystem.IsWindows())
                throw new InvalidOperationException("DML provider is only available on Windows.");
            options.AppendExecutionProvider_DML(0);
            resolvedProvider = "dml";
            return options;
        }

        // gpu or auto: always try CUDA first. DML and CUDA are mutually exclusive
        // OnnxRuntime packages — trying DML when CUDA is installed throws immediately.
        TryBootstrapCudaRuntimePaths();
        try
        {
            options.AppendExecutionProvider_CUDA(0);
            resolvedProvider = "cuda";
            return options;
        }
        catch (Exception cudaEx)
        {
            if (normalized == "gpu")
            {
                // DML as last resort on Windows only
                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        options.AppendExecutionProvider_DML(0);
                        resolvedProvider = "dml";
                        return options;
                    }
                    catch (Exception dmlEx)
                    {
                        throw new InvalidOperationException(
                            $"GPU was selected but neither CUDA nor DML are available. CUDA: {cudaEx.Message}; DML: {dmlEx.Message}");
                    }
                }
                throw new InvalidOperationException(
                    $"GPU was selected but CUDA is unavailable: {cudaEx.Message}");
            }

            // auto: CUDA failed, try DML, then fall back to CPU silently
            if (OperatingSystem.IsWindows())
            {
                try
                {
                    options.AppendExecutionProvider_DML(0);
                    resolvedProvider = "dml";
                    return options;
                }
                catch { }
            }

            resolvedProvider = "cpu";
            return options;
        }
    }

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
                .Where(dir => !string.IsNullOrWhiteSpace(dir))
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

            var merged = string.Join(";", additions) + ";" + existingPath;
            Environment.SetEnvironmentVariable("PATH", merged);
        }
    }

    private static InferenceSession CreateSessionOrThrow(string modelPath, SessionOptions options)
    {
        try
        {
            return new InferenceSession(modelPath, options);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Qwen split ONNX file is invalid or incomplete: {modelPath}. Remove cache and re-download.",
                ex);
        }
    }

    private static string BuildAssistantPrompt(string text)
    {
        // Do NOT put RefText in the prompt: the model would speak it and produce garbage. RefText is reserved for future ref_text/ICL input when the ONNX graph supports it.
        return $"<|im_start|>assistant\n{text}<|im_end|>\n<|im_start|>assistant\n";
    }

    private static float[] RunSpeakerEncoder(InferenceSession session, float[] melInput)
    {
        var seq = melInput.Length / 128;
        var melTensor = new DenseTensor<float>(melInput, new[] { 1, seq, 128 });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("mel_spectrogram", melTensor) });
        return TensorToFloat(results.First());
    }

    private sealed class InputState
    {
        public float[] InputsEmbeds { get; set; } = Array.Empty<float>();
        public float[] AttentionMask { get; set; } = Array.Empty<float>();
        public float[] TrailingTextHidden { get; set; } = Array.Empty<float>();
        public float[] TtsPadEmbed { get; set; } = Array.Empty<float>();
    }

    private static InputState BuildInputState(
        long[] inputIds,
        float[] speakerEmbedding,
        QwenSplitConfig cfg,
        InferenceSession textEmbedding,
        InferenceSession codecEmbedding)
    {
        if (inputIds.Length < 10)
        {
            throw new InvalidOperationException("Qwen tokenizer prompt is too short.");
        }

        var hidden = cfg.HiddenSize;
        var ttsEmbeds = RunTextEmbedding(textEmbedding, new[] { (long)cfg.TtsBosTokenId, (long)cfg.TtsEosTokenId, (long)cfg.TtsPadTokenId });
        var ttsBos = SliceSeq(ttsEmbeds, hidden, 0, 1);
        var ttsEos = SliceSeq(ttsEmbeds, hidden, 1, 1);
        var ttsPad = SliceSeq(ttsEmbeds, hidden, 2, 1);

        long[] codecPrefix;
        if (cfg.CodecLanguageId.TryGetValue("english", out var langId))
        {
            codecPrefix = new[] { cfg.CodecThinkId, cfg.CodecThinkBosId, langId, cfg.CodecThinkEosId };
        }
        else if (cfg.CodecLanguageId.TryGetValue("chinese", out var zhId))
        {
            codecPrefix = new[] { cfg.CodecThinkId, cfg.CodecThinkBosId, zhId, cfg.CodecThinkEosId };
        }
        else
        {
            codecPrefix = new[] { cfg.CodecNoThinkId, cfg.CodecThinkBosId, cfg.CodecThinkEosId };
        }

        var codecPrefixEmb = RunCodecEmbedding(codecEmbedding, codecPrefix);
        var codecPadBosEmb = RunCodecEmbedding(codecEmbedding, new[] { cfg.CodecPadId, cfg.CodecBosId });
        var speakerEmb = new float[hidden];
        Array.Copy(speakerEmbedding, speakerEmb, Math.Min(hidden, speakerEmbedding.Length));
        var codecInputEmb = ConcatSeq(codecPrefixEmb, speakerEmb, codecPadBosEmb);

        var roleEmbed = RunTextEmbedding(textEmbedding, SliceIds(inputIds, 0, 3));
        var padRepeat = Math.Max(0, SeqLen(codecInputEmb, hidden) - 2);
        var ttsPrefix = ConcatSeq(RepeatSeq(ttsPad, hidden, padRepeat), ttsBos);
        var combined = AddSeq(ttsPrefix, SliceSeq(codecInputEmb, hidden, 0, SeqLen(codecInputEmb, hidden) - 1));
        var talkerInput = ConcatSeq(roleEmbed, combined);

        var firstText = AddSeq(
            RunTextEmbedding(textEmbedding, SliceIds(inputIds, 3, 1)),
            SliceSeq(codecInputEmb, hidden, SeqLen(codecInputEmb, hidden) - 1, 1));
        talkerInput = ConcatSeq(talkerInput, firstText);

        var trailingTextIds = SliceIds(inputIds, 4, Math.Max(0, inputIds.Length - 9));
        var trailingText = trailingTextIds.Length > 0
            ? RunTextEmbedding(textEmbedding, trailingTextIds)
            : Array.Empty<float>();
        var trailingHidden = ConcatSeq(trailingText, ttsEos);

        return new InputState
        {
            InputsEmbeds = talkerInput,
            AttentionMask = Enumerable.Repeat(1.0f, SeqLen(talkerInput, hidden)).ToArray(),
            TrailingTextHidden = trailingHidden,
            TtsPadEmbed = ttsPad
        };
    }

    private long[,] GenerateCodes(
        InputState state,
        QwenSplitConfig cfg,
        InferenceSession talkerDecode,
        InferenceSession? codePredictor,
        InferenceSession codecEmbedding,
        IReadOnlyList<InferenceSession?> codePredictorEmbeds,
        bool codePredictorHasKv,
        QwenDecodeOptions decode,
        CancellationToken ct)
    {
        var hidden = cfg.HiddenSize;
        var minNewTokens = 2;
        var suppressStart = Math.Max(0, cfg.TalkerVocabSize - 1024);
        var trailingSeq = SeqLen(state.TrailingTextHidden, hidden);
        if (trailingSeq <= 0)
            return new long[0, cfg.NumCodeGroups];
        var trailingIndex = 0;
        var firstTokenHistory = new List<long>();
        var generated = new List<long[]>();

        var talkerPast = new float[0];
        var talkerPastSeq = 0;
        var prefillSeq = SeqLen(state.InputsEmbeds, hidden);
        var prefillPos = BuildPositionIdsRange(prefillSeq);
        var prefill = RunTalker(
            talkerDecode,
            state.InputsEmbeds,
            prefillSeq,
            state.AttentionMask,
            prefillPos,
            talkerPast,
            talkerPastSeq,
            cfg);

        var logits = prefill.LogitsLast;
        var lastHidden = prefill.HiddenLast;
        talkerPast = prefill.PresentKeyValues;
        talkerPastSeq = prefill.PresentSeqLen;

        for (var step = 0; ; step++)
        {
            ct.ThrowIfCancellationRequested();

            var stepLogits = (float[])logits.Clone();
            for (var tok = suppressStart; tok < cfg.TalkerVocabSize && tok < stepLogits.Length; tok++)
            {
                if (tok == cfg.CodecEosTokenId)
                {
                    continue;
                }
                stepLogits[tok] = -1.0e9f;
            }
            if (step < minNewTokens && cfg.CodecEosTokenId >= 0 && cfg.CodecEosTokenId < stepLogits.Length)
            {
                stepLogits[cfg.CodecEosTokenId] = -1.0e9f;
            }

            ApplyRepetitionPenalty(stepLogits, firstTokenHistory, decode.RepetitionPenalty);
            var firstCode = SelectToken(stepLogits, decode);
            if (firstCode == cfg.CodecEosTokenId)
            {
                break;
            }

            firstTokenHistory.Add(firstCode);
            var stepCodes = new List<long> { firstCode };
            var firstCodeEmbed = RunCodecEmbedding(codecEmbedding, new[] { firstCode });
            var talkerHidden = lastHidden;

            if (codePredictor is not null && talkerHidden.Length == hidden)
            {
                if (codePredictorHasKv)
                {
                    var cpPast = new float[0];
                    var cpPastSeq = 0;
                    var cpInput = ConcatSeq(talkerHidden, firstCodeEmbed);

                    var pred = RunCodePredictorKv(codePredictor, cpInput, hidden, cpPast, cpPastSeq, cfg, generationStep: 0);
                    var predCode = SelectToken(pred.Logits, decode);
                    stepCodes.Add(predCode);
                    cpPast = pred.PresentKeyValues;
                    cpPastSeq = pred.PresentSeqLen;

                    for (var g = 1; g < cfg.NumCodeGroups - 1; g++)
                    {
                        var subEmb = RunCodePredictorEmbedding(codePredictorEmbeds, codecEmbedding, g, predCode);
                        pred = RunCodePredictorKv(codePredictor, subEmb, hidden, cpPast, cpPastSeq, cfg, g);
                        predCode = SelectToken(pred.Logits, decode);
                        stepCodes.Add(predCode);
                        cpPast = pred.PresentKeyValues;
                        cpPastSeq = pred.PresentSeqLen;
                    }
                }
                else
                {
                    var predictorInput = ConcatSeq(talkerHidden, firstCodeEmbed);
                    for (var g = 0; g < cfg.NumCodeGroups - 1; g++)
                    {
                        var predLogits = RunCodePredictor(codePredictor, predictorInput, hidden, g);
                        var predCode = SelectToken(predLogits, decode);
                        stepCodes.Add(predCode);
                        var embedIdx = g + 1;
                        var subEmb = RunCodePredictorEmbedding(codePredictorEmbeds, codecEmbedding, embedIdx, predCode);
                        predictorInput = ConcatSeq(predictorInput, subEmb);
                    }
                }
            }
            else
            {
                while (stepCodes.Count < cfg.NumCodeGroups)
                {
                    stepCodes.Add(0);
                }
            }

            if (stepCodes.Count > cfg.NumCodeGroups)
            {
                stepCodes = stepCodes.Take(cfg.NumCodeGroups).ToList();
            }
            while (stepCodes.Count < cfg.NumCodeGroups)
            {
                stepCodes.Add(0);
            }

            var nextCodecEmbed = firstCodeEmbed;
            for (var g = 1; g < stepCodes.Count; g++)
            {
                var embedG = RunCodePredictorEmbedding(codePredictorEmbeds, codecEmbedding, g, stepCodes[g]);
                nextCodecEmbed = AddSeq(nextCodecEmbed, embedG);
            }

            // Save codes for this step before deciding whether to continue.
            generated.Add(stepCodes.ToArray());

            // All real text tokens consumed — input is finished, stop generating.
            // Never feed padding: fake input causes the model to hallucinate silence.
            if (trailingIndex >= trailingSeq)
                break;
            var nextTextEmbed = SliceSeq(state.TrailingTextHidden, hidden, trailingIndex++, 1);
            var nextEmbed = AddSeq(nextTextEmbed, nextCodecEmbed);

            var currLen = talkerPastSeq + 1;
            var attn = Enumerable.Repeat(1.0f, currLen).ToArray();
            var pos = BuildPositionIdsSingle(currLen - 1);
            var stepOut = RunTalker(
                talkerDecode,
                nextEmbed,
                1,
                attn,
                pos,
                talkerPast,
                talkerPastSeq,
                cfg);

            logits = stepOut.LogitsLast;
            lastHidden = stepOut.HiddenLast;
            talkerPast = stepOut.PresentKeyValues;
            talkerPastSeq = stepOut.PresentSeqLen;
        }

        if (generated.Count == 0)
        {
            return new long[0, cfg.NumCodeGroups];
        }

        var output = new long[generated.Count, cfg.NumCodeGroups];
        for (var i = 0; i < generated.Count; i++)
        {
            for (var g = 0; g < cfg.NumCodeGroups; g++)
            {
                output[i, g] = generated[i][g];
            }
        }
        return output;
    }

    private static float[] RunTextEmbedding(InferenceSession session, long[] ids)
    {
        if (ids.Length == 0)
        {
            return Array.Empty<float>();
        }

        var tensor = new DenseTensor<long>(ids, new[] { 1, ids.Length });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("input_ids", tensor) });
        return TensorToFloat(results.First());
    }

    private static float[] RunCodecEmbedding(InferenceSession session, long[] ids)
    {
        if (ids.Length == 0)
        {
            return Array.Empty<float>();
        }

        var tensor = new DenseTensor<long>(ids, new[] { 1, ids.Length });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("codec_ids", tensor) });
        return TensorToFloat(results.First());
    }

    private static float[] RunCodePredictorEmbedding(
        IReadOnlyList<InferenceSession?> embedSessions,
        InferenceSession codecEmbedding,
        int group,
        long codeId)
    {
        if (group >= 0 && group < embedSessions.Count && embedSessions[group] is not null)
        {
            var tensor = new DenseTensor<long>(new[] { codeId }, new[] { 1, 1 });
            using var results = embedSessions[group]!.Run(new[] { NamedOnnxValue.CreateFromTensor("code_ids", tensor) });
            return TensorToFloat(results.First());
        }

        return RunCodecEmbedding(codecEmbedding, new[] { codeId });
    }

    private static float[] RunCodePredictor(InferenceSession session, float[] inputsEmbeds, int hidden, int generationStep)
    {
        var seq = SeqLen(inputsEmbeds, hidden);
        var embedTensor = new DenseTensor<float>(inputsEmbeds, new[] { 1, seq, hidden });
        var stepTensor = new DenseTensor<long>(new[] { (long)generationStep }, new[] { 1 });
        using var results = session.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("inputs_embeds", embedTensor),
            NamedOnnxValue.CreateFromTensor("generation_step", stepTensor)
        });

        var first = results.First();
        var logits = TensorToFloat(first);
        var dims = first.AsTensor<float>().Dimensions.ToArray();
        if (dims.Length == 3)
        {
            var seqLen = dims[1];
            var vocab = dims[2];
            var last = new float[vocab];
            Array.Copy(logits, (seqLen - 1) * vocab, last, 0, vocab);
            return last;
        }

        return logits;
    }

    private sealed class CodePredictorKvOutput
    {
        public float[] Logits { get; init; } = Array.Empty<float>();
        public float[] PresentKeyValues { get; init; } = Array.Empty<float>();
        public int PresentSeqLen { get; init; }
    }

    private static CodePredictorKvOutput RunCodePredictorKv(
        InferenceSession session,
        float[] inputsEmbeds,
        int hidden,
        float[] pastKeyValues,
        int pastSeqLen,
        QwenSplitConfig cfg,
        int generationStep)
    {
        var seq = SeqLen(inputsEmbeds, hidden);
        var embedTensor = new DenseTensor<float>(inputsEmbeds, new[] { 1, seq, hidden });
        var pastTensor = new DenseTensor<float>(
            pastKeyValues,
            new[] { cfg.CpNumLayers, 2, 1, cfg.CpNumKvHeads, pastSeqLen, cfg.CpHeadDim });
        var stepTensor = new DenseTensor<long>(new[] { (long)generationStep }, new[] { 1 });

        using var results = session.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("inputs_embeds", embedTensor),
            NamedOnnxValue.CreateFromTensor("past_key_values", pastTensor),
            NamedOnnxValue.CreateFromTensor("generation_step", stepTensor)
        });

        var list = results.ToList();
        var first = list[0];
        var logitsRaw = TensorToFloat(first);
        var logitsDims = first.AsTensor<float>().Dimensions.ToArray();
        float[] logits;
        if (logitsDims.Length == 3)
        {
            var seqLen = logitsDims[1];
            var vocab = logitsDims[2];
            logits = new float[vocab];
            Array.Copy(logitsRaw, (seqLen - 1) * vocab, logits, 0, vocab);
        }
        else
        {
            logits = logitsRaw;
        }

        var presentTensor = list.Count > 1 ? list[1].AsTensor<float>() : null;
        var present = presentTensor is null ? Array.Empty<float>() : presentTensor.ToArray();
        var presentSeq = pastSeqLen;
        if (presentTensor is not null)
        {
            var dims = presentTensor.Dimensions.ToArray();
            if (dims.Length >= 5)
            {
                presentSeq = dims[4];
            }
        }

        return new CodePredictorKvOutput
        {
            Logits = logits,
            PresentKeyValues = present,
            PresentSeqLen = presentSeq
        };
    }

    private sealed class TalkerStepOutput
    {
        public float[] LogitsLast { get; init; } = Array.Empty<float>();
        public float[] HiddenLast { get; init; } = Array.Empty<float>();
        public float[] PresentKeyValues { get; init; } = Array.Empty<float>();
        public int PresentSeqLen { get; init; }
    }

    private static TalkerStepOutput RunTalker(
        InferenceSession session,
        float[] inputsEmbeds,
        int seq,
        float[] attentionMask,
        long[] positionIds,
        float[] pastKeyValues,
        int pastSeqLen,
        QwenSplitConfig cfg)
    {
        var embedTensor = new DenseTensor<float>(inputsEmbeds, new[] { 1, seq, cfg.HiddenSize });
        var maskTensor = new DenseTensor<float>(attentionMask, new[] { 1, attentionMask.Length });
        var posTensor = new DenseTensor<long>(positionIds, new[] { 3, 1, seq });
        var pastTensor = new DenseTensor<float>(
            pastKeyValues,
            new[] { cfg.NumLayers, 2, 1, cfg.NumKvHeads, pastSeqLen, cfg.HeadDim });

        using var results = session.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("inputs_embeds", embedTensor),
            NamedOnnxValue.CreateFromTensor("attention_mask", maskTensor),
            NamedOnnxValue.CreateFromTensor("position_ids", posTensor),
            NamedOnnxValue.CreateFromTensor("past_key_values", pastTensor)
        });

        var list = results.ToList();
        var logitsTensor = list[0].AsTensor<float>();
        var logits = logitsTensor.ToArray();
        var logitsSeq = logitsTensor.Dimensions[1];
        var logitsVocab = logitsTensor.Dimensions[2];
        var logitsLast = new float[logitsVocab];
        Array.Copy(logits, (logitsSeq - 1) * logitsVocab, logitsLast, 0, logitsVocab);

        float[] hiddenLast = new float[cfg.HiddenSize];
        if (list.Count > 2)
        {
            var hiddenTensor = list[2].AsTensor<float>();
            var hidden = hiddenTensor.ToArray();
            var hiddenSeq = hiddenTensor.Dimensions[1];
            Array.Copy(hidden, (hiddenSeq - 1) * cfg.HiddenSize, hiddenLast, 0, cfg.HiddenSize);
        }

        var presentKvTensor = list.Count > 1 ? list[1].AsTensor<float>() : null;
        var presentKv = presentKvTensor?.ToArray() ?? Array.Empty<float>();
        var presentSeq = presentKvTensor is null ? pastSeqLen : presentKvTensor.Dimensions[4];

        return new TalkerStepOutput
        {
            LogitsLast = logitsLast,
            HiddenLast = hiddenLast,
            PresentKeyValues = presentKv,
            PresentSeqLen = presentSeq
        };
    }

    private static float[] DecodeAudioCodes(InferenceSession session, long[,] codes)
    {
        var steps = codes.GetLength(0);
        var groups = codes.GetLength(1);
        var flat = new long[steps * groups];
        var idx = 0;
        for (var i = 0; i < steps; i++)
        {
            for (var g = 0; g < groups; g++)
            {
                flat[idx++] = codes[i, g];
            }
        }

        var codeTensor = new DenseTensor<long>(flat, new[] { 1, steps, groups });
        var codeInputName = session.InputNames.Count > 0 ? session.InputNames[0] : "codes";
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor(codeInputName, codeTensor) });
        return TensorToFloat(results.First());
    }

    private long SelectToken(float[] logits, QwenDecodeOptions decode)
    {
        if (!decode.DoSample || decode.TopK <= 1 || decode.Temperature <= 0.05f)
        {
            return ArgMaxToken(logits);
        }

        return SampleToken(logits, decode.TopK, decode.TopP, decode.Temperature);
    }

    private static long ArgMaxToken(float[] logits)
    {
        if (logits.Length == 0)
        {
            return 0;
        }

        var bestIndex = 0;
        var best = logits[0];
        for (var i = 1; i < logits.Length; i++)
        {
            if (logits[i] > best)
            {
                best = logits[i];
                bestIndex = i;
            }
        }
        return bestIndex;
    }

    private long SampleToken(float[] logits, int topK, float topP, float temperature)
    {
        var temp = Math.Max(0.01f, temperature);
        var order = Enumerable.Range(0, logits.Length)
            .OrderByDescending(i => logits[i] / temp)
            .ToArray();
        if (topK > 0 && topK < order.Length)
        {
            order = order.Take(topK).ToArray();
        }

        var max = order.Max(i => logits[i] / temp);
        var probs = new double[order.Length];
        var sum = 0.0;
        for (var i = 0; i < order.Length; i++)
        {
            probs[i] = Math.Exp((logits[order[i]] / temp) - max);
            sum += probs[i];
        }
        if (sum <= 0 || double.IsNaN(sum))
        {
            return order[0];
        }
        for (var i = 0; i < probs.Length; i++)
        {
            probs[i] /= sum;
        }

        if (topP < 1.0f)
        {
            var cumulative = 0.0;
            var keep = probs.Length;
            for (var i = 0; i < probs.Length; i++)
            {
                cumulative += probs[i];
                if (cumulative >= topP)
                {
                    keep = i + 1;
                    break;
                }
            }
            order = order.Take(keep).ToArray();
            probs = probs.Take(keep).ToArray();
            var renorm = probs.Sum();
            if (renorm > 0)
            {
                for (var i = 0; i < probs.Length; i++)
                {
                    probs[i] /= renorm;
                }
            }
        }

        var p = _rng.NextDouble();
        var walk = 0.0;
        for (var i = 0; i < probs.Length; i++)
        {
            walk += probs[i];
            if (p <= walk)
            {
                return order[i];
            }
        }

        return order[^1];
    }

    private static QwenAudioQuality AnalyzeAudioQuality(float[] audio, int sampleRate)
    {
        if (audio.Length == 0 || sampleRate <= 0)
        {
            return new QwenAudioQuality(true, "empty");
        }

        var peak = 0f;
        double sumSq = 0;
        var clipCount = 0;
        var zeroCrossings = 0;
        var nearZeroDiff = 0;
        for (var i = 0; i < audio.Length; i++)
        {
            var s = audio[i];
            var abs = Math.Abs(s);
            if (abs > peak)
            {
                peak = abs;
            }
            if (abs >= 0.995f)
            {
                clipCount++;
            }
            sumSq += s * s;
            if (i > 0)
            {
                var prev = audio[i - 1];
                if ((prev < 0 && s >= 0) || (prev >= 0 && s < 0))
                {
                    zeroCrossings++;
                }
                if (Math.Abs(s - prev) < 1e-5f)
                {
                    nearZeroDiff++;
                }
            }
        }

        var rms = Math.Sqrt(sumSq / Math.Max(1, audio.Length));
        var zcr = zeroCrossings / (double)Math.Max(1, audio.Length - 1);
        var clipRatio = clipCount / (double)audio.Length;
        var flatRatio = nearZeroDiff / (double)Math.Max(1, audio.Length - 1);
        var durationSec = audio.Length / (double)sampleRate;

        var suspicious =
            durationSec < 0.20 ||
            peak < 0.02f ||
            rms < 0.003 ||
            clipRatio > 0.03 ||
            (rms < 0.015 && zcr > 0.20) ||
            flatRatio > 0.90;

        return new QwenAudioQuality(
            suspicious,
            $"dur={durationSec:0.00}s rms={rms:0.0000} peak={peak:0.000} zcr={zcr:0.000} clip={clipRatio:P0}");
    }

    private sealed record QwenAudioQuality(bool IsSuspicious, string Summary);

    private sealed record QwenDecodeOptions(
        bool DoSample,
        int TopK,
        float TopP,
        float Temperature,
        float RepetitionPenalty)
    {
        public static QwenDecodeOptions From(QwenSplitConfig cfg, TtsRequest request)
        {
            var doSample = request.QwenDoSample ?? cfg.CodePredictorDoSample;
            var topK = request.QwenTopK ?? cfg.CodePredictorTopK;
            var topP = request.QwenTopP ?? cfg.CodePredictorTopP;
            var temperature = request.QwenTemperature ?? cfg.CodePredictorTemperature;
            var repetitionPenalty = request.QwenRepetitionPenalty ?? cfg.CodePredictorRepetitionPenalty;
            return new QwenDecodeOptions(
                doSample,
                Math.Clamp(topK <= 0 ? 1 : topK, 1, 200),
                Math.Clamp(topP <= 0 ? 1.0f : topP, 0.05f, 1.0f),
                Math.Clamp(temperature <= 0 ? 1.0f : temperature, 0.05f, 2.0f),
                Math.Clamp(repetitionPenalty <= 0 ? 1.0f : repetitionPenalty, 1.0f, 2.0f));
        }

        public QwenDecodeOptions MakeSafer(int attempt)
        {
            if (attempt >= 2)
            {
                // Final fallback: fully deterministic greedy decoding.
                return this with
                {
                    DoSample = false,
                    TopK = 1,
                    TopP = 1.0f,
                    Temperature = 0.1f,
                    RepetitionPenalty = Math.Clamp(Math.Max(RepetitionPenalty, 1.08f), 1.0f, 2.0f)
                };
            }

            if (attempt >= 1)
            {
                // Second attempt: tighter sampling, lower temperature.
                return this with
                {
                    DoSample = false,
                    TopK = 1,
                    TopP = 1.0f,
                    Temperature = 0.2f,
                    RepetitionPenalty = Math.Clamp(Math.Max(RepetitionPenalty, 1.05f), 1.0f, 2.0f)
                };
            }

            // First retry: reduce sampling diversity but still allow minor variation.
            return this with
            {
                DoSample = false,
                TopK = Math.Clamp(Math.Min(TopK, 8), 1, 200),
                TopP = Math.Clamp(Math.Min(TopP, 0.88f), 0.05f, 1.0f),
                Temperature = Math.Clamp(Math.Min(Temperature, 0.50f), 0.05f, 2.0f),
                RepetitionPenalty = Math.Clamp(Math.Max(RepetitionPenalty, 1.05f), 1.0f, 2.0f)
            };
        }
    }

    private static void ApplyRepetitionPenalty(float[] logits, IReadOnlyList<long> history, float penalty)
    {
        if (history.Count == 0 || Math.Abs(penalty - 1f) < 1e-6)
        {
            return;
        }
        foreach (var tok in history.Distinct())
        {
            if (tok < 0 || tok >= logits.Length)
            {
                continue;
            }

            var s = logits[tok];
            logits[tok] = s >= 0 ? s / penalty : s * penalty;
        }
    }

    private static long[] BuildPositionIdsRange(int seqLen)
    {
        var output = new long[3 * seqLen];
        var idx = 0;
        for (var row = 0; row < 3; row++)
        {
            for (var i = 0; i < seqLen; i++)
            {
                output[idx++] = i;
            }
        }
        return output;
    }

    private static long[] BuildPositionIdsSingle(int value) => new[] { (long)value, (long)value, (long)value };

    private static int SeqLen(float[] seqTensor, int hidden) => hidden <= 0 ? 0 : seqTensor.Length / hidden;

    private static long[] SliceIds(long[] ids, int start, int len)
    {
        if (start < 0) start = 0;
        if (start >= ids.Length || len <= 0) return Array.Empty<long>();
        var take = Math.Min(len, ids.Length - start);
        var output = new long[take];
        Array.Copy(ids, start, output, 0, take);
        return output;
    }

    private static float[] SliceSeq(float[] src, int hidden, int start, int len)
    {
        var seq = SeqLen(src, hidden);
        if (seq == 0 || len <= 0 || start >= seq)
        {
            return Array.Empty<float>();
        }
        if (start < 0) start = 0;
        var take = Math.Min(len, seq - start);
        var output = new float[take * hidden];
        Array.Copy(src, start * hidden, output, 0, output.Length);
        return output;
    }

    private static float[] RepeatSeq(float[] oneStepSeq, int hidden, int count)
    {
        if (count <= 0) return Array.Empty<float>();
        var output = new float[count * hidden];
        for (var i = 0; i < count; i++)
        {
            Array.Copy(oneStepSeq, 0, output, i * hidden, hidden);
        }
        return output;
    }

    private static float[] ConcatSeq(params float[][] items)
    {
        var total = items.Sum(x => x.Length);
        var output = new float[total];
        var pos = 0;
        foreach (var x in items)
        {
            if (x.Length == 0)
            {
                continue;
            }
            Array.Copy(x, 0, output, pos, x.Length);
            pos += x.Length;
        }
        return output;
    }

    private static float[] AddSeq(float[] a, float[] b)
    {
        if (a.Length != b.Length)
        {
            throw new InvalidOperationException("AddSeq tensor shape mismatch.");
        }
        var output = new float[a.Length];
        for (var i = 0; i < a.Length; i++)
        {
            output[i] = a[i] + b[i];
        }
        return output;
    }

    private static float[] TensorToFloat(DisposableNamedOnnxValue value)
    {
        try
        {
            return value.AsTensor<float>().ToArray();
        }
        catch
        {
            var half = value.AsTensor<Half>().ToArray();
            var output = new float[half.Length];
            for (var i = 0; i < half.Length; i++)
            {
                output[i] = (float)half[i];
            }
            return output;
        }
    }

    private static float[] ApplySpeedResample(float[] input, float speed)
    {
        speed = Math.Clamp(speed, 0.55f, 1.75f);
        if (input.Length == 0 || Math.Abs(speed - 1.0f) < 0.001f)
        {
            return input;
        }

        var outputLength = Math.Max(1, (int)Math.Round(input.Length / speed));
        var output = new float[outputLength];
        for (var i = 0; i < outputLength; i++)
        {
            var sourceIndex = i * speed;
            var lo = (int)Math.Floor(sourceIndex);
            if (lo >= input.Length)
            {
                output[i] = input[^1];
                continue;
            }
            var hi = Math.Min(lo + 1, input.Length - 1);
            var frac = sourceIndex - lo;
            output[i] = (float)((input[lo] * (1.0 - frac)) + (input[hi] * frac));
        }
        return output;
    }

    private static float[] TransposeMelToInput(float[,] mel)
    {
        var rows = mel.GetLength(0);
        var cols = mel.GetLength(1);
        var output = new float[rows * cols];
        var idx = 0;
        for (var c = 0; c < cols; c++)
        {
            for (var r = 0; r < rows; r++)
            {
                output[idx++] = mel[r, c];
            }
        }
        return output;
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

        bw.Write(Encoding.ASCII.GetBytes("RIFF"));
        bw.Write(36 + dataSize);
        bw.Write(Encoding.ASCII.GetBytes("WAVEfmt "));
        bw.Write(16);
        bw.Write((short)1);
        bw.Write(channels);
        bw.Write(sampleRate);
        bw.Write(byteRate);
        bw.Write(blockAlign);
        bw.Write(bitsPerSample);
        bw.Write(Encoding.ASCII.GetBytes("data"));
        bw.Write(dataSize);
        foreach (var s in samples)
        {
            bw.Write((short)Math.Round(Math.Clamp(s, -1f, 1f) * 32767f));
        }
    }

    private sealed class QwenSplitConfig
    {
        public int TtsBosTokenId { get; init; }
        public int TtsEosTokenId { get; init; }
        public int TtsPadTokenId { get; init; }
        public int HiddenSize { get; init; }
        public int NumLayers { get; init; }
        public int NumKvHeads { get; init; }
        public int HeadDim { get; init; }
        public int NumCodeGroups { get; init; }
        public int TalkerVocabSize { get; init; }
        public long CodecBosId { get; init; }
        public int CodecEosTokenId { get; init; }
        public long CodecPadId { get; init; }
        public long CodecThinkId { get; init; }
        public long CodecNoThinkId { get; init; }
        public long CodecThinkBosId { get; init; }
        public long CodecThinkEosId { get; init; }
        public Dictionary<string, long> CodecLanguageId { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public int SpeakerSampleRate { get; init; } = 24000;
        public int SpeakerNfft { get; init; } = 1024;
        public int SpeakerHop { get; init; } = 256;
        public int SpeakerWin { get; init; } = 1024;
        public int SpeakerMels { get; init; } = 128;
        public float SpeakerFMin { get; init; } = 0f;
        public float SpeakerFMax { get; init; } = 12000f;
        public int OutputSampleRate { get; init; } = 24000;
        public int CpNumLayers { get; init; } = 5;
        public int CpNumKvHeads { get; init; } = 8;
        public int CpHeadDim { get; init; } = 128;
        // Defaults match the official Python qwen-tts package inference params.
        public bool CodePredictorDoSample { get; init; } = true;
        public int CodePredictorTopK { get; init; } = 50;
        public float CodePredictorTopP { get; init; } = 1.0f;
        public float CodePredictorTemperature { get; init; } = 0.9f;
        public float CodePredictorRepetitionPenalty { get; init; } = 1.05f;
        public string OnnxSharedDir { get; set; } = string.Empty;
        public string OnnxVoiceCloneDir { get; set; } = string.Empty;
        public string VocabPath { get; set; } = string.Empty;
        public string MergesPath { get; set; } = string.Empty;
        public string TokenizerConfigPath { get; set; } = string.Empty;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MelConfigNative
    {
        public uint sample_rate;
        public nuint n_fft;
        public nuint hop_length;
        public nuint win_length;
        public nuint n_mels;
        public float fmin;
        public float fmax;
    }

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_last_error_message")]
    private static extern nuint LastErrorMessage(nint outBuf, nuint outLen);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_read_wav_f32")]
    private static extern nuint ReadWavNative([MarshalAs(UnmanagedType.LPUTF8Str)] string path, nint outBuf, nuint outLen, ref uint outSr);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_resample_f32")]
    private static extern nuint ResampleNative(nint input, nuint inputLen, uint srcRate, uint dstRate, nint outBuf, nuint outLen);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_mel_f32")]
    private static extern nuint MelNative(nint input, nuint inputLen, ref MelConfigNative cfg, nint outBuf, nuint outLen, ref nuint rows, ref nuint cols);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_create")]
    private static extern nint CreateTokenizer([MarshalAs(UnmanagedType.LPUTF8Str)] string vocabPath, [MarshalAs(UnmanagedType.LPUTF8Str)] string mergesPath, [MarshalAs(UnmanagedType.LPUTF8Str)] string tokenizerConfigPath);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_free")]
    private static extern void FreeTokenizer(nint handle);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_encode")]
    private static extern nuint TokenizerEncodeNative(nint handle, [MarshalAs(UnmanagedType.LPUTF8Str)] string text, nint outIds, nuint outLen);

    private static (float[] Samples, int SampleRate) ReadWav(string path)
    {
        uint sr = 0;
        var needed = ReadWavNative(path, nint.Zero, 0, ref sr);
        if (needed == 0)
        {
            throw new InvalidOperationException(GetLastNativeError());
        }
        var output = new float[(int)needed];
        var handle = GCHandle.Alloc(output, GCHandleType.Pinned);
        try
        {
            var got = ReadWavNative(path, handle.AddrOfPinnedObject(), (nuint)output.Length, ref sr);
            if (got == 0)
            {
                throw new InvalidOperationException(GetLastNativeError());
            }
            if ((int)got != output.Length)
            {
                Array.Resize(ref output, (int)got);
            }
            return (output, (int)sr);
        }
        finally
        {
            handle.Free();
        }
    }

    private static float[] Resample(float[] input, int srcRate, int dstRate)
    {
        var inHandle = GCHandle.Alloc(input, GCHandleType.Pinned);
        try
        {
            var needed = ResampleNative(inHandle.AddrOfPinnedObject(), (nuint)input.Length, (uint)srcRate, (uint)dstRate, nint.Zero, 0);
            if (needed == 0)
            {
                throw new InvalidOperationException(GetLastNativeError());
            }
            var output = new float[(int)needed];
            var outHandle = GCHandle.Alloc(output, GCHandleType.Pinned);
            try
            {
                var got = ResampleNative(inHandle.AddrOfPinnedObject(), (nuint)input.Length, (uint)srcRate, (uint)dstRate, outHandle.AddrOfPinnedObject(), (nuint)output.Length);
                if (got == 0)
                {
                    throw new InvalidOperationException(GetLastNativeError());
                }
                if ((int)got != output.Length)
                {
                    Array.Resize(ref output, (int)got);
                }
                return output;
            }
            finally
            {
                outHandle.Free();
            }
        }
        finally
        {
            inHandle.Free();
        }
    }

    private static float[,] ComputeMel(float[] input, QwenSplitConfig cfg)
    {
        var melCfg = new MelConfigNative
        {
            sample_rate = (uint)cfg.SpeakerSampleRate,
            n_fft = (nuint)cfg.SpeakerNfft,
            hop_length = (nuint)cfg.SpeakerHop,
            win_length = (nuint)cfg.SpeakerWin,
            n_mels = (nuint)cfg.SpeakerMels,
            fmin = cfg.SpeakerFMin,
            fmax = cfg.SpeakerFMax
        };
        var inHandle = GCHandle.Alloc(input, GCHandleType.Pinned);
        try
        {
            nuint rows = 0;
            nuint cols = 0;
            var needed = MelNative(inHandle.AddrOfPinnedObject(), (nuint)input.Length, ref melCfg, nint.Zero, 0, ref rows, ref cols);
            if (needed == 0)
            {
                throw new InvalidOperationException(GetLastNativeError());
            }

            var buffer = new float[(int)needed];
            var outHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            try
            {
                var got = MelNative(inHandle.AddrOfPinnedObject(), (nuint)input.Length, ref melCfg, outHandle.AddrOfPinnedObject(), (nuint)buffer.Length, ref rows, ref cols);
                if (got == 0)
                {
                    throw new InvalidOperationException(GetLastNativeError());
                }
                var r = (int)rows;
                var c = (int)cols;
                var mel = new float[r, c];
                var idx = 0;
                for (var i = 0; i < r; i++)
                {
                    for (var j = 0; j < c; j++)
                    {
                        mel[i, j] = buffer[idx++];
                    }
                }
                return mel;
            }
            finally
            {
                outHandle.Free();
            }
        }
        finally
        {
            inHandle.Free();
        }
    }

    private static long[] TokenizerEncode(nint handle, string text)
    {
        var needed = TokenizerEncodeNative(handle, text, nint.Zero, 0);
        if (needed == 0)
        {
            throw new InvalidOperationException(GetLastNativeError());
        }

        var output = new long[(int)needed];
        var gc = GCHandle.Alloc(output, GCHandleType.Pinned);
        try
        {
            var got = TokenizerEncodeNative(handle, text, gc.AddrOfPinnedObject(), (nuint)output.Length);
            if (got == 0)
            {
                throw new InvalidOperationException(GetLastNativeError());
            }
            if ((int)got != output.Length)
            {
                Array.Resize(ref output, (int)got);
            }
            return output;
        }
        finally
        {
            gc.Free();
        }
    }

    private static string GetLastNativeError()
    {
        var buffer = new byte[4096];
        var gc = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            _ = LastErrorMessage(gc.AddrOfPinnedObject(), (nuint)buffer.Length);
            var len = Array.IndexOf(buffer, (byte)0);
            if (len < 0)
            {
                len = buffer.Length;
            }
            return Encoding.UTF8.GetString(buffer, 0, len);
        }
        finally
        {
            gc.Free();
        }
    }
}
