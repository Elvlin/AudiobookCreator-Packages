using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using App.Core.Runtime;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace App.Inference;

public sealed class Qwen3OnnxDllBackend : ITtsBackend, IDisposable
{
    private static readonly object CudaBootstrapSync = new();
    private static bool _cudaBootstrapDone;

    private readonly LocalInferenceOptions _options;
    private readonly Random _rng = new();

    private InferenceSession? _speakerEncoder;
    private InferenceSession? _talkerPrefill;
    private InferenceSession? _codePredictor;
    private InferenceSession? _textProject;
    private InferenceSession? _codecEmbed;
    private InferenceSession? _codePredictorEmbed;
    private InferenceSession? _tokenizerDecode;
    private QwenConfig? _config;
    private string _loadedKey = string.Empty;

    public Qwen3OnnxDllBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "qwen3-onnx-dll-native";

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        return Task.Run(() => SynthesizeInternal(request, ct), ct);
    }

    public void Dispose()
    {
        _speakerEncoder?.Dispose();
        _talkerPrefill?.Dispose();
        _codePredictor?.Dispose();
        _textProject?.Dispose();
        _codecEmbed?.Dispose();
        _codePredictorEmbed?.Dispose();
        _tokenizerDecode?.Dispose();
        _speakerEncoder = null;
        _talkerPrefill = null;
        _codePredictor = null;
        _textProject = null;
        _codecEmbed = null;
        _codePredictorEmbed = null;
        _tokenizerDecode = null;
        _config = null;
        _loadedKey = string.Empty;
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
            ? "zukky/Qwen3-TTS-ONNX-DLL"
            : _options.ModelRepoId.Trim();
        var cacheDir = ModelCachePath.ResolveAbsolute(_options.ModelCacheDir, RuntimePathResolver.AppRoot);
        EnsureLoaded(cacheDir, repoId);
        if (_config is null || _speakerEncoder is null || _talkerPrefill is null || _codePredictor is null ||
            _textProject is null || _codecEmbed is null || _codePredictorEmbed is null || _tokenizerDecode is null)
        {
            throw new InvalidOperationException("Qwen backend failed to initialize.");
        }

        var (wav, sr) = ReadWav(request.VoicePath);
        if (sr != _config.SpeakerSampleRate)
        {
            wav = Resample(wav, sr, _config.SpeakerSampleRate);
            sr = _config.SpeakerSampleRate;
        }

        var mel = ComputeMel(wav, _config);
        var melFlat = TransposeMelToInput(mel);
        var speakerEmbed = RunSpeakerEncoder(_speakerEncoder, melFlat);

        var tokenizer = CreateTokenizer(_config.VocabPath, _config.MergesPath, _config.TokenizerConfigPath);
        try
        {
            // Do NOT put RefText in the prompt: the model would speak it (reference + beeps/repeats). RefText is for future ref_text input when supported.
            var prompt = BuildAssistantText(request.Text);
            var inputIds = TokenizerEncode(tokenizer, prompt);
            var audio = GenerateAudioWithRetries(
                request,
                inputIds,
                speakerEmbed,
                _config,
                _textProject,
                _codecEmbed,
                _talkerPrefill,
                _codePredictor,
                _codePredictorEmbed,
                _tokenizerDecode,
                ct);
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
        QwenConfig cfg,
        InferenceSession textProject,
        InferenceSession codecEmbed,
        InferenceSession talkerPrefill,
        InferenceSession codePredictor,
        InferenceSession codePredictorEmbed,
        InferenceSession tokenizerDecode,
        CancellationToken ct)
    {
        // Use one decode config for all chunks so voice/level stay consistent (no quality-guard retry with MakeSafer).
        var decode = QwenDecodeOptions.From(cfg, request);
        var state = BuildInputState(inputIds, speakerEmbed, cfg, textProject, codecEmbed);
        var codes = GenerateCodes(state, cfg, talkerPrefill, codePredictor, codecEmbed, codePredictorEmbed, decode, ct);
        if (codes.GetLength(0) == 0)
        {
            throw new InvalidOperationException("Qwen generation produced no audio codes.");
        }
        var audio = DecodeAudioCodes(tokenizerDecode, codes, cfg.DecodeUpsampleRate);
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

        var dllSource = Path.Combine(repoRoot, "qwen3_tts_rust.dll");
        if (!File.Exists(dllSource))
        {
            dllSource = Directory.GetFiles(repoRoot, "qwen3_tts_rust.dll", SearchOption.AllDirectories).FirstOrDefault() ?? string.Empty;
        }
        if (!File.Exists(dllSource))
        {
            dllSource = QwenEmbeddedRuntime.EnsureBundledRustDllAtAppRoot();
        }
        if (!File.Exists(dllSource))
        {
            throw new FileNotFoundException("qwen3_tts_rust.dll missing. Re-download Qwen ONNX DLL preset.", dllSource);
        }
        var appDll = Path.Combine(RuntimePathResolver.AppRoot, "qwen3_tts_rust.dll");
        if (!File.Exists(appDll) || new FileInfo(appDll).Length != new FileInfo(dllSource).Length)
        {
            File.Copy(dllSource, appDll, true);
        }

        var onnx06Dir = Path.Combine(repoRoot, "onnx_kv_06b");
        var onnx17Dir = Path.Combine(repoRoot, "onnx_kv");
        var onnxDir = Directory.Exists(onnx06Dir) ? onnx06Dir : onnx17Dir;

        var model06Dir = Path.Combine(repoRoot, "models", "Qwen3-TTS-12Hz-0.6B-Base");
        var model17Dir = Path.Combine(repoRoot, "models", "Qwen3-TTS-12Hz-1.7B-Base");
        var modelDir = Directory.Exists(model06Dir) ? model06Dir : model17Dir;
        if (!Directory.Exists(onnxDir) || !Directory.Exists(modelDir))
        {
            throw new InvalidOperationException("Qwen ONNX model layout not found in cache.");
        }

        if (string.Equals(Path.GetFileName(onnxDir), "onnx_kv", StringComparison.OrdinalIgnoreCase))
        {
            var prefillPath = Path.Combine(onnxDir, "talker_prefill.onnx");
            if (File.Exists(prefillPath) && new FileInfo(prefillPath).Length > int.MaxValue)
            {
                throw new InvalidOperationException(
                    "Qwen 1.7B ONNX set is not loadable by strict native ORT (talker_prefill.onnx > 2GB). " +
                    "Use onnx_kv_06b / Qwen3 0.6B ONNX DLL preset.");
            }
        }

        _config = LoadQwenConfig(modelDir);
        _config.OnnxDir = onnxDir;
        _config.VocabPath = Path.Combine(modelDir, "vocab.json");
        _config.MergesPath = Path.Combine(modelDir, "merges.txt");
        _config.TokenizerConfigPath = Path.Combine(modelDir, "tokenizer_config.json");

        var requiredOnnx = new[]
        {
            "speaker_encoder.onnx",
            "talker_prefill.onnx",
            "code_predictor.onnx",
            "text_project.onnx",
            "codec_embed.onnx",
            "code_predictor_embed.onnx",
            "tokenizer12hz_decode.onnx"
        };
        foreach (var file in requiredOnnx)
        {
            var path = Path.Combine(onnxDir, file);
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Required Qwen ONNX file missing: {file}. Re-download Qwen ONNX DLL preset.", path);
            }
        }

        var sessionOptions = CreateSessionOptions(_options.PreferDevice);
        _speakerEncoder = CreateSessionOrThrow(Path.Combine(onnxDir, "speaker_encoder.onnx"), sessionOptions);
        _talkerPrefill = CreateSessionOrThrow(Path.Combine(onnxDir, "talker_prefill.onnx"), sessionOptions);
        _codePredictor = CreateSessionOrThrow(Path.Combine(onnxDir, "code_predictor.onnx"), sessionOptions);
        _textProject = CreateSessionOrThrow(Path.Combine(onnxDir, "text_project.onnx"), sessionOptions);
        _codecEmbed = CreateSessionOrThrow(Path.Combine(onnxDir, "codec_embed.onnx"), sessionOptions);
        _codePredictorEmbed = CreateSessionOrThrow(Path.Combine(onnxDir, "code_predictor_embed.onnx"), sessionOptions);
        _tokenizerDecode = CreateSessionOrThrow(Path.Combine(onnxDir, "tokenizer12hz_decode.onnx"), sessionOptions);

        _loadedKey = key;
    }

    private static SessionOptions CreateSessionOptions(string preferDevice)
    {
        var options = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_EXTENDED };
        var normalized = (preferDevice ?? string.Empty).Trim().ToLowerInvariant();
        if (normalized == "cpu")
        {
            return options;
        }

        if (normalized == "dml")
        {
            if (!OperatingSystem.IsWindows())
            {
                throw new InvalidOperationException("DML provider is only available on Windows.");
            }

            options.AppendExecutionProvider_DML(0);
            return options;
        }

        if (normalized == "gpu")
        {
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
                catch (Exception dmlEx)
                {
                    throw new InvalidOperationException(
                        $"GPU was selected but CUDA/DML execution providers are unavailable. CUDA load error: {cudaError?.Message}; DML load error: {dmlEx.Message}");
                }
            }

            throw new InvalidOperationException(
                $"GPU was selected but CUDA execution provider is unavailable. CUDA load error: {cudaError?.Message}");
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
                // Fall back to CUDA below.
            }
        }

        TryBootstrapCudaRuntimePaths();
        Exception? cudaErrorFallback = null;
        try
        {
            options.AppendExecutionProvider_CUDA(0);
            return options;
        }
        catch (Exception ex)
        {
            cudaErrorFallback = ex;
        }

        return options;
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
                $"Qwen ONNX file is invalid or incomplete: {modelPath}. Remove cache and re-download Qwen ONNX DLL preset.",
                ex);
        }
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

    private static float[] RunSpeakerEncoder(InferenceSession session, float[] melInput)
    {
        var seq = melInput.Length / 128;
        var melTensor = new DenseTensor<float>(melInput, new[] { 1, seq, 128 });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("mels", melTensor) });
        return results.First().AsTensor<float>().ToArray();
    }

    private sealed class InputState
    {
        public float[] InputsEmbeds { get; set; } = Array.Empty<float>();
        public long[] AttentionMask { get; set; } = Array.Empty<long>();
        public float[] TrailingTextHidden { get; set; } = Array.Empty<float>();
        public float[] TtsPadEmbed { get; set; } = Array.Empty<float>();
    }

    private static InputState BuildInputState(long[] inputIds, float[] speakerEmbedding, QwenConfig cfg, InferenceSession textProject, InferenceSession codecEmbed)
    {
        if (inputIds.Length < 10)
        {
            throw new InvalidOperationException("Qwen tokenizer prompt is too short.");
        }

        var hidden = cfg.HiddenSize;

        var ttsEmbeds = RunTextProject(textProject, new[] { (long)cfg.TtsBosTokenId, (long)cfg.TtsEosTokenId, (long)cfg.TtsPadTokenId });
        var ttsBos = SliceSeq(ttsEmbeds, hidden, 0, 1);
        var ttsEos = SliceSeq(ttsEmbeds, hidden, 1, 1);
        var ttsPad = SliceSeq(ttsEmbeds, hidden, 2, 1);

        long[] codecPrefix;
        if (cfg.CodecLanguageId.TryGetValue("english", out var langId))
        {
            codecPrefix = new[] { cfg.CodecThinkId, cfg.CodecThinkBosId, langId, cfg.CodecThinkEosId };
        }
        else
        {
            codecPrefix = new[] { cfg.CodecNoThinkId, cfg.CodecThinkBosId, cfg.CodecThinkEosId };
        }

        var codecPrefixEmb = RunCodecEmbed(codecEmbed, codecPrefix);
        var codecPadBosEmb = RunCodecEmbed(codecEmbed, new[] { cfg.CodecPadId, cfg.CodecBosId });
        var speakerEmb3 = new float[hidden];
        Array.Copy(speakerEmbedding, speakerEmb3, Math.Min(hidden, speakerEmbedding.Length));

        var codecInputEmb = ConcatSeq(codecPrefixEmb, speakerEmb3, codecPadBosEmb);

        var roleEmbed = RunTextProject(textProject, SliceIds(inputIds, 0, 3));
        var padRepeat = Math.Max(0, SeqLen(codecInputEmb, hidden) - 2);
        var padBlock = RepeatSeq(ttsPad, hidden, padRepeat);
        var talkerEmb = AddSeq(
            ConcatSeq(padBlock, ttsBos),
            SliceSeq(codecInputEmb, hidden, 0, SeqLen(codecInputEmb, hidden) - 1));
        var talkerInput = ConcatSeq(roleEmbed, talkerEmb);

        var firstText = AddSeq(
            RunTextProject(textProject, SliceIds(inputIds, 3, 1)),
            SliceSeq(codecInputEmb, hidden, SeqLen(codecInputEmb, hidden) - 1, 1));
        talkerInput = ConcatSeq(talkerInput, firstText);

        var trailingText = RunTextProject(textProject, SliceIds(inputIds, 4, Math.Max(1, inputIds.Length - 9)));
        var trailingHidden = ConcatSeq(trailingText, ttsEos);

        return new InputState
        {
            InputsEmbeds = talkerInput,
            AttentionMask = Enumerable.Repeat(1L, SeqLen(talkerInput, hidden)).ToArray(),
            TrailingTextHidden = trailingHidden,
            TtsPadEmbed = ttsPad
        };
    }

    private long[,] GenerateCodes(InputState state, QwenConfig cfg, InferenceSession talkerPrefill, InferenceSession codePredictor, InferenceSession codecEmbed, InferenceSession codePredictorEmbed, QwenDecodeOptions decode, CancellationToken ct)
    {
        var hidden = cfg.HiddenSize;
        var seq = SeqLen(state.InputsEmbeds, hidden);
        var trailSeq = SeqLen(state.TrailingTextHidden, hidden);
        if (trailSeq <= 0)
            return new long[0, cfg.NumCodeGroups];

        var generated = new List<long[]>();
        var firstTokenHistory = new List<long>();
        var suppressStart = Math.Max(0, cfg.CodecVocabSize - 1024);

        for (var step = 0; ; step++)
        {
            ct.ThrowIfCancellationRequested();

            var (logits, lastHidden) = RunTalkerPrefill(talkerPrefill, state.InputsEmbeds, seq, hidden, state.AttentionMask);
            for (var tok = suppressStart; tok < cfg.CodecVocabSize; tok++)
            {
                if (tok == cfg.CodecEosTokenId)
                    continue;
                logits[tok] = -1.0e9f;
            }

            ApplyRepetitionPenalty(logits, firstTokenHistory, decode.RepetitionPenalty);
            var firstCode = SelectToken(logits, decode);
            firstTokenHistory.Add(firstCode);

            var row = new long[cfg.NumCodeGroups];
            row[0] = firstCode;

            var firstEmbed = RunCodecEmbed(codecEmbed, new[] { firstCode });
            var sum = firstEmbed;
            var predictorInput = ConcatSeq(lastHidden, firstEmbed);
            for (var g = 0; g < cfg.NumCodeGroups - 1; g++)
            {
                var subLogits = RunCodePredictor(codePredictor, predictorInput, hidden, g);
                var subCode = SelectToken(subLogits, decode);
                row[g + 1] = subCode;
                var subEmb = RunCodePredictorEmbed(codePredictorEmbed, subCode, g);
                sum = AddSeq(sum, subEmb);
                predictorInput = ConcatSeq(predictorInput, subEmb);
            }

            generated.Add(row);
            if (firstCode == cfg.CodecEosTokenId)
                break;
            if (step >= trailSeq)
                break;

            var trail = SliceSeq(state.TrailingTextHidden, hidden, step, 1);
            var codecStep = AddSeq(sum, trail);
            state.InputsEmbeds = ConcatSeq(state.InputsEmbeds, codecStep);
            state.AttentionMask = AppendMask(state.AttentionMask);
            seq++;
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

    private static float[] DecodeAudioCodes(InferenceSession tokenizerDecode, long[,] codes, int decodeUpsampleRate)
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
        using var results = tokenizerDecode.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("audio_codes", codeTensor)
        });
        var list = results.ToList();
        var audioTensor = list[0].AsTensor<float>();
        var audio = audioTensor.ToArray();
        var audioLen = audioTensor.Dimensions[^1];
        var targetLen = Math.Min(audioLen, Math.Max(1, steps * decodeUpsampleRate));
        if (list.Count > 1)
        {
            var lenTensor = list[1].AsTensor<long>();
            var reported = (int)Math.Max(0, lenTensor.ToArray().FirstOrDefault());
            if (reported > 0)
            {
                targetLen = Math.Min(targetLen, reported);
            }
        }

        var output = new float[targetLen];
        Array.Copy(audio, output, targetLen);
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

    private sealed class QwenConfig
    {
        public int TtsBosTokenId { get; init; }
        public int TtsEosTokenId { get; init; }
        public int TtsPadTokenId { get; init; }
        public int HiddenSize { get; init; }
        public int NumCodeGroups { get; init; }
        public int CodecVocabSize { get; init; }
        public long CodecBosId { get; init; }
        public long CodecEosTokenId { get; init; }
        public long CodecPadId { get; init; }
        public long CodecThinkId { get; init; }
        public long CodecNoThinkId { get; init; }
        public long CodecThinkBosId { get; init; }
        public long CodecThinkEosId { get; init; }
        public int SpeakerSampleRate { get; init; } = 24000;
        public int SpeakerNfft { get; init; } = 1024;
        public int SpeakerHop { get; init; } = 256;
        public int SpeakerWin { get; init; } = 1024;
        public int SpeakerMels { get; init; } = 128;
        public float SpeakerFMin { get; init; } = 0;
        public float SpeakerFMax { get; init; } = 12000;
        public int DecodeUpsampleRate { get; init; } = 1920;
        public int OutputSampleRate { get; init; } = 24000;
        public Dictionary<string, long> CodecLanguageId { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public bool CodePredictorDoSample { get; init; } = false;
        public int CodePredictorTopK { get; init; } = 50;
        public float CodePredictorTopP { get; init; } = 1.0f;
        public float CodePredictorTemperature { get; init; } = 1.0f;
        public float CodePredictorRepetitionPenalty { get; init; } = 1.0f;
        public string OnnxDir { get; set; } = string.Empty;
        public string VocabPath { get; set; } = string.Empty;
        public string MergesPath { get; set; } = string.Empty;
        public string TokenizerConfigPath { get; set; } = string.Empty;
    }

    private static QwenConfig LoadQwenConfig(string modelDir)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(modelDir, "config.json")));
        var root = doc.RootElement;
        var talker = root.GetProperty("talker_config");
        var speaker = root.GetProperty("speaker_encoder_config");
        var codePredictorCfg = talker.TryGetProperty("code_predictor_config", out var cp) ? cp : default;

        var lang = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        if (talker.TryGetProperty("codec_language_id", out var langObj))
        {
            foreach (var p in langObj.EnumerateObject())
            {
                if (p.Value.TryGetInt64(out var v))
                {
                    lang[p.Name] = v;
                }
            }
        }

        return new QwenConfig
        {
            TtsBosTokenId = root.GetProperty("tts_bos_token_id").GetInt32(),
            TtsEosTokenId = root.GetProperty("tts_eos_token_id").GetInt32(),
            TtsPadTokenId = root.GetProperty("tts_pad_token_id").GetInt32(),
            HiddenSize = talker.GetProperty("hidden_size").GetInt32(),
            NumCodeGroups = talker.GetProperty("num_code_groups").GetInt32(),
            CodecVocabSize = talker.GetProperty("vocab_size").GetInt32(),
            CodecBosId = talker.GetProperty("codec_bos_id").GetInt64(),
            CodecEosTokenId = talker.GetProperty("codec_eos_token_id").GetInt64(),
            CodecPadId = talker.GetProperty("codec_pad_id").GetInt64(),
            CodecThinkId = talker.GetProperty("codec_think_id").GetInt64(),
            CodecNoThinkId = talker.GetProperty("codec_nothink_id").GetInt64(),
            CodecThinkBosId = talker.GetProperty("codec_think_bos_id").GetInt64(),
            CodecThinkEosId = talker.GetProperty("codec_think_eos_id").GetInt64(),
            SpeakerSampleRate = speaker.TryGetProperty("sample_rate", out var sr) ? sr.GetInt32() : 24000,
            SpeakerNfft = speaker.TryGetProperty("n_fft", out var nfft) ? nfft.GetInt32() : 1024,
            SpeakerHop = speaker.TryGetProperty("hop_size", out var hop) ? hop.GetInt32() : 256,
            SpeakerWin = speaker.TryGetProperty("win_size", out var win) ? win.GetInt32() : 1024,
            SpeakerMels = speaker.TryGetProperty("num_mels", out var nmels) ? nmels.GetInt32() : 128,
            SpeakerFMin = speaker.TryGetProperty("fmin", out var fmin) ? (float)fmin.GetDouble() : 0f,
            SpeakerFMax = speaker.TryGetProperty("fmax", out var fmax) ? (float)fmax.GetDouble() : 12000f,
            CodecLanguageId = lang,
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
    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_build_assistant_text")]
    private static extern nuint BuildAssistantTextNative([MarshalAs(UnmanagedType.LPUTF8Str)] string text, nint outBuf, nuint outLen);

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

    private static float[,] ComputeMel(float[] input, QwenConfig cfg)
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

    private static string BuildAssistantText(string text)
    {
        var cap = Math.Max(4096, text.Length * 4 + 128);
        var buffer = new byte[cap];
        var handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            _ = BuildAssistantTextNative(text, handle.AddrOfPinnedObject(), (nuint)buffer.Length);
            var len = Array.IndexOf(buffer, (byte)0);
            if (len < 0)
            {
                len = buffer.Length;
            }
            return Encoding.UTF8.GetString(buffer, 0, len);
        }
        finally
        {
            handle.Free();
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
            (rms < 0.012 && zcr > 0.22) ||
            flatRatio > 0.96;

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
        public static QwenDecodeOptions From(QwenConfig cfg, TtsRequest request)
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
            if (attempt >= 1)
            {
                return this with
                {
                    DoSample = false,
                    TopK = 1,
                    TopP = 1.0f,
                    Temperature = 0.2f,
                    RepetitionPenalty = Math.Clamp(Math.Max(RepetitionPenalty, 1.05f), 1.0f, 2.0f)
                };
            }

            return this with
            {
                DoSample = true,
                TopK = Math.Clamp(Math.Min(TopK, 16), 1, 200),
                TopP = Math.Clamp(Math.Min(TopP, 0.92f), 0.05f, 1.0f),
                Temperature = Math.Clamp(Math.Min(Temperature, 0.65f), 0.05f, 2.0f),
                RepetitionPenalty = Math.Clamp(Math.Max(RepetitionPenalty, 1.03f), 1.0f, 2.0f)
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

    private static (float[] Logits, float[] LastHidden) RunTalkerPrefill(InferenceSession session, float[] embeds, int seq, int hidden, long[] mask)
    {
        var embedTensor = new DenseTensor<float>(embeds, new[] { 1, seq, hidden });
        var maskTensor = new DenseTensor<long>(mask, new[] { 1, mask.Length });
        using var results = session.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("inputs_embeds", embedTensor),
            NamedOnnxValue.CreateFromTensor("attention_mask", maskTensor)
        });
        var list = results.ToList();
        var logitsTensor = list[0].AsTensor<float>();
        var hiddenTensor = list[1].AsTensor<float>();
        var logitsAll = logitsTensor.ToArray();
        var vocab = logitsTensor.Dimensions[^1];
        var outSeq = logitsTensor.Dimensions[^2];
        var logits = new float[vocab];
        Array.Copy(logitsAll, (outSeq - 1) * vocab, logits, 0, vocab);
        return (logits, hiddenTensor.ToArray());
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
        return results.First().AsTensor<float>().ToArray();
    }

    private static float[] RunCodePredictorEmbed(InferenceSession session, long codeId, int generationStep)
    {
        var idTensor = new DenseTensor<long>(new[] { codeId }, new[] { 1, 1 });
        var stepTensor = new DenseTensor<long>(new[] { (long)generationStep }, new[] { 1 });
        using var results = session.Run(new[]
        {
            NamedOnnxValue.CreateFromTensor("input_ids", idTensor),
            NamedOnnxValue.CreateFromTensor("generation_step", stepTensor)
        });
        return results.First().AsTensor<float>().ToArray();
    }

    private static float[] RunTextProject(InferenceSession session, long[] ids)
    {
        var tensor = new DenseTensor<long>(ids, new[] { 1, ids.Length });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("input_ids", tensor) });
        return results.First().AsTensor<float>().ToArray();
    }

    private static float[] RunCodecEmbed(InferenceSession session, long[] ids)
    {
        var tensor = new DenseTensor<long>(ids, new[] { 1, ids.Length });
        using var results = session.Run(new[] { NamedOnnxValue.CreateFromTensor("input_ids", tensor) });
        return results.First().AsTensor<float>().ToArray();
    }

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
            if (x.Length == 0) continue;
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
        var outArr = new float[a.Length];
        for (var i = 0; i < a.Length; i++)
        {
            outArr[i] = a[i] + b[i];
        }
        return outArr;
    }

    private static long[] AppendMask(long[] mask)
    {
        var outMask = new long[mask.Length + 1];
        Array.Copy(mask, outMask, mask.Length);
        outMask[^1] = 1;
        return outMask;
    }
}
