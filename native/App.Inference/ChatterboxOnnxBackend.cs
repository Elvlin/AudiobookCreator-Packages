using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using App.Core.Runtime;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace App.Inference;

public sealed class ChatterboxOnnxBackend : ITtsBackend, IDisposable
{
    private const int SampleRate = 24000;
    private const int HiddenSize = 1024;
    private const int NumHiddenLayers = 30;
    private const int NumKeyValueHeads = 16;
    private const int HeadDim = 64;
    private const int StartSpeechToken = 6561;
    private const int StopSpeechToken = 6562;
    private const int DefaultMaxNewTokens = 512;
    private const float DefaultExaggeration = 0.5f;
    private const float DefaultRepetitionPenalty = 1.2f;

    private static readonly string[] RequiredModelFiles =
    {
        "onnx/conditional_decoder.onnx",
        "onnx/conditional_decoder.onnx_data",
        "onnx/embed_tokens.onnx",
        "onnx/embed_tokens.onnx_data",
        "onnx/language_model.onnx",
        "onnx/language_model.onnx_data",
        "onnx/speech_encoder.onnx",
        "onnx/speech_encoder.onnx_data",
        "tokenizer.json"
    };

    private readonly LocalInferenceOptions _options;
    private readonly object _sync = new();

    private string _loadedKey = string.Empty;
    private InferenceSession? _speechEncoder;
    private InferenceSession? _embedTokens;
    private InferenceSession? _languageModel;
    private InferenceSession? _conditionalDecoder;
    private ChatterboxTokenizer? _tokenizer;
    private float _repetitionPenalty = DefaultRepetitionPenalty;
    private static readonly object CudaBootstrapSync = new();
    private static bool _cudaBootstrapDone;

    public ChatterboxOnnxBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "chatterbox-onnx-native";
    public string ActiveExecutionProvider { get; private set; } = "unloaded";

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        return Task.Run(() => SynthesizeInternal(request, ct), ct);
    }

    public void Dispose()
    {
        lock (_sync)
        {
            UnloadLocked();
        }

        GC.SuppressFinalize(this);
    }

    private void SynthesizeInternal(TtsRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            throw new ArgumentException("Input text is empty.");
        }
        if (string.IsNullOrWhiteSpace(request.VoicePath))
        {
            throw new ArgumentException("Voice path is required.");
        }
        if (!File.Exists(request.VoicePath))
        {
            throw new FileNotFoundException($"Voice file not found: {request.VoicePath}");
        }
        if (string.IsNullOrWhiteSpace(request.OutputPath))
        {
            throw new ArgumentException("Output path is required.");
        }

        var outputDir = Path.GetDirectoryName(request.OutputPath);
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var modelRepoId = string.IsNullOrWhiteSpace(_options.ModelRepoId)
            ? "onnx-community/chatterbox-ONNX"
            : _options.ModelRepoId.Trim();
        if (!modelRepoId.Contains("chatterbox", StringComparison.OrdinalIgnoreCase) ||
            !modelRepoId.Contains("onnx", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Selected local model is not a Chatterbox ONNX model. Full native inference currently supports chatterbox-ONNX only.");
        }

        var modelCacheDir = ModelCachePath.ResolveAbsolute(_options.ModelCacheDir, RuntimePathResolver.AppRoot);
        EnsureLoaded(modelCacheDir, modelRepoId);

        var voice = ReadWavAsMonoFloat(request.VoicePath);
        if (voice.SampleRate != SampleRate)
        {
            voice = voice with { Samples = ResampleLinear(voice.Samples, voice.SampleRate, SampleRate), SampleRate = SampleRate };
        }

        long[] modelInputIds;
        lock (_sync)
        {
            modelInputIds = _tokenizer!.EncodeForModel(request.Text).ToArray();
        }

        ct.ThrowIfCancellationRequested();

        float[] waveform;
        lock (_sync)
        {
            waveform = RunInferenceLocked(modelInputIds, voice.Samples, ct, request.ChatterboxExaggeration);
        }

        if (Math.Abs(request.Speed - 1.0f) > 0.01f)
        {
            var speed = Math.Clamp(request.Speed, 0.55f, 1.75f);
            waveform = ApplySpeedResample(waveform, speed);
        }

        WriteWav16Mono(request.OutputPath, waveform, SampleRate);
    }

    private void EnsureLoaded(string modelCacheDir, string modelRepoId)
    {
        var packRoot = ResolvePackRoot(modelCacheDir, modelRepoId);
        var device = NormalizePreferDevice(_options.PreferDevice);
        var key = $"{packRoot.ToLowerInvariant()}|{device}";

        lock (_sync)
        {
            if (_loadedKey == key &&
                _speechEncoder is not null &&
                _embedTokens is not null &&
                _languageModel is not null &&
                _conditionalDecoder is not null &&
                _tokenizer is not null)
            {
                return;
            }

            ValidateModelPack(packRoot);
            UnloadLocked();

            var loaded = false;
            Exception? firstError = null;
            var loadedProvider = "cpu";

            using (var preferredOptions = BuildSessionOptions(device, out var preferredProvider))
            {
                loaded = TryLoadSessions(packRoot, preferredOptions, out firstError);
                if (loaded)
                {
                    loadedProvider = preferredProvider;
                }
            }

            if (!loaded && device == "auto")
            {
                using var cpuOptions = BuildSessionOptions("cpu", out var cpuProvider);
                if (!TryLoadSessions(packRoot, cpuOptions, out var cpuError))
                {
                    throw new InvalidOperationException(
                        $"Failed to load Chatterbox ONNX on preferred device and CPU fallback. Preferred error: {firstError?.Message}. CPU error: {cpuError?.Message}");
                }
                loadedProvider = cpuProvider;
            }
            else if (!loaded)
            {
                throw new InvalidOperationException($"Failed to load Chatterbox ONNX model sessions: {firstError?.Message}", firstError);
            }

            _tokenizer = ChatterboxTokenizer.Load(Path.Combine(packRoot, "tokenizer.json"));
            _repetitionPenalty = ReadRepetitionPenalty(packRoot);
            ActiveExecutionProvider = loadedProvider;
            _loadedKey = key;
        }
    }

    private bool TryLoadSessions(string packRoot, SessionOptions options, out Exception? error)
    {
        InferenceSession? speech = null;
        InferenceSession? embed = null;
        InferenceSession? language = null;
        InferenceSession? decoder = null;

        try
        {
            speech = new InferenceSession(Path.Combine(packRoot, "onnx", "speech_encoder.onnx"), options);
            embed = new InferenceSession(Path.Combine(packRoot, "onnx", "embed_tokens.onnx"), options);
            language = new InferenceSession(Path.Combine(packRoot, "onnx", "language_model.onnx"), options);
            decoder = new InferenceSession(Path.Combine(packRoot, "onnx", "conditional_decoder.onnx"), options);

            _speechEncoder = speech;
            _embedTokens = embed;
            _languageModel = language;
            _conditionalDecoder = decoder;
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            speech?.Dispose();
            embed?.Dispose();
            language?.Dispose();
            decoder?.Dispose();
            error = ex;
            return false;
        }
    }

    private float[] RunInferenceLocked(long[] inputIds, float[] voiceSamples, CancellationToken ct, float? exaggerationOverride = null)
    {
        if (_speechEncoder is null || _embedTokens is null || _languageModel is null || _conditionalDecoder is null || _tokenizer is null)
        {
            throw new InvalidOperationException("Chatterbox ONNX sessions are not loaded.");
        }

        var exaggeration = exaggerationOverride.HasValue
            ? Math.Clamp(exaggerationOverride.Value, 0.0f, 1.0f)
            : (_options.Exaggeration <= 0 ? DefaultExaggeration : Math.Clamp(_options.Exaggeration, 0.0f, 1.0f));
        var maxNewTokens = _options.MaxNewTokens <= 0 ? DefaultMaxNewTokens : Math.Clamp(_options.MaxNewTokens, 64, 2048);

        var positionIds = BuildInitialPositionIds(inputIds);
        var textEmbeds = RunEmbedTokens(_embedTokens, inputIds, positionIds, exaggeration);
        var speechEncoderOut = RunSpeechEncoder(_speechEncoder, voiceSamples);

        var currentEmbeds = ConcatEmbeddings(speechEncoderOut.CondEmbeddings, speechEncoderOut.CondSeqLen, textEmbeds.Values, textEmbeds.SeqLen);
        var currentSeqLen = speechEncoderOut.CondSeqLen + textEmbeds.SeqLen;
        var attentionMask = Enumerable.Repeat(1L, currentSeqLen).ToArray();
        var pastKv = InitPastKeyValues();
        var generated = new List<long>(Math.Min(maxNewTokens + 1, 4096)) { StartSpeechToken };

        for (var step = 0; step < maxNewTokens; step++)
        {
            ct.ThrowIfCancellationRequested();

            var lmOut = RunLanguageModel(_languageModel, currentEmbeds, currentSeqLen, attentionMask, pastKv);
            pastKv = lmOut.PresentKeyValues;

            var nextToken = SelectNextToken(lmOut.LogitsLastStep, generated, _repetitionPenalty);
            generated.Add(nextToken);

            if (nextToken == StopSpeechToken)
            {
                break;
            }

            var nextIds = new[] { nextToken };
            var nextPositions = new[] { step + 1L };
            var nextEmbed = RunEmbedTokens(_embedTokens, nextIds, nextPositions, exaggeration);
            currentEmbeds = nextEmbed.Values;
            currentSeqLen = nextEmbed.SeqLen;
            attentionMask = AppendOne(attentionMask);
        }

        var speechTail = generated.Skip(1).ToList();
        if (speechTail.Count > 0 && speechTail[^1] == StopSpeechToken)
        {
            speechTail.RemoveAt(speechTail.Count - 1);
        }

        var speechTokens = new long[speechEncoderOut.PromptTokens.Length + speechTail.Count];
        Array.Copy(speechEncoderOut.PromptTokens, 0, speechTokens, 0, speechEncoderOut.PromptTokens.Length);
        for (var i = 0; i < speechTail.Count; i++)
        {
            speechTokens[speechEncoderOut.PromptTokens.Length + i] = speechTail[i];
        }

        return RunConditionalDecoder(
            _conditionalDecoder,
            speechTokens,
            speechEncoderOut.SpeakerEmbeddings,
            speechEncoderOut.SpeakerFeatures,
            speechEncoderOut.SpeakerFeatureSeqLen);
    }

    private static EmbedOut RunEmbedTokens(InferenceSession session, long[] inputIds, long[] positionIds, float exaggeration)
    {
        var inputIdsTensor = CreateTensor(inputIds, 1, inputIds.Length);
        var positionIdsTensor = CreateTensor(positionIds, 1, positionIds.Length);
        var exaggerationTensor = CreateTensor(new[] { exaggeration }, 1);

        var inputs = new List<NamedOnnxValue>(3)
        {
            NamedOnnxValue.CreateFromTensor("input_ids", inputIdsTensor),
            NamedOnnxValue.CreateFromTensor("position_ids", positionIdsTensor),
            NamedOnnxValue.CreateFromTensor("exaggeration", exaggerationTensor)
        };

        using var results = session.Run(inputs);
        var embeds = results.First(r => r.Name == "inputs_embeds").AsTensor<float>();
        return new EmbedOut(embeds.ToArray(), embeds.Dimensions[1]);
    }

    private static SpeechEncoderOut RunSpeechEncoder(InferenceSession session, float[] voiceSamples)
    {
        var audioTensor = CreateTensor(voiceSamples, 1, voiceSamples.Length);
        var inputs = new List<NamedOnnxValue>(1)
        {
            NamedOnnxValue.CreateFromTensor("audio_values", audioTensor)
        };

        using var results = session.Run(inputs);
        var condEmb = results.First(r => r.Name == "audio_features").AsTensor<float>();
        var promptTokens = results.First(r => r.Name == "audio_tokens").AsTensor<long>();
        var speakerEmb = results.First(r => r.Name == "speaker_embeddings").AsTensor<float>();
        var speakerFeat = results.First(r => r.Name == "speaker_features").AsTensor<float>();

        return new SpeechEncoderOut(
            condEmb.ToArray(),
            condEmb.Dimensions[1],
            promptTokens.ToArray(),
            speakerEmb.ToArray(),
            speakerFeat.ToArray(),
            speakerFeat.Dimensions[1]);
    }

    private static LanguageModelOut RunLanguageModel(
        InferenceSession session,
        float[] inputsEmbeds,
        int seqLen,
        long[] attentionMask,
        Dictionary<string, KvTensorData> pastKv)
    {
        var embedsTensor = CreateTensor(inputsEmbeds, 1, seqLen, HiddenSize);
        var attentionTensor = CreateTensor(attentionMask, 1, attentionMask.Length);

        var inputs = new List<NamedOnnxValue>(2 + pastKv.Count)
        {
            NamedOnnxValue.CreateFromTensor("inputs_embeds", embedsTensor),
            NamedOnnxValue.CreateFromTensor("attention_mask", attentionTensor)
        };

        foreach (var kv in pastKv)
        {
            var tensor = CreateTensor(kv.Value.Values, 1, NumKeyValueHeads, kv.Value.SeqLen, HeadDim);
            inputs.Add(NamedOnnxValue.CreateFromTensor(kv.Key, tensor));
        }

        using var results = session.Run(inputs);
        var logitsTensor = results.First(r => r.Name == "logits").AsTensor<float>();
        var logits = logitsTensor.ToArray();
        var seqOut = logitsTensor.Dimensions[1];
        var vocab = logitsTensor.Dimensions[2];
        var lastStep = new float[vocab];
        Array.Copy(logits, (seqOut - 1) * vocab, lastStep, 0, vocab);

        var present = new Dictionary<string, KvTensorData>(pastKv.Count, StringComparer.Ordinal);
        foreach (var output in results)
        {
            if (!output.Name.StartsWith("present.", StringComparison.Ordinal))
            {
                continue;
            }

            var tensor = output.AsTensor<float>();
            var key = "past_key_values." + output.Name["present.".Length..];
            present[key] = new KvTensorData(tensor.ToArray(), tensor.Dimensions[2]);
        }

        return new LanguageModelOut(lastStep, present);
    }

    private static float[] RunConditionalDecoder(
        InferenceSession session,
        long[] speechTokens,
        float[] speakerEmbeddings,
        float[] speakerFeatures,
        int speakerFeatureSeqLen)
    {
        var speechTensor = CreateTensor(speechTokens, 1, speechTokens.Length);
        var embeddingsTensor = CreateTensor(speakerEmbeddings, 1, 192);
        var featuresTensor = CreateTensor(speakerFeatures, 1, speakerFeatureSeqLen, 80);

        var inputs = new List<NamedOnnxValue>(3)
        {
            NamedOnnxValue.CreateFromTensor("speech_tokens", speechTensor),
            NamedOnnxValue.CreateFromTensor("speaker_embeddings", embeddingsTensor),
            NamedOnnxValue.CreateFromTensor("speaker_features", featuresTensor)
        };

        using var results = session.Run(inputs);
        var waveform = results.First(r => r.Name == "waveform").AsTensor<float>();
        return waveform.ToArray();
    }

    private static long SelectNextToken(float[] logitsLastStep, List<long> generatedTokens, float repetitionPenalty)
    {
        if (repetitionPenalty > 1.0001f)
        {
            foreach (var token in generatedTokens)
            {
                var idx = (int)token;
                if (idx < 0 || idx >= logitsLastStep.Length)
                {
                    continue;
                }

                var score = logitsLastStep[idx];
                logitsLastStep[idx] = score < 0 ? score * repetitionPenalty : score / repetitionPenalty;
            }
        }

        var bestIdx = 0;
        var bestValue = logitsLastStep[0];
        for (var i = 1; i < logitsLastStep.Length; i++)
        {
            if (logitsLastStep[i] > bestValue)
            {
                bestValue = logitsLastStep[i];
                bestIdx = i;
            }
        }

        return bestIdx;
    }

    private static Dictionary<string, KvTensorData> InitPastKeyValues()
    {
        var result = new Dictionary<string, KvTensorData>(NumHiddenLayers * 2, StringComparer.Ordinal);
        for (var layer = 0; layer < NumHiddenLayers; layer++)
        {
            result[$"past_key_values.{layer}.key"] = new KvTensorData(Array.Empty<float>(), 0);
            result[$"past_key_values.{layer}.value"] = new KvTensorData(Array.Empty<float>(), 0);
        }

        return result;
    }

    private static long[] BuildInitialPositionIds(long[] inputIds)
    {
        var output = new long[inputIds.Length];
        for (var i = 0; i < inputIds.Length; i++)
        {
            output[i] = inputIds[i] >= StartSpeechToken ? 0 : i - 1;
        }

        return output;
    }

    private static float[] ConcatEmbeddings(float[] left, int leftSeqLen, float[] right, int rightSeqLen)
    {
        var output = new float[(leftSeqLen + rightSeqLen) * HiddenSize];
        Buffer.BlockCopy(left, 0, output, 0, left.Length * sizeof(float));
        Buffer.BlockCopy(right, 0, output, left.Length * sizeof(float), right.Length * sizeof(float));
        return output;
    }

    private static long[] AppendOne(long[] source)
    {
        var output = new long[source.Length + 1];
        Array.Copy(source, output, source.Length);
        output[^1] = 1;
        return output;
    }

    private static SessionOptions BuildSessionOptions(string device, out string provider)
    {
        var options = new SessionOptions
        {
            GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_EXTENDED
        };

        if (device == "cpu")
        {
            provider = "cpu";
            return options;
        }

        TryBootstrapCudaRuntimePaths();
        Exception? cudaError = null;
        try
        {
            options.AppendExecutionProvider_CUDA(0);
            provider = "cuda";
            return options;
        }
        catch (Exception ex)
        {
            cudaError = ex;
            // Try DML.
        }

        if (OperatingSystem.IsWindows())
        {
            try
            {
                options.AppendExecutionProvider_DML(0);
                provider = "dml";
                return options;
            }
            catch
            {
                // Ignore and fall back.
            }
        }

        if (device == "gpu")
        {
            throw new InvalidOperationException(
                $"GPU was selected but CUDA/DML execution providers are unavailable. CUDA load error: {cudaError?.Message}");
        }

        provider = "cpu";
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

    private static void ValidateModelPack(string packRoot)
    {
        foreach (var file in RequiredModelFiles)
        {
            var path = Path.Combine(packRoot, file.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"ONNX model incomplete. Missing file: {file}", path);
            }
        }
    }

    private static string ResolvePackRoot(string modelCacheDir, string repoId)
    {
        var repoFolder = "models--" + repoId.Replace("/", "--");
        return Path.Combine(modelCacheDir, "hf-cache", repoFolder);
    }

    private static string NormalizePreferDevice(string? value)
    {
        return (value ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "gpu" => "gpu",
            "cuda" => "gpu",
            "cpu" => "cpu",
            _ => "auto"
        };
    }

    private static float ReadRepetitionPenalty(string packRoot)
    {
        var path = Path.Combine(packRoot, "generation_config.json");
        if (!File.Exists(path))
        {
            return DefaultRepetitionPenalty;
        }

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(path));
            if (doc.RootElement.TryGetProperty("repetition_penalty", out var value) && value.TryGetSingle(out var penalty))
            {
                return Math.Clamp(penalty, 1.0f, 2.0f);
            }
        }
        catch
        {
            // Keep default.
        }

        return DefaultRepetitionPenalty;
    }

    private void UnloadLocked()
    {
        _speechEncoder?.Dispose();
        _embedTokens?.Dispose();
        _languageModel?.Dispose();
        _conditionalDecoder?.Dispose();

        _speechEncoder = null;
        _embedTokens = null;
        _languageModel = null;
        _conditionalDecoder = null;
        _tokenizer = null;
        _repetitionPenalty = DefaultRepetitionPenalty;
        ActiveExecutionProvider = "unloaded";
        _loadedKey = string.Empty;
    }

    private static DenseTensor<long> CreateTensor(long[] data, params int[] dims)
    {
        var tensor = new DenseTensor<long>(dims);
        data.CopyTo(tensor.Buffer.Span);
        return tensor;
    }

    private static DenseTensor<float> CreateTensor(float[] data, params int[] dims)
    {
        var tensor = new DenseTensor<float>(dims);
        data.CopyTo(tensor.Buffer.Span);
        return tensor;
    }

    private static WavData ReadWavAsMonoFloat(string path)
    {
        using var fs = File.OpenRead(path);
        using var br = new BinaryReader(fs);

        if (Encoding.ASCII.GetString(br.ReadBytes(4)) != "RIFF")
        {
            throw new InvalidOperationException("Voice file is not RIFF WAV.");
        }

        _ = br.ReadUInt32();
        if (Encoding.ASCII.GetString(br.ReadBytes(4)) != "WAVE")
        {
            throw new InvalidOperationException("Voice file is not WAVE.");
        }

        ushort channels = 0;
        uint sampleRate = 0;
        ushort bitsPerSample = 0;
        byte[]? pcm = null;

        while (fs.Position + 8 <= fs.Length)
        {
            var chunkId = Encoding.ASCII.GetString(br.ReadBytes(4));
            var chunkSize = br.ReadUInt32();
            if (chunkSize > int.MaxValue || fs.Position + chunkSize > fs.Length)
            {
                throw new InvalidOperationException("Invalid WAV chunk.");
            }

            if (chunkId == "fmt ")
            {
                var audioFormat = br.ReadUInt16();
                channels = br.ReadUInt16();
                sampleRate = br.ReadUInt32();
                _ = br.ReadUInt32();
                _ = br.ReadUInt16();
                bitsPerSample = br.ReadUInt16();

                var extra = (int)chunkSize - 16;
                if (extra > 0)
                {
                    br.ReadBytes(extra);
                }

                if (audioFormat != 1)
                {
                    throw new InvalidOperationException("Voice WAV must be PCM.");
                }
            }
            else if (chunkId == "data")
            {
                pcm = br.ReadBytes((int)chunkSize);
            }
            else
            {
                br.ReadBytes((int)chunkSize);
            }

            if ((chunkSize & 1) == 1 && fs.Position < fs.Length)
            {
                fs.Position++;
            }
        }

        if (pcm is null || pcm.Length < 2)
        {
            throw new InvalidOperationException("Voice WAV has no audio data.");
        }
        if (channels is not (1 or 2))
        {
            throw new InvalidOperationException("Voice WAV must be mono or stereo.");
        }
        if (bitsPerSample != 16)
        {
            throw new InvalidOperationException("Voice WAV must be 16-bit PCM.");
        }
        if (sampleRate < 8000)
        {
            throw new InvalidOperationException("Voice WAV sample rate is too low.");
        }

        var samples = new float[pcm.Length / 2 / channels];
        var src = 0;
        var dst = 0;
        while (src + (2 * channels) <= pcm.Length)
        {
            short left = BitConverter.ToInt16(pcm, src);
            short value = left;
            if (channels == 2)
            {
                short right = BitConverter.ToInt16(pcm, src + 2);
                value = (short)((left + right) / 2);
            }

            samples[dst++] = value / 32768.0f;
            src += 2 * channels;
        }

        return new WavData(samples, (int)sampleRate);
    }

    private static float[] ResampleLinear(float[] input, int fromSampleRate, int toSampleRate)
    {
        if (input.Length == 0 || fromSampleRate <= 0 || toSampleRate <= 0 || fromSampleRate == toSampleRate)
        {
            return input;
        }

        var ratio = (double)toSampleRate / fromSampleRate;
        var outputLength = Math.Max(1, (int)Math.Round(input.Length * ratio));
        var output = new float[outputLength];

        for (var i = 0; i < outputLength; i++)
        {
            var sourceIndex = i / ratio;
            var lo = (int)Math.Floor(sourceIndex);
            var hi = Math.Min(lo + 1, input.Length - 1);
            var frac = sourceIndex - lo;
            output[i] = (float)((input[lo] * (1.0 - frac)) + (input[hi] * frac));
        }

        return output;
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

        for (var i = 0; i < samples.Length; i++)
        {
            var clamped = Math.Clamp(samples[i], -1.0f, 1.0f);
            var pcm = (short)Math.Round(clamped * 32767.0f);
            bw.Write(pcm);
        }
    }

    private readonly record struct EmbedOut(float[] Values, int SeqLen);

    private readonly record struct SpeechEncoderOut(
        float[] CondEmbeddings,
        int CondSeqLen,
        long[] PromptTokens,
        float[] SpeakerEmbeddings,
        float[] SpeakerFeatures,
        int SpeakerFeatureSeqLen);

    private readonly record struct KvTensorData(float[] Values, int SeqLen);

    private readonly record struct LanguageModelOut(float[] LogitsLastStep, Dictionary<string, KvTensorData> PresentKeyValues);

    private readonly record struct WavData(float[] Samples, int SampleRate);

    private sealed class ChatterboxTokenizer
    {
        private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

        private readonly Dictionary<string, int> _vocab;
        private readonly Dictionary<string, int> _mergeRanks;
        private readonly int _unknownId;
        private readonly int _maxTokenLength;
        private readonly int _startId;
        private readonly int _stopId;
        private readonly int _startSpeechId;
        private readonly int _exaggerationId;

        private ChatterboxTokenizer(
            Dictionary<string, int> vocab,
            Dictionary<string, int> mergeRanks,
            int unknownId,
            int startId,
            int stopId,
            int startSpeechId,
            int exaggerationId)
        {
            _vocab = vocab;
            _mergeRanks = mergeRanks;
            _unknownId = unknownId;
            _startId = startId;
            _stopId = stopId;
            _startSpeechId = startSpeechId;
            _exaggerationId = exaggerationId;
            _maxTokenLength = Math.Max(1, vocab.Keys.Max(k => k.Length));
        }

        public static ChatterboxTokenizer Load(string tokenizerPath)
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(tokenizerPath));
            var root = doc.RootElement;
            var model = root.GetProperty("model");
            var vocabObj = model.GetProperty("vocab");

            var vocab = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var token in vocabObj.EnumerateObject())
            {
                vocab[token.Name] = token.Value.GetInt32();
            }

            var merges = new Dictionary<string, int>(StringComparer.Ordinal);
            if (model.TryGetProperty("merges", out var mergeArray) && mergeArray.ValueKind == JsonValueKind.Array)
            {
                var rank = 0;
                foreach (var mergeItem in mergeArray.EnumerateArray())
                {
                    var merge = mergeItem.GetString();
                    if (string.IsNullOrWhiteSpace(merge))
                    {
                        continue;
                    }

                    var split = merge.IndexOf(' ');
                    if (split <= 0 || split >= merge.Length - 1)
                    {
                        continue;
                    }

                    var left = merge[..split];
                    var right = merge[(split + 1)..];
                    merges[$"{left}\u0001{right}"] = rank++;
                }
            }

            var unkToken = model.TryGetProperty("unk_token", out var unkProp) ? unkProp.GetString() : "[UNK]";
            var unkId = vocab.TryGetValue(unkToken ?? "[UNK]", out var foundUnk) ? foundUnk : 1;

            var startId = 255;
            var stopId = 0;
            var startSpeechId = StartSpeechToken;
            var exaggerationId = 6563;
            if (root.TryGetProperty("post_processor", out var postProcessor) &&
                postProcessor.TryGetProperty("special_tokens", out var specialTokens))
            {
                startId = ReadSpecialTokenId(specialTokens, "[START]", startId);
                stopId = ReadSpecialTokenId(specialTokens, "[STOP]", stopId);
                startSpeechId = ReadSpecialTokenId(specialTokens, "[START_SPEECH]", startSpeechId);
                exaggerationId = ReadSpecialTokenId(specialTokens, "[EXAGGERATION]", exaggerationId);
            }

            return new ChatterboxTokenizer(vocab, merges, unkId, startId, stopId, startSpeechId, exaggerationId);
        }

        public IReadOnlyList<long> EncodeForModel(string text)
        {
            var normalized = WhitespaceRegex.Replace(text ?? string.Empty, " ");
            var symbols = normalized.EnumerateRunes().Select(r => r.ToString()).ToList();
            ApplyBpeMerges(symbols);

            var ids = new List<long>(symbols.Count + 6)
            {
                _exaggerationId,
                _startId
            };

            foreach (var symbol in symbols)
            {
                AppendSymbolIds(symbol, ids);
            }

            ids.Add(_stopId);
            ids.Add(_startSpeechId);
            ids.Add(_startSpeechId);
            return ids;
        }

        private static int ReadSpecialTokenId(JsonElement specialTokens, string tokenName, int fallback)
        {
            if (!specialTokens.TryGetProperty(tokenName, out var tokenData))
            {
                return fallback;
            }

            if (!tokenData.TryGetProperty("ids", out var ids) || ids.ValueKind != JsonValueKind.Array)
            {
                return fallback;
            }

            var first = ids.EnumerateArray().FirstOrDefault();
            return first.ValueKind == JsonValueKind.Number ? first.GetInt32() : fallback;
        }

        private void ApplyBpeMerges(List<string> symbols)
        {
            if (symbols.Count < 2 || _mergeRanks.Count == 0)
            {
                return;
            }

            while (symbols.Count > 1)
            {
                var bestIndex = -1;
                var bestRank = int.MaxValue;
                for (var i = 0; i < symbols.Count - 1; i++)
                {
                    var key = $"{symbols[i]}\u0001{symbols[i + 1]}";
                    if (_mergeRanks.TryGetValue(key, out var rank) && rank < bestRank)
                    {
                        bestRank = rank;
                        bestIndex = i;
                    }
                }

                if (bestIndex < 0)
                {
                    break;
                }

                symbols[bestIndex] = symbols[bestIndex] + symbols[bestIndex + 1];
                symbols.RemoveAt(bestIndex + 1);
            }
        }

        private void AppendSymbolIds(string symbol, List<long> ids)
        {
            if (_vocab.TryGetValue(symbol, out var exact))
            {
                ids.Add(exact);
                return;
            }

            if (symbol.Length <= 1)
            {
                ids.Add(_unknownId);
                return;
            }

            var pos = 0;
            while (pos < symbol.Length)
            {
                var bestId = _unknownId;
                var bestLen = 1;
                var maxLen = Math.Min(_maxTokenLength, symbol.Length - pos);

                for (var len = maxLen; len >= 1; len--)
                {
                    var piece = symbol.Substring(pos, len);
                    if (_vocab.TryGetValue(piece, out var pieceId))
                    {
                        bestId = pieceId;
                        bestLen = len;
                        break;
                    }
                }

                ids.Add(bestId);
                pos += bestLen;
            }
        }
    }
}
