using System.Diagnostics;
using System.Text;
using System.Text.Json;
using App.Core.Runtime;

namespace App.Inference;

public sealed class QwenPythonBackend : ITtsBackend, IDisposable
{
    private const string ProtocolPrefix = "@@QWENJSON@@";
    private static readonly object WorkerLogSync = new();

    private readonly object _sync = new();
    private readonly LocalInferenceOptions _options;
    private Process? _proc;
    private Task? _stderrPump;
    private string _loadedKey = string.Empty;
    private bool _disposed;

    public QwenPythonBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "qwen-python-worker";
    public string ActiveExecutionProvider { get; private set; } = "python";
    private static string WorkerStdErrLogPath => Path.Combine(RuntimePathResolver.AppRoot, "qwen_worker_stderr.log");

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        return Task.Run(() => SynthesizeInternal(request, ct), ct);
    }

    public Task GenerateVoiceDesignAsync(
        string text,
        string instruct,
        string language,
        string outputPath,
        bool? doSample = null,
        float? temperature = null,
        int? topK = null,
        float? topP = null,
        float? repetitionPenalty = null,
        IProgress<QwenWorkerProgress>? progress = null,
        CancellationToken ct = default)
    {
        return Task.Run(() => GenerateVoiceDesignInternal(
            text,
            instruct,
            language,
            outputPath,
            doSample,
            temperature,
            topK,
            topP,
            repetitionPenalty,
            progress,
            ct), ct);
    }

    private void SynthesizeInternal(TtsRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.Text))
            throw new ArgumentException("Input text is empty.");
        if (string.IsNullOrWhiteSpace(request.VoicePath))
            throw new FileNotFoundException("Voice file not found.", request.VoicePath);
        if (!CanUseVoicePathForCurrentModel(request.VoicePath))
            throw new FileNotFoundException("Voice file not found.", request.VoicePath);
        if (string.IsNullOrWhiteSpace(request.OutputPath))
            throw new ArgumentException("Output path is required.");

        var outDir = Path.GetDirectoryName(request.OutputPath);
        if (!string.IsNullOrWhiteSpace(outDir))
            Directory.CreateDirectory(outDir);

        EnsureWorkerStarted(ct);

        var hasRefText = !string.IsNullOrWhiteSpace(request.RefText);
        var xVectorOnlyMode = request.QwenXVectorOnlyMode ?? !hasRefText;
        if (!hasRefText)
        {
            xVectorOnlyMode = true;
        }

        var cmd = new
        {
            cmd = "synthesize",
            text = request.Text,
            voice_path = request.VoicePath,
            ref_text = hasRefText ? request.RefText : null,
            output_path = request.OutputPath,
            qwen = new
            {
                do_sample = request.QwenDoSample,
                temperature = request.QwenTemperature,
                top_k = request.QwenTopK,
                top_p = request.QwenTopP,
                repetition_penalty = request.QwenRepetitionPenalty,
                max_new_tokens = _options.MaxNewTokens,
                x_vector_only_mode = xVectorOnlyMode
            }
        };

        var res = SendCommand(cmd, ct);
        if (!res.TryGetProperty("ok", out var okNode) || !okNode.GetBoolean())
        {
            var err = res.TryGetProperty("error", out var errNode) ? errNode.GetString() : "Unknown Qwen Python worker error.";
            throw new InvalidOperationException(err);
        }
    }

    private bool CanUseVoicePathForCurrentModel(string voicePath)
    {
        var raw = (voicePath ?? string.Empty).Trim();
        if (raw.Length == 0)
        {
            return false;
        }

        // Only CustomVoice models accept token-based speaker IDs.
        if (IsCustomVoiceRepo() && raw.StartsWith("qwen-customvoice://", StringComparison.OrdinalIgnoreCase))
        {
            var speaker = raw["qwen-customvoice://".Length..].Trim();
            return speaker.Length > 0;
        }

        return File.Exists(raw);
    }

    private bool IsCustomVoiceRepo()
    {
        var repo = (_options.ModelRepoId ?? string.Empty).Trim();
        return repo.Contains("customvoice", StringComparison.OrdinalIgnoreCase) ||
               repo.Contains("custom_voice", StringComparison.OrdinalIgnoreCase);
    }

    private void GenerateVoiceDesignInternal(
        string text,
        string instruct,
        string language,
        string outputPath,
        bool? doSample,
        float? temperature,
        int? topK,
        float? topP,
        float? repetitionPenalty,
        IProgress<QwenWorkerProgress>? progress,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("Input text is empty.");
        if (string.IsNullOrWhiteSpace(instruct))
            throw new ArgumentException("Voice design prompt is empty.");
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path is required.");

        var outDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outDir))
            Directory.CreateDirectory(outDir);

        EnsureWorkerStarted(ct);

        var cmd = new
        {
            cmd = "voice_design",
            text = text,
            instruct = instruct,
            language = string.IsNullOrWhiteSpace(language) ? "auto" : language.Trim(),
            output_path = outputPath,
            qwen = new
            {
                do_sample = doSample,
                temperature = temperature,
                top_k = topK,
                top_p = topP,
                repetition_penalty = repetitionPenalty,
                max_new_tokens = _options.MaxNewTokens
            }
        };

        var res = SendCommand(cmd, ct, evt =>
        {
            if (!evt.TryGetProperty("event", out var eventNode) ||
                !string.Equals(eventNode.GetString(), "progress", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var message = evt.TryGetProperty("message", out var msgNode) ? msgNode.GetString() ?? string.Empty : string.Empty;
            var stage = evt.TryGetProperty("stage", out var stageNode) ? stageNode.GetString() : null;
            int? step = null;
            int? totalSteps = null;
            if (evt.TryGetProperty("step", out var stepNode) && stepNode.ValueKind == JsonValueKind.Number && stepNode.TryGetInt32(out var s))
            {
                step = s;
            }
            if (evt.TryGetProperty("total_steps", out var totalNode) && totalNode.ValueKind == JsonValueKind.Number && totalNode.TryGetInt32(out var t))
            {
                totalSteps = t;
            }

            progress?.Report(new QwenWorkerProgress(message, stage, step, totalSteps));
        });
        if (!res.TryGetProperty("ok", out var okNode) || !okNode.GetBoolean())
        {
            var err = res.TryGetProperty("error", out var errNode) ? errNode.GetString() : "Unknown Qwen Python worker error.";
            throw new InvalidOperationException(err);
        }
    }

    private void EnsureWorkerStarted(CancellationToken ct)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            var key = BuildWorkerKey();
            if (_proc is not null && !_proc.HasExited && string.Equals(_loadedKey, key, StringComparison.Ordinal))
                return;

            StopWorkerLocked();

            var pythonExe = ResolvePythonExePath();
            var workerScript = QwenPythonWorkerScript.EnsureWorkerScript();
            var modelRepo = ResolveOfficialQwenRepoForPython(_options.ModelRepoId);
            var cacheRoot = ResolveModelCacheForPython(_options.ModelCacheDir);
            Directory.CreateDirectory(cacheRoot);

            var psi = new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = $"\"{workerScript}\"",
                WorkingDirectory = RuntimePathResolver.AppRoot,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            psi.Environment["PYTHONIOENCODING"] = "utf-8";
            psi.Environment["HF_HOME"] = cacheRoot;
            psi.Environment["HUGGINGFACE_HUB_CACHE"] = cacheRoot;
            if (string.Equals(_options.PreferDevice, "gpu", StringComparison.OrdinalIgnoreCase))
            {
                psi.Environment["QWEN_WORKER_DEVICE"] = "cuda";
            }
            else if (string.Equals(_options.PreferDevice, "cpu", StringComparison.OrdinalIgnoreCase))
            {
                psi.Environment["QWEN_WORKER_DEVICE"] = "cpu";
            }
            else
            {
                psi.Environment["QWEN_WORKER_DEVICE"] = "auto";
            }
            psi.Environment["QWEN_WORKER_MODEL_REPO"] = modelRepo;

            var proc = new Process { StartInfo = psi, EnableRaisingEvents = false };
            if (!proc.Start())
                throw new InvalidOperationException("Failed to start Qwen Python worker.");

            _stderrPump = Task.Run(async () =>
            {
                try
                {
                    while (!proc.HasExited)
                    {
                        var line = await proc.StandardError.ReadLineAsync();
                        if (line is null)
                            break;
                        AppendWorkerStdErr(line);
                    }
                }
                catch
                {
                    // Best effort pump only.
                }
            });

            _proc = proc;
            var initRes = ReadProtocolResponse(ct);
            if (!initRes.TryGetProperty("ok", out var initOk) || !initOk.GetBoolean())
            {
                var err = initRes.TryGetProperty("error", out var errNode) ? errNode.GetString() : "Qwen Python worker init failed.";
                StopWorkerLocked();
                throw new InvalidOperationException(err);
            }

            if (initRes.TryGetProperty("provider", out var providerNode))
            {
                ActiveExecutionProvider = $"python-{providerNode.GetString()}";
            }
            else
            {
                ActiveExecutionProvider = "python";
            }

            _loadedKey = key;
        }
    }

    private JsonElement SendCommand<T>(T command, CancellationToken ct, Action<JsonElement>? onEvent = null)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            if (_proc is null || _proc.HasExited)
                throw new InvalidOperationException("Qwen Python worker is not running.");

            var json = JsonSerializer.Serialize(command);
            _proc.StandardInput.WriteLine(json);
            _proc.StandardInput.Flush();
            return ReadProtocolResponse(ct, onEvent);
        }
    }

    private JsonElement ReadProtocolResponse(CancellationToken ct, Action<JsonElement>? onEvent = null)
    {
        if (_proc is null)
            throw new InvalidOperationException("Qwen Python worker is not running.");

        while (true)
        {
            ct.ThrowIfCancellationRequested();
            var line = _proc.StandardOutput.ReadLine();
            if (line is null)
            {
                throw new InvalidOperationException("Qwen Python worker terminated unexpectedly.");
            }

            if (!line.StartsWith(ProtocolPrefix, StringComparison.Ordinal))
            {
                // Worker can print warnings/noise to stdout (qwen_tts/sox warnings). Ignore.
                continue;
            }

            var payload = line.Substring(ProtocolPrefix.Length);
            using var doc = JsonDocument.Parse(payload);
            var root = doc.RootElement.Clone();
            if (root.TryGetProperty("event", out var eventNode))
            {
                onEvent?.Invoke(root);
                if (string.Equals(eventNode.GetString(), "progress", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }
            return root;
        }
    }

    private static void AppendWorkerStdErr(string line)
    {
        try
        {
            lock (WorkerLogSync)
            {
                File.AppendAllText(WorkerStdErrLogPath, $"[{DateTime.UtcNow:O}] {line}{Environment.NewLine}");
            }
        }
        catch
        {
            // Never crash while writing worker logs.
        }
    }

    private static string ResolveOfficialQwenRepoForPython(string repo)
    {
        var r = (repo ?? string.Empty).Trim();
        if (r.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            if (r.Contains("voicedesign", StringComparison.OrdinalIgnoreCase))
            {
                return "Qwen/Qwen3-TTS-12Hz-1.7B-VoiceDesign";
            }
            if (r.Contains("customvoice", StringComparison.OrdinalIgnoreCase))
            {
                return r.Contains("0.6", StringComparison.OrdinalIgnoreCase)
                    ? "Qwen/Qwen3-TTS-12Hz-0.6B-CustomVoice"
                    : "Qwen/Qwen3-TTS-12Hz-1.7B-CustomVoice";
            }
            if (r.Contains("0.6", StringComparison.OrdinalIgnoreCase) || r.Contains("onnx-dll", StringComparison.OrdinalIgnoreCase))
            {
                return "Qwen/Qwen3-TTS-12Hz-0.6B-Base";
            }

            return "Qwen/Qwen3-TTS-12Hz-1.7B-Base";
        }

        return "Qwen/Qwen3-TTS-12Hz-1.7B-Base";
    }

    private static string ResolveModelCacheForPython(string cacheSetting)
    {
        if (string.IsNullOrWhiteSpace(cacheSetting))
            return Path.Combine(RuntimePathResolver.AppRoot, "models", "hf-cache-qwen-py");
        if (Path.IsPathRooted(cacheSetting))
            return Path.Combine(cacheSetting, "hf-cache-qwen-py");
        return Path.Combine(RuntimePathResolver.AppRoot, cacheSetting, "hf-cache-qwen-py");
    }

    private static string ResolvePythonExePath()
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var candidates = new[]
        {
            Path.Combine(appRoot, "python_qwen", "python.exe"),
            Path.Combine(appRoot, "python_qwen", "Scripts", "python.exe"),
            Path.Combine(appRoot, "tools", "python_qwen", "python.exe"),
            Path.Combine(appRoot, "tools", "python_qwen", "Scripts", "python.exe"),
            Path.GetFullPath(Path.Combine(appRoot, "..", ".venv", "Scripts", "python.exe"))
        };

        foreach (var path in candidates)
        {
            if (File.Exists(path))
                return path;
        }

        throw new FileNotFoundException(
            "Qwen Python runtime not found. Bundle a private Python env at customer_release\\python_qwen\\python.exe (or \\Scripts\\python.exe).");
    }

    private string BuildWorkerKey()
    {
        var repo = ResolveOfficialQwenRepoForPython(_options.ModelRepoId);
        var device = (_options.PreferDevice ?? "auto").Trim().ToLowerInvariant();
        var cache = ResolveModelCacheForPython(_options.ModelCacheDir).ToLowerInvariant();
        return $"{repo}|{device}|{cache}";
    }

    private void StopWorkerLocked()
    {
        if (_proc is not null)
        {
            try
            {
                if (!_proc.HasExited)
                {
                    try
                    {
                        _proc.StandardInput.WriteLine("{\"cmd\":\"shutdown\"}");
                        _proc.StandardInput.Flush();
                    }
                    catch { }
                    if (!_proc.WaitForExit(1500))
                    {
                        _proc.Kill(entireProcessTree: true);
                        _proc.WaitForExit(2000);
                    }
                }
            }
            catch { }
            finally
            {
                _proc.Dispose();
                _proc = null;
            }
        }

        _stderrPump = null;
        _loadedKey = string.Empty;
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
                return;
            _disposed = true;
            StopWorkerLocked();
        }
        GC.SuppressFinalize(this);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(QwenPythonBackend));
    }
}

public sealed record QwenWorkerProgress(string Message, string? Stage = null, int? Step = null, int? TotalSteps = null);

internal static class QwenPythonWorkerScript
{
    private const string ResourceName = "App.Inference.Resources.qwen_python_worker.py";

    public static string EnsureWorkerScript()
    {
        var toolsDir = Path.Combine(RuntimePathResolver.AppRoot, "tools", "qwen_python");
        Directory.CreateDirectory(toolsDir);
        var path = Path.Combine(toolsDir, "qwen_python_worker.py");
        using var stream = typeof(QwenPythonWorkerScript).Assembly.GetManifestResourceStream(ResourceName)
            ?? throw new FileNotFoundException($"Embedded resource missing: {ResourceName}");
        using var fs = File.Create(path);
        stream.CopyTo(fs);
        return path;
    }
}
