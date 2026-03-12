using System.Diagnostics;
using System.Text;
using App.Core.Runtime;

namespace App.Inference;

public sealed class KittenPythonBackend : ITtsBackend
{
    private readonly LocalInferenceOptions _options;

    public KittenPythonBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "kitten-python";

    public async Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
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

        var pythonExe = ResolvePythonExePath();
        var workerScript = KittenPythonScript.EnsureRunnerScript();
        var hfCache = ResolvePythonCacheRoot(_options.ModelCacheDir);

        var psi = new ProcessStartInfo
        {
            FileName = pythonExe,
            WorkingDirectory = RuntimePathResolver.AppRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        psi.ArgumentList.Add(workerScript);
        psi.ArgumentList.Add("--text");
        psi.ArgumentList.Add(request.Text);
        psi.ArgumentList.Add("--output");
        psi.ArgumentList.Add(request.OutputPath);
        psi.ArgumentList.Add("--repo");
        psi.ArgumentList.Add(string.IsNullOrWhiteSpace(_options.ModelRepoId) ? "KittenML/kitten-tts-mini-0.8" : _options.ModelRepoId.Trim());
        psi.ArgumentList.Add("--voice");
        psi.ArgumentList.Add(ResolveVoiceKey(request.VoicePath));
        psi.ArgumentList.Add("--speed");
        psi.ArgumentList.Add(Math.Clamp(request.Speed, 0.5f, 1.5f).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));

        psi.Environment["PYTHONIOENCODING"] = "utf-8";
        psi.Environment["HF_HOME"] = hfCache;
        psi.Environment["HUGGINGFACE_HUB_CACHE"] = hfCache;

        using var proc = new Process { StartInfo = psi };
        if (!proc.Start())
        {
            throw new InvalidOperationException("Failed to start Kitten Python backend.");
        }

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync(ct);
        var stdout = (await stdoutTask).Trim();
        var stderr = (await stderrTask).Trim();
        if (proc.ExitCode == 0)
        {
            return;
        }

        var detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        if (string.IsNullOrWhiteSpace(detail))
        {
            detail = $"Kitten Python backend failed with exit code {proc.ExitCode}.";
        }

        throw new InvalidOperationException(detail);
    }

    private static string ResolveVoiceKey(string? rawSelection)
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

        return string.IsNullOrWhiteSpace(value) ? "expr-voice-5-m" : value;
    }

    private static string ResolvePythonExePath()
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var candidates = new[]
        {
            Path.Combine(appRoot, "python_kitten", "python.exe"),
            Path.Combine(appRoot, "python_kitten", "Scripts", "python.exe"),
            Path.Combine(appRoot, "tools", "python_kitten", "python.exe"),
            Path.Combine(appRoot, "tools", "python_kitten", "Scripts", "python.exe"),
            Path.GetFullPath(Path.Combine(appRoot, "..", ".venv", "Scripts", "python.exe"))
        };

        foreach (var path in candidates)
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        throw new FileNotFoundException("Bundled Python runtime not found for Kitten backend (python_kitten).");
    }

    private static string ResolvePythonCacheRoot(string modelCacheDir)
    {
        if (string.IsNullOrWhiteSpace(modelCacheDir))
        {
            return Path.Combine(RuntimePathResolver.AppRoot, "models", "hf-cache");
        }

        var absolute = Path.IsPathRooted(modelCacheDir)
            ? modelCacheDir
            : Path.Combine(RuntimePathResolver.AppRoot, modelCacheDir);
        return Path.Combine(absolute, "hf-cache");
    }
}

internal static class KittenPythonScript
{
    private const string ResourceName = "App.Inference.Resources.kitten_python_runner.py";

    public static string EnsureRunnerScript()
    {
        var toolsDir = Path.Combine(RuntimePathResolver.AppRoot, "tools", "python_backends");
        Directory.CreateDirectory(toolsDir);
        var path = Path.Combine(toolsDir, "kitten_python_runner.py");
        using var stream = typeof(KittenPythonScript).Assembly.GetManifestResourceStream(ResourceName)
            ?? throw new FileNotFoundException($"Embedded resource missing: {ResourceName}");
        using var fs = File.Create(path);
        stream.CopyTo(fs);
        return path;
    }
}
