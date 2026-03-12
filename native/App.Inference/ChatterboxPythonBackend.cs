using System.Diagnostics;
using System.Text;
using App.Core.Runtime;

namespace App.Inference;

public sealed class ChatterboxPythonBackend : ITtsBackend
{
    private readonly LocalInferenceOptions _options;

    public ChatterboxPythonBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public string Name => "chatterbox-python";

    public async Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            throw new ArgumentException("Input text is empty.");
        }
        if (string.IsNullOrWhiteSpace(request.VoicePath) || !File.Exists(request.VoicePath))
        {
            throw new FileNotFoundException("Voice file not found.", request.VoicePath);
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
        var workerScript = ChatterboxPythonScript.EnsureRunnerScript();
        var repoDir = ResolveRepoFolder(_options.ModelCacheDir, _options.ModelRepoId);
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
        psi.ArgumentList.Add("--voice");
        psi.ArgumentList.Add(request.VoicePath);
        psi.ArgumentList.Add("--output");
        psi.ArgumentList.Add(request.OutputPath);
        psi.ArgumentList.Add("--device");
        psi.ArgumentList.Add(NormalizePythonDevice(_options.PreferDevice));
        psi.ArgumentList.Add("--exaggeration");
        psi.ArgumentList.Add((request.ChatterboxExaggeration ?? _options.Exaggeration).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        if (!string.IsNullOrWhiteSpace(repoDir))
        {
            psi.ArgumentList.Add("--repo-dir");
            psi.ArgumentList.Add(repoDir);
        }

        psi.Environment["PYTHONIOENCODING"] = "utf-8";
        psi.Environment["HF_HOME"] = hfCache;
        psi.Environment["HUGGINGFACE_HUB_CACHE"] = hfCache;

        using var proc = new Process { StartInfo = psi };
        if (!proc.Start())
        {
            throw new InvalidOperationException("Failed to start Chatterbox Python backend.");
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
            detail = $"Chatterbox Python backend failed with exit code {proc.ExitCode}.";
        }

        throw new InvalidOperationException(detail);
    }

    private static string NormalizePythonDevice(string? preferDevice)
    {
        return (preferDevice ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "gpu" => "cuda",
            "cpu" => "cpu",
            _ => "auto"
        };
    }

    private static string ResolvePythonExePath()
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var candidates = new[]
        {
            Path.Combine(appRoot, "python_chatterbox", "python.exe"),
            Path.Combine(appRoot, "python_chatterbox", "Scripts", "python.exe"),
            Path.Combine(appRoot, "tools", "python_chatterbox", "python.exe"),
            Path.Combine(appRoot, "tools", "python_chatterbox", "Scripts", "python.exe"),
            Path.GetFullPath(Path.Combine(appRoot, "..", ".venv", "Scripts", "python.exe"))
        };

        foreach (var path in candidates)
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        throw new FileNotFoundException("Bundled Python runtime not found for Chatterbox backend (python_chatterbox).");
    }

    private static string ResolvePythonCacheRoot(string modelCacheDir)
    {
        var absolute = ResolveModelCacheRoot(modelCacheDir);
        return Path.Combine(absolute, "hf-cache");
    }

    private static string ResolveRepoFolder(string modelCacheDir, string repoId)
    {
        var absolute = ResolveModelCacheRoot(modelCacheDir);
        var tokens = (repoId ?? string.Empty).Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2)
        {
            return string.Empty;
        }

        var repoFolder = Path.Combine(absolute, "hf-cache", $"models--{tokens[0]}--{tokens[1]}");
        return Directory.Exists(repoFolder) ? repoFolder : string.Empty;
    }

    private static string ResolveModelCacheRoot(string modelCacheDir)
    {
        if (string.IsNullOrWhiteSpace(modelCacheDir))
        {
            return Path.Combine(RuntimePathResolver.AppRoot, "models");
        }

        return Path.IsPathRooted(modelCacheDir)
            ? modelCacheDir
            : Path.Combine(RuntimePathResolver.AppRoot, modelCacheDir);
    }
}

internal static class ChatterboxPythonScript
{
    private const string ResourceName = "App.Inference.Resources.chatterbox_python_runner.py";

    public static string EnsureRunnerScript()
    {
        var toolsDir = Path.Combine(RuntimePathResolver.AppRoot, "tools", "python_backends");
        Directory.CreateDirectory(toolsDir);
        var path = Path.Combine(toolsDir, "chatterbox_python_runner.py");
        using var stream = typeof(ChatterboxPythonScript).Assembly.GetManifestResourceStream(ResourceName)
            ?? throw new FileNotFoundException($"Embedded resource missing: {ResourceName}");
        using var fs = File.Create(path);
        stream.CopyTo(fs);
        return path;
    }
}
