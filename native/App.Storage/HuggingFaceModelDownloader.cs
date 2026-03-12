using System.Net.Http;
using System.Text.Json;
using System.Diagnostics;
using App.Core.Models;
using App.Core.Runtime;

namespace App.Storage;

public sealed class HuggingFaceModelDownloader
{
    private static readonly string[] LegacyRequiredFiles =
    {
        "conds.pt",
        "s3gen.pt",
        "s3gen.safetensors",
        "t3_cfg.pt",
        "t3_cfg.safetensors",
        "tokenizer.json",
        "ve.pt",
        "ve.safetensors",
        "README.md"
    };
    private static readonly string[] OnnxRequiredFiles =
    {
        "README.md",
        "onnx/conditional_decoder.onnx",
        "onnx/conditional_decoder.onnx_data",
        "onnx/embed_tokens.onnx",
        "onnx/embed_tokens.onnx_data",
        "onnx/language_model.onnx",
        "onnx/language_model.onnx_data",
        "onnx/speech_encoder.onnx",
        "onnx/speech_encoder.onnx_data",
        "config.json",
        "generation_config.json",
        "preprocessor_config.json",
        "tokenizer_config.json",
        "tokenizer.json",
        "default_voice.wav"
    };
    private static readonly string[] QwenOnnxDllRequiredFiles =
    {
        "README.md",
        "THIRD_PARTY_LICENSES.txt",
        "qwen3_tts_rust.dll",
        "models/Qwen3-TTS-12Hz-0.6B-Base/config.json",
        "models/Qwen3-TTS-12Hz-0.6B-Base/merges.txt",
        "models/Qwen3-TTS-12Hz-0.6B-Base/tokenizer_config.json",
        "models/Qwen3-TTS-12Hz-0.6B-Base/vocab.json",
        "models/Qwen3-TTS-12Hz-1.7B-Base/config.json",
        "models/Qwen3-TTS-12Hz-1.7B-Base/merges.txt",
        "models/Qwen3-TTS-12Hz-1.7B-Base/tokenizer_config.json",
        "models/Qwen3-TTS-12Hz-1.7B-Base/vocab.json",
        "onnx_kv_06b/code_predictor.onnx",
        "onnx_kv_06b/code_predictor_embed.onnx",
        "onnx_kv_06b/codec_embed.onnx",
        "onnx_kv_06b/speaker_encoder.onnx",
        "onnx_kv_06b/talker_decode.onnx",
        "onnx_kv_06b/talker_prefill.onnx",
        "onnx_kv_06b/text_project.onnx",
        "onnx_kv_06b/tokenizer12hz_decode.onnx",
        "onnx_kv_06b/tokenizer12hz_encode.onnx"
    };
    private static readonly string[] QwenOnnxXkosRequiredPrefixes =
    {
        "onnx/shared/",
        "onnx/voice_clone/",
        "tokenizer/"
    };
    private static readonly string[] QwenOnnxXkosRootRequiredFiles =
    {
        "voice_clone_config.json",
        "README.md"
    };
    private static readonly string[] KittenTtsRequiredFiles =
    {
        "config.json",
        "kitten_tts_mini_v0_8.onnx",
        "voices.npz"
    };

    private readonly HttpClient _httpClient;
    private static readonly HashSet<string> GenericAllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".json", ".txt", ".model", ".tiktoken", ".safetensors", ".bin", ".pt", ".onnx", ".onnx_data", ".wav", ".md", ".dll", ".npz"
    };

    public HuggingFaceModelDownloader(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
    }

    public async Task DownloadAsync(
        AppConfig config,
        IProgress<string>? progress = null,
        IProgress<ModelDownloadTelemetry>? telemetry = null,
        CancellationToken ct = default,
        bool forceRedownload = false)
    {
        var repoId = string.IsNullOrWhiteSpace(config.ModelRepoId) ? "onnx-community/chatterbox-ONNX" : config.ModelRepoId.Trim();
        var state = new DownloadTelemetryState();
        await DownloadRepoAsync(repoId, config.ModelCacheDir, progress, telemetry, state, ct, forceRedownload);

        var preset = (config.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        if (preset != "qwen3_tts")
        {
            return;
        }

        var tokenizerRepo = string.IsNullOrWhiteSpace(config.AdditionalModelRepoId)
            ? string.Empty
            : config.AdditionalModelRepoId.Trim();
        if (string.IsNullOrWhiteSpace(tokenizerRepo))
        {
            return;
        }
        if (string.Equals(tokenizerRepo, repoId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        await DownloadRepoAsync(tokenizerRepo, config.ModelCacheDir, progress, telemetry, state, ct, forceRedownload);
    }

    private static bool IsChatterboxOnnxRepo(string repoId)
    {
        return repoId.Contains("chatterbox-onnx", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsQwenOnnxDllRepo(string repoId)
    {
        return repoId.Contains("qwen3-tts-onnx-dll", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsQwenOnnxXkosRepo(string repoId)
    {
        return repoId.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase) &&
               repoId.Contains("onnx", StringComparison.OrdinalIgnoreCase) &&
               !IsQwenOnnxDllRepo(repoId);
    }

    private static bool IsKittenTtsRepo(string repoId)
    {
        return repoId.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveModelCacheDir(string modelCacheDir)
    {
        return ModelCachePath.ResolveAbsolute(modelCacheDir, RuntimePaths.AppRoot);
    }

    private async Task DownloadRepoAsync(
        string repoId,
        string modelCacheDirRaw,
        IProgress<string>? progress,
        IProgress<ModelDownloadTelemetry>? telemetry,
        DownloadTelemetryState state,
        CancellationToken ct,
        bool forceRedownload)
    {
        var modelCacheDir = ResolveModelCacheDir(modelCacheDirRaw);
        var (owner, name) = ParseRepo(repoId);
        var targetDir = Path.Combine(modelCacheDir, "hf-cache", $"models--{owner}--{name}");
        Directory.CreateDirectory(targetDir);

        IReadOnlyList<string> requiredFiles;
        if (IsChatterboxOnnxRepo(repoId))
        {
            requiredFiles = OnnxRequiredFiles;
        }
        else if (IsQwenOnnxDllRepo(repoId))
        {
            requiredFiles = QwenOnnxDllRequiredFiles;
        }
        else if (IsQwenOnnxXkosRepo(repoId))
        {
            requiredFiles = await DiscoverRepoFilesByPrefixAsync(repoId, QwenOnnxXkosRequiredPrefixes, QwenOnnxXkosRootRequiredFiles, ct);
        }
        else if (IsKittenTtsRepo(repoId))
        {
            requiredFiles = KittenTtsRequiredFiles;
        }
        else if (IsLegacyChatterboxRepo(repoId))
        {
            requiredFiles = LegacyRequiredFiles;
        }
        else
        {
            requiredFiles = await DiscoverGenericRepoFilesAsync(repoId, ct);
            if (requiredFiles.Count == 0)
            {
                throw new InvalidOperationException($"No downloadable model files discovered for repo: {repoId}");
            }
        }

        state.FilesTotal += requiredFiles.Count;

        var remoteSizes = new Dictionary<string, long?>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in requiredFiles)
        {
            ct.ThrowIfCancellationRequested();
            var url = $"https://huggingface.co/{repoId}/resolve/main/{file}";
            var len = await TryGetContentLengthAsync(url, ct);
            remoteSizes[file] = len;
            if (len.HasValue && len.Value > 0)
            {
                state.TotalExpected += len.Value;
            }
            else
            {
                state.HasUnknownTotals = true;
            }
        }

        foreach (var file in requiredFiles)
        {
            ct.ThrowIfCancellationRequested();
            var targetPath = Path.Combine(targetDir, file.Replace('/', Path.DirectorySeparatorChar));
            var targetFolder = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            var url = $"https://huggingface.co/{repoId}/resolve/main/{file}";
            var expectedFileSize = remoteSizes.TryGetValue(file, out var known) ? known : null;
            var shouldDownload = forceRedownload || await ShouldDownloadAsync(url, targetPath, ct);
            if (!shouldDownload)
            {
                var accounted = expectedFileSize ?? new FileInfo(targetPath).Length;
                if (accounted > 0)
                {
                    state.TotalDownloaded += accounted;
                }

                state.FilesCompleted++;
                progress?.Report($"[{repoId}] Skip existing: {file}");
                ReportTelemetry(
                    telemetry,
                    state,
                    repoId,
                    file,
                    fileBytesDownloaded: accounted,
                    fileBytesTotal: expectedFileSize,
                    message: $"[{repoId}] Skip existing: {file}");
                continue;
            }

            progress?.Report($"[{repoId}] Downloading {file} ...");
            using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Failed to download {file} from {repoId}: {(int)response.StatusCode} {response.ReasonPhrase}");
            }

            var tmpPath = $"{targetPath}.download-{Guid.NewGuid():N}.tmp";
            var fileTotal = expectedFileSize ?? response.Content.Headers.ContentLength;
            var fileDownloaded = 0L;
            var buffer = new byte[1024 * 64];
            var reportWatch = Stopwatch.StartNew();
            var throughputWatch = Stopwatch.StartNew();
            var throughputBaselineBytes = state.TotalDownloaded;

            try
            {
                await using (var source = await response.Content.ReadAsStreamAsync(ct))
                await using (var dest = File.Create(tmpPath))
                {
                    while (true)
                    {
                        var read = await source.ReadAsync(buffer, ct);
                        if (read <= 0)
                        {
                            break;
                        }

                        await dest.WriteAsync(buffer.AsMemory(0, read), ct);
                        fileDownloaded += read;
                        state.TotalDownloaded += read;

                        if (reportWatch.ElapsedMilliseconds >= 220)
                        {
                            var elapsedSec = Math.Max(0.001, throughputWatch.Elapsed.TotalSeconds);
                            var bytesPerSecond = (state.TotalDownloaded - throughputBaselineBytes) / elapsedSec;
                            ReportTelemetry(
                                telemetry,
                                state,
                                repoId,
                                file,
                                fileDownloaded,
                                fileTotal,
                                $"[{repoId}] Downloading {file} ...",
                                bytesPerSecond);
                            reportWatch.Restart();
                        }
                    }
                }

                if (fileTotal.HasValue && fileTotal.Value > 0 && fileDownloaded != fileTotal.Value)
                {
                    throw new IOException(
                        $"Incomplete download for {file} from {repoId}. Expected {fileTotal.Value} bytes, got {fileDownloaded} bytes.");
                }

                await FinalizeDownloadedFileAsync(
                    tempPath: tmpPath,
                    targetPath: targetPath,
                    downloadedBytes: fileDownloaded,
                    expectedFileSize: fileTotal,
                    ct: ct);
            }
            finally
            {
                TryDeleteFile(tmpPath);
            }

            state.FilesCompleted++;
            var finalElapsedSec = Math.Max(0.001, throughputWatch.Elapsed.TotalSeconds);
            var finalBps = (state.TotalDownloaded - throughputBaselineBytes) / finalElapsedSec;
            ReportTelemetry(
                telemetry,
                state,
                repoId,
                file,
                fileDownloaded,
                fileTotal,
                $"[{repoId}] Downloaded: {file}",
                finalBps);

            progress?.Report($"[{repoId}] Downloaded: {file}");
        }

        var missing = requiredFiles
            .Where(file => !File.Exists(Path.Combine(targetDir, file.Replace('/', Path.DirectorySeparatorChar))))
            .Take(5)
            .ToList();
        if (missing.Count > 0)
        {
            throw new InvalidOperationException(
                $"Model download incomplete for {repoId}. Missing: {string.Join(", ", missing)}");
        }
    }

    private static void ReportTelemetry(
        IProgress<ModelDownloadTelemetry>? telemetry,
        DownloadTelemetryState state,
        string repoId,
        string file,
        long fileBytesDownloaded,
        long? fileBytesTotal,
        string message,
        double bytesPerSecond = 0.0)
    {
        telemetry?.Report(new ModelDownloadTelemetry(
            RepoId: repoId,
            FilePath: file,
            BytesDownloaded: state.TotalDownloaded,
            BytesTotal: state.HasUnknownTotals ? null : state.TotalExpected,
            FileBytesDownloaded: fileBytesDownloaded,
            FileBytesTotal: fileBytesTotal,
            BytesPerSecond: Math.Max(0.0, bytesPerSecond),
            FilesCompleted: state.FilesCompleted,
            FilesTotal: Math.Max(state.FilesTotal, state.FilesCompleted),
            Message: message));
    }

    private static async Task FinalizeDownloadedFileAsync(
        string tempPath,
        string targetPath,
        long downloadedBytes,
        long? expectedFileSize,
        CancellationToken ct)
    {
        const int maxAttempts = 4;
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                File.Move(tempPath, targetPath, overwrite: true);
                return;
            }
            catch (IOException) when (attempt < maxAttempts - 1)
            {
                await Task.Delay(250 * (attempt + 1), ct);
            }
            catch (UnauthorizedAccessException) when (attempt < maxAttempts - 1)
            {
                await Task.Delay(250 * (attempt + 1), ct);
            }
        }

        TryDeleteFile(tempPath);
        throw new IOException(
            $"The process cannot access the file '{targetPath}' because it is being used by another process. " +
            "Stop generation (if running), close other app instances, then retry download.");
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Ignore cleanup failures.
        }
    }

    private async Task<long?> TryGetContentLengthAsync(string url, CancellationToken ct)
    {
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Head, url);
            using var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            if (!res.IsSuccessStatusCode)
            {
                return null;
            }

            var len = res.Content.Headers.ContentLength;
            return len.HasValue && len.Value > 0 ? len.Value : null;
        }
        catch
        {
            return null;
        }
    }

    private async Task<bool> ShouldDownloadAsync(string url, string targetPath, CancellationToken ct)
    {
        if (!File.Exists(targetPath))
        {
            return true;
        }

        var localLen = new FileInfo(targetPath).Length;
        if (localLen <= 0)
        {
            return true;
        }

        // If remote content length is known and mismatched, force re-download.
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Head, url);
            using var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            if (res.IsSuccessStatusCode)
            {
                var remoteLen = res.Content.Headers.ContentLength;
                if (remoteLen.HasValue && remoteLen.Value > 0 && remoteLen.Value != localLen)
                {
                    return true;
                }
            }
        }
        catch
        {
            // Keep current file when remote metadata probe fails.
        }

        return false;
    }

    private async Task<IReadOnlyList<string>> DiscoverGenericRepoFilesAsync(string repoId, CancellationToken ct)
    {
        var url = $"https://huggingface.co/api/models/{repoId}";
        using var response = await _httpClient.GetAsync(url, ct);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Failed to inspect repo {repoId}: {(int)response.StatusCode} {response.ReasonPhrase}");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
        if (!doc.RootElement.TryGetProperty("siblings", out var siblings) || siblings.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<string>();
        }

        var files = new List<string>();
        foreach (var sibling in siblings.EnumerateArray())
        {
            if (!sibling.TryGetProperty("rfilename", out var fileProp))
            {
                continue;
            }
            var file = fileProp.GetString();
            if (string.IsNullOrWhiteSpace(file))
            {
                continue;
            }
            if (file.Contains("/.", StringComparison.Ordinal) || file.StartsWith(".", StringComparison.Ordinal))
            {
                continue;
            }
            if (file.EndsWith(".lock", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var ext = Path.GetExtension(file);
            if (!string.IsNullOrWhiteSpace(ext) && GenericAllowedExtensions.Contains(ext))
            {
                files.Add(file);
                continue;
            }

            if (file.Equals("README.md", StringComparison.OrdinalIgnoreCase) ||
                file.Equals("config.json", StringComparison.OrdinalIgnoreCase) ||
                file.Equals("generation_config.json", StringComparison.OrdinalIgnoreCase) ||
                file.Contains("tokenizer", StringComparison.OrdinalIgnoreCase))
            {
                files.Add(file);
            }
        }

        return files;
    }

    private async Task<IReadOnlyList<string>> DiscoverRepoFilesByPrefixAsync(
        string repoId,
        IReadOnlyList<string> prefixes,
        IReadOnlyList<string>? rootRequiredFiles,
        CancellationToken ct)
    {
        var url = $"https://huggingface.co/api/models/{repoId}/tree/main?recursive=1";
        using var response = await _httpClient.GetAsync(url, ct);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Failed to inspect repo tree {repoId}: {(int)response.StatusCode} {response.ReasonPhrase}");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<string>();
        }

        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            if (!entry.TryGetProperty("type", out var typeProp) ||
                !string.Equals(typeProp.GetString(), "file", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            if (!entry.TryGetProperty("path", out var pathProp))
            {
                continue;
            }
            var path = pathProp.GetString();
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }
            if (path.Contains("/.", StringComparison.Ordinal) || path.StartsWith(".", StringComparison.Ordinal))
            {
                continue;
            }
            if (path.EndsWith(".lock", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (prefixes.Any(prefix => path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            {
                files.Add(path);
            }
        }

        if (rootRequiredFiles is not null)
        {
            foreach (var file in rootRequiredFiles)
            {
                files.Add(file);
            }
        }

        return files.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static (string Owner, string Name) ParseRepo(string repoId)
    {
        var tokens = repoId.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2)
        {
            throw new InvalidOperationException($"Invalid model repo id: {repoId}");
        }
        return (tokens[0], tokens[1]);
    }

    private static bool IsLegacyChatterboxRepo(string repoId)
    {
        return repoId.Contains("chatterbox", StringComparison.OrdinalIgnoreCase) && !IsChatterboxOnnxRepo(repoId);
    }

    private sealed class DownloadTelemetryState
    {
        public long TotalDownloaded { get; set; }
        public long TotalExpected { get; set; }
        public bool HasUnknownTotals { get; set; }
        public int FilesCompleted { get; set; }
        public int FilesTotal { get; set; }
    }
}
