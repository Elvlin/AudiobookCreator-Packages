using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;

namespace AudiobookCreator.UI;

internal static class ReleaseIntegrityVerifier
{
    private const string ManifestFileName = "integrity_manifest.json";
    private const string StateFileName = "integrity_state.json";
    private static readonly string[] QuickCriticalFiles =
    {
        "AudiobookCreator.exe",
        "AudiobookCreator.dll",
        "App.Core.dll",
        "App.Storage.dll",
        "App.Inference.dll",
        "App.Diagnostics.dll",
        "native_tts_engine.dll"
    };

    public static IntegrityCheckResult VerifyForStartup(string appRoot, bool strict)
    {
        var quick = BuildQuickStamp(appRoot);
        if (!quick.Success)
        {
            return quick;
        }

        var manifestPath = Path.Combine(appRoot, ManifestFileName);
        if (!File.Exists(manifestPath))
        {
            return strict
                ? IntegrityCheckResult.Fail("Integrity manifest is missing.")
                : IntegrityCheckResult.Ok("Integrity manifest missing (non-strict mode).");
        }

        var manifestHash = ComputeSha256(manifestPath);
        var statePath = GetStatePath();
        var appRootStamp = ComputeStringSha256(appRoot.ToLowerInvariant());
        var state = TryLoadState(statePath);
        var stateMatches = state is not null &&
                           string.Equals(state.AppRootSha256, appRootStamp, StringComparison.OrdinalIgnoreCase) &&
                           string.Equals(state.QuickStamp, quick.Message, StringComparison.OrdinalIgnoreCase) &&
                           string.Equals(state.ManifestSha256, manifestHash, StringComparison.OrdinalIgnoreCase);

        if (stateMatches)
        {
            return IntegrityCheckResult.Ok("Quick integrity check passed.");
        }

        var full = VerifyRelease(appRoot, strict);
        if (!full.Success)
        {
            return full;
        }

        SaveState(statePath, new IntegrityState
        {
            AppRootSha256 = appRootStamp,
            QuickStamp = quick.Message,
            ManifestSha256 = manifestHash,
            VerifiedUtc = DateTime.UtcNow.ToString("O")
        });
        return IntegrityCheckResult.Ok("Full integrity verification completed.");
    }

    public static IntegrityCheckResult VerifyRelease(string appRoot, bool strict)
    {
        var logPath = Path.Combine(appRoot, "integrity_check.log");
        void Log(string message)
        {
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.UtcNow:O}] {message}{Environment.NewLine}");
            }
            catch
            {
                // ignore logging failures
            }
        }

        try
        {
            var manifestPath = Path.Combine(appRoot, ManifestFileName);
            Log($"strict={strict}; appRoot={appRoot}; manifest={manifestPath}");
            if (!File.Exists(manifestPath))
            {
                Log("manifest missing");
                return strict
                    ? IntegrityCheckResult.Fail("Integrity manifest is missing.")
                    : IntegrityCheckResult.Ok("Integrity manifest not present (non-strict mode).");
            }

            var json = File.ReadAllText(manifestPath);
            var manifest = JsonSerializer.Deserialize<IntegrityManifest>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (manifest is null || manifest.Files is null || manifest.Files.Count == 0)
            {
                Log("manifest invalid/empty");
                return IntegrityCheckResult.Fail("Integrity manifest is invalid or empty.");
            }
            Log($"manifest entries={manifest.Files.Count}");

            var byPath = manifest.Files
                .GroupBy(x => NormalizeRelPath(x.Path))
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            Log($"unique entries={byPath.Count}");

            foreach (var entry in byPath.Values)
            {
                var absPath = Path.Combine(appRoot, entry.Path.Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(absPath))
                {
                    Log($"missing file: {entry.Path}");
                    return IntegrityCheckResult.Fail($"Missing file: {entry.Path}");
                }

                var info = new FileInfo(absPath);
                if (info.Length != entry.Size)
                {
                    Log($"size mismatch: {entry.Path}; manifest={entry.Size}; actual={info.Length}");
                    return IntegrityCheckResult.Fail($"File size mismatch: {entry.Path}");
                }

                var actualHash = ComputeSha256(absPath);
                if (!string.Equals(actualHash, entry.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    Log($"hash mismatch: {entry.Path}");
                    return IntegrityCheckResult.Fail($"File hash mismatch: {entry.Path}");
                }
            }

            Log("integrity ok");
            return IntegrityCheckResult.Ok($"Integrity verified ({byPath.Count} files).");
        }
        catch (Exception ex)
        {
            Log("exception: " + ex);
            return IntegrityCheckResult.Fail("Integrity check failed: " + ex.Message);
        }
    }

    private static string NormalizeRelPath(string path)
    {
        return (path ?? string.Empty).Replace('\\', '/').Trim();
    }

    private static string ComputeSha256(string filePath)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(filePath);
        var hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash);
    }

    private static string ComputeStringSha256(string value)
    {
        using var sha = SHA256.Create();
        var bytes = System.Text.Encoding.UTF8.GetBytes(value);
        var hash = sha.ComputeHash(bytes);
        return Convert.ToHexString(hash);
    }

    private static string GetStatePath()
    {
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AudiobookCreator");
        return Path.Combine(root, StateFileName);
    }

    private static IntegrityCheckResult BuildQuickStamp(string appRoot)
    {
        try
        {
            using var sha = SHA256.Create();
            foreach (var rel in QuickCriticalFiles)
            {
                var full = Path.Combine(appRoot, rel);
                if (!File.Exists(full))
                {
                    return IntegrityCheckResult.Fail($"Critical runtime file missing: {rel}");
                }

                var hash = ComputeSha256(full);
                var info = new FileInfo(full);
                var line = $"{rel}|{info.Length}|{hash}\n";
                var bytes = System.Text.Encoding.UTF8.GetBytes(line);
                sha.TransformBlock(bytes, 0, bytes.Length, null, 0);
            }

            sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return IntegrityCheckResult.Ok(Convert.ToHexString(sha.Hash ?? Array.Empty<byte>()));
        }
        catch (Exception ex)
        {
            return IntegrityCheckResult.Fail("Quick integrity check failed: " + ex.Message);
        }
    }

    private static IntegrityState? TryLoadState(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<IntegrityState>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch
        {
            return null;
        }
    }

    private static void SaveState(string path, IntegrityState state)
    {
        try
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }
            var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch
        {
            // ignore state write issues; verification is already complete
        }
    }

    internal sealed class IntegrityManifest
    {
        public int Version { get; set; }
        public string GeneratedUtc { get; set; } = string.Empty;
        public List<IntegrityFileEntry> Files { get; set; } = new();
    }

    internal sealed class IntegrityFileEntry
    {
        public string Path { get; set; } = string.Empty;
        public long Size { get; set; }
        public string Sha256 { get; set; } = string.Empty;
    }

    internal sealed class IntegrityState
    {
        public string AppRootSha256 { get; set; } = string.Empty;
        public string QuickStamp { get; set; } = string.Empty;
        public string ManifestSha256 { get; set; } = string.Empty;
        public string VerifiedUtc { get; set; } = string.Empty;
    }
}

internal readonly struct IntegrityCheckResult
{
    private IntegrityCheckResult(bool success, string message)
    {
        Success = success;
        Message = message;
    }

    public bool Success { get; }
    public string Message { get; }

    public static IntegrityCheckResult Ok(string message) => new(true, message);
    public static IntegrityCheckResult Fail(string message) => new(false, message);
}
