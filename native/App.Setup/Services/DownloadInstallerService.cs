using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Diagnostics;
using AudiobookCreator.Setup.Models;

namespace AudiobookCreator.Setup.Services;

public sealed class DownloadInstallerService
{
    private readonly HttpClient _http;
    private readonly SetupConfig _config;

    public DownloadInstallerService(HttpClient http, SetupConfig config)
    {
        _http = http;
        _config = config;
    }

    public async Task InstallAsync(
        PackageManifest manifest,
        ResolvedInstallPlan plan,
        string installDirectory,
        IProgress<InstallProgress> progress,
        CancellationToken ct)
    {
        var orderedPackages = plan.PackageIds
            .Select(id => manifest.Packages.First(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        var fullInstallDirectory = Path.GetFullPath(installDirectory);
        var installParent = Directory.GetParent(fullInstallDirectory)?.FullName
            ?? throw new InvalidOperationException("Install folder must have a valid parent directory.");
        Directory.CreateDirectory(installParent);

        EnsureInstallTargetReady(fullInstallDirectory);

        var cacheRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            $"{_config.AppId}Setup",
            "cache",
            _config.ReleaseTag);
        var packageCacheRoot = Path.Combine(cacheRoot, "packages");
        var archiveCacheRoot = Path.Combine(cacheRoot, "archives");
        var stageRoot = Path.Combine(installParent, $".{Path.GetFileName(fullInstallDirectory)}.stage.{Guid.NewGuid():N}");
        var backupRoot = Path.Combine(installParent, $".{Path.GetFileName(fullInstallDirectory)}.backup");

        Directory.CreateDirectory(packageCacheRoot);
        Directory.CreateDirectory(archiveCacheRoot);
        Directory.CreateDirectory(stageRoot);

        try
        {
            for (var i = 0; i < orderedPackages.Count; i++)
            {
                var package = orderedPackages[i];
                progress.Report(new InstallProgress
                {
                    Percent = i * 100d / orderedPackages.Count,
                    Message = $"Downloading {FriendlyName(package.Id)}..."
                });

                var archivePath = await GetVerifiedArchiveAsync(package, packageCacheRoot, archiveCacheRoot, progress, ct);

                progress.Report(new InstallProgress
                {
                    Percent = (i + 0.35) * 100d / orderedPackages.Count,
                    Message = $"Extracting {FriendlyName(package.Id)} into staging..."
                });

                ExtractArchiveToStage(package, archivePath, stageRoot);
            }

            ValidateStageInstall(stageRoot, orderedPackages);

            progress.Report(new InstallProgress
            {
                Percent = 96,
                Message = "Committing staged install..."
            });

            CommitStage(stageRoot, fullInstallDirectory, backupRoot);
        }
        catch
        {
            if (Directory.Exists(stageRoot))
            {
                Directory.Delete(stageRoot, true);
            }

            throw;
        }
        finally
        {
            if (Directory.Exists(stageRoot))
            {
                Directory.Delete(stageRoot, true);
            }
        }

        progress.Report(new InstallProgress
        {
            Percent = 100,
            Message = "Installation completed."
        });
    }

    private async Task<string> GetVerifiedArchiveAsync(
        PackageEntry package,
        string packageCacheRoot,
        string archiveCacheRoot,
        IProgress<InstallProgress> progress,
        CancellationToken ct)
    {
        if (!package.IsMultipart)
        {
            var archivePath = Path.Combine(packageCacheRoot, package.File!);
            await DownloadFileAsync(_config.BuildReleaseUri(package.RelativeUrl ?? package.File!), archivePath, package.Sha256!, ct);
            return archivePath;
        }

        foreach (var part in package.Parts.OrderBy(x => x.Index))
        {
            var partPath = Path.Combine(packageCacheRoot, part.File);
            progress.Report(new InstallProgress
            {
                Percent = 0,
                Message = $"Downloading part {part.Index} of {package.Parts.Count} for {FriendlyName(package.Id)}..."
            });
            await DownloadFileAsync(_config.BuildReleaseUri(part.File), partPath, part.Sha256, ct);
        }

        var reassembledPath = Path.Combine(archiveCacheRoot, package.OriginalFile!);
        var needsAssembly = !File.Exists(reassembledPath);
        if (!needsAssembly && package.OriginalSizeBytes is long expectedSize)
        {
            needsAssembly = new FileInfo(reassembledPath).Length != expectedSize;
        }

        if (needsAssembly)
        {
            await using var output = File.Create(reassembledPath);
            foreach (var part in package.Parts.OrderBy(x => x.Index))
            {
                var partPath = Path.Combine(packageCacheRoot, part.File);
                await using var input = File.OpenRead(partPath);
                await input.CopyToAsync(output, ct);
            }
        }

        if (package.OriginalSizeBytes is long originalSize && new FileInfo(reassembledPath).Length != originalSize)
        {
            throw new InvalidOperationException($"Multipart reassembly failed for {package.OriginalFile}.");
        }

        return reassembledPath;
    }

    private static void ExtractArchiveToStage(PackageEntry package, string archivePath, string stageRoot)
    {
        var targetRoot = package.InstallTarget switch
        {
            "app_root" => stageRoot,
            _ => throw new InvalidOperationException($"Unsupported install target '{package.InstallTarget}'.")
        };

        Directory.CreateDirectory(targetRoot);
        var tempExtractRoot = Path.Combine(Path.GetTempPath(), "AudiobookCreatorSetup", "extract", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempExtractRoot);
        try
        {
            try
            {
                ZipFile.ExtractToDirectory(archivePath, tempExtractRoot, true);
            }
            catch (InvalidDataException)
            {
                if (Directory.Exists(tempExtractRoot))
                {
                    Directory.Delete(tempExtractRoot, true);
                }

                Directory.CreateDirectory(tempExtractRoot);
                ExtractArchiveWithTar(archivePath, tempExtractRoot);
            }

            MergeDirectory(tempExtractRoot, targetRoot);
        }
        finally
        {
            if (Directory.Exists(tempExtractRoot))
            {
                Directory.Delete(tempExtractRoot, true);
            }
        }
    }

    private static void ExtractArchiveWithTar(string archivePath, string destinationDirectory)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "tar",
            WorkingDirectory = destinationDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-xf");
        psi.ArgumentList.Add(archivePath);
        psi.ArgumentList.Add("-C");
        psi.ArgumentList.Add(destinationDirectory);

        using var process = Process.Start(psi);
        if (process is null)
        {
            throw new InvalidOperationException("Failed to start tar for archive extraction.");
        }

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            var detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
            throw new InvalidOperationException($"tar extraction failed for {Path.GetFileName(archivePath)}: {detail}");
        }
    }

    private static void MergeDirectory(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);

        foreach (var directory in Directory.EnumerateDirectories(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(sourceDirectory, directory);
            Directory.CreateDirectory(Path.Combine(destinationDirectory, relative));
        }

        foreach (var file in Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(sourceDirectory, file);
            var destinationPath = Path.Combine(destinationDirectory, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
            File.Copy(file, destinationPath, true);
        }
    }

    private static void ValidateStageInstall(string stageRoot, IReadOnlyCollection<PackageEntry> packages)
    {
        if (!packages.Any(pkg => string.Equals(pkg.Id, "app_core_no_audiox", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var requiredFiles = new[]
        {
            "AudiobookCreator.exe",
            Path.Combine("defaults", "app_config.json")
        };

        foreach (var relative in requiredFiles)
        {
            var fullPath = Path.Combine(stageRoot, relative);
            if (!File.Exists(fullPath))
            {
                throw new InvalidOperationException($"Staged install is missing required file: {relative}");
            }
        }
    }

    private static void CommitStage(string stageRoot, string installDirectory, string backupRoot)
    {
        if (Directory.Exists(backupRoot))
        {
            Directory.Delete(backupRoot, true);
        }

        var hadExistingInstall = Directory.Exists(installDirectory);
        if (hadExistingInstall)
        {
            Directory.Move(installDirectory, backupRoot);
        }

        try
        {
            Directory.Move(stageRoot, installDirectory);
            if (Directory.Exists(backupRoot))
            {
                Directory.Delete(backupRoot, true);
            }
        }
        catch
        {
            if (!Directory.Exists(installDirectory) && Directory.Exists(backupRoot))
            {
                Directory.Move(backupRoot, installDirectory);
            }

            throw;
        }
    }

    private static void EnsureInstallTargetReady(string installDirectory)
    {
        if (!Directory.Exists(installDirectory))
        {
            return;
        }

        var entries = Directory.EnumerateFileSystemEntries(installDirectory).ToList();
        if (entries.Count == 0)
        {
            return;
        }

        var installStatePath = Path.Combine(installDirectory, "defaults", "install_state.json");
        if (File.Exists(installStatePath))
        {
            return;
        }

        throw new InvalidOperationException("Install folder already contains files that are not a managed Audiobook Creator install. Choose an empty folder or an existing Audiobook Creator folder.");
    }

    private async Task DownloadFileAsync(Uri uri, string destination, string expectedSha, CancellationToken ct)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);

        if (File.Exists(destination))
        {
            var existing = await ComputeSha256Async(destination, ct);
            if (string.Equals(existing, expectedSha, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            File.Delete(destination);
        }

        await using (var input = await _http.GetStreamAsync(uri, ct))
        await using (var output = File.Create(destination))
        {
            await input.CopyToAsync(output, ct);
        }

        var sha = await ComputeSha256Async(destination, ct);
        if (!string.Equals(sha, expectedSha, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"SHA verification failed for {Path.GetFileName(destination)}");
        }
    }

    private static async Task<string> ComputeSha256Async(string path, CancellationToken ct)
    {
        await using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        var hash = await sha.ComputeHashAsync(stream, ct);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static string FriendlyName(string packageId)
    {
        return packageId
            .Replace("_", " ", StringComparison.Ordinal)
            .Replace("qwen", "Qwen", StringComparison.OrdinalIgnoreCase)
            .Replace("onnx", "ONNX", StringComparison.OrdinalIgnoreCase);
    }
}
