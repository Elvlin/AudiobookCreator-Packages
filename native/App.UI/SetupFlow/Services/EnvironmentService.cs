using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using App.Diagnostics;
using AudiobookCreator.UI.SetupFlow.Models;
using Microsoft.Win32;

namespace AudiobookCreator.UI.SetupFlow.Services;

public sealed class EnvironmentService
{
    public DetectedMachineProfile Detect(string installDirectory, InstalledState? existingState = null, bool internetReachable = false)
    {
        var system = SystemProbe.Detect();
        var dotNetInstalled = HasDesktopRuntime7();
        var internetAvailable = internetReachable || HasInternetFallback();
        return new DetectedMachineProfile
        {
            OperatingSystem = Environment.OSVersion.VersionString,
            Is64Bit = Environment.Is64BitOperatingSystem,
            DotNetDesktopRuntimeInstalled = dotNetInstalled,
            DotNetStatus = dotNetInstalled ? ".NET Desktop Runtime 7.x detected" : ".NET Desktop Runtime 7.x missing",
            GpuVendor = system.GpuVendor,
            GpuName = system.GpuName,
            RamGb = system.RamGb,
            GpuVramGb = system.GpuVramGb,
            CpuLogical = system.CpuLogical,
            InternetAvailable = internetAvailable,
            InternetStatus = internetAvailable ? "Release feed reachable" : "Release feed not reachable",
            AvailableInstallDriveBytes = GetAvailableBytes(installDirectory),
            DiskStatus = FormatDiskStatus(installDirectory),
            ExistingInstallDetected = existingState is not null,
            ExistingInstallStatus = existingState is null
                ? "No existing install detected"
                : $"Existing install detected ({existingState.CompletionState})"
        };
    }

    public async Task<bool> EnsureDotNetDesktopRuntimeAsync(SetupConfig config, string downloadDir, CancellationToken ct)
    {
        if (HasDesktopRuntime7())
        {
            return true;
        }

        Directory.CreateDirectory(downloadDir);
        var installerPath = Path.Combine(downloadDir, "windowsdesktop-runtime-7-installer.exe");
        using var http = new HttpClient();
        await using (var input = await http.GetStreamAsync(config.DotNetRuntimeUrl, ct))
        await using (var output = File.Create(installerPath))
        {
            await input.CopyToAsync(output, ct);
        }

        var process = Process.Start(new ProcessStartInfo
        {
            FileName = installerPath,
            Arguments = "/install /quiet /norestart",
            UseShellExecute = true,
            Verb = "runas"
        });

        if (process is null)
        {
            return false;
        }

        await process.WaitForExitAsync(ct);
        return process.ExitCode == 0 && HasDesktopRuntime7();
    }

    public async Task<bool> CanReachReleaseFeedAsync(SetupConfig config, CancellationToken ct)
    {
        try
        {
            using var http = new HttpClient();
            using var response = await http.GetAsync(config.BuildReleaseUri(config.ManifestFileName), HttpCompletionOption.ResponseHeadersRead, ct);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private static bool HasDesktopRuntime7()
    {
        if (HasDesktopRuntimeInRegistry())
        {
            return true;
        }

        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = "--list-runtimes",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            });

            if (process is null)
            {
                return false;
            }

            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit(5000);
            return output.Contains("Microsoft.WindowsDesktop.App 7.", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static bool HasDesktopRuntimeInRegistry()
    {
        try
        {
            var roots = new[]
            {
                RegistryView.Registry64,
                RegistryView.Registry32
            };

            foreach (var view in roots)
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                using var key = baseKey.OpenSubKey(@"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App");
                if (key is null)
                {
                    continue;
                }

                foreach (var name in key.GetValueNames())
                {
                    if (name.StartsWith("7.", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
        }
        catch
        {
            // Fall through to command-based detection.
        }

        return false;
    }

    private static bool HasInternetFallback()
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "nslookup",
                Arguments = "github.com",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            });

            if (process is null)
            {
                return false;
            }

            process.WaitForExit(5000);
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static long GetAvailableBytes(string installDirectory)
    {
        var root = Path.GetPathRoot(Path.GetFullPath(installDirectory));
        if (string.IsNullOrWhiteSpace(root))
        {
            return 0;
        }

        var drive = new DriveInfo(root);
        return drive.AvailableFreeSpace;
    }

    private static string FormatDiskStatus(string installDirectory)
    {
        var bytes = GetAvailableBytes(installDirectory);
        var gb = bytes / 1024d / 1024d / 1024d;
        return $"{gb:F1} GB free on target drive";
    }
}
