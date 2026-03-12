using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using AudiobookCreator.SetupNetFx.Models;
using Microsoft.Win32;

namespace AudiobookCreator.SetupNetFx.Services;

public sealed class EnvironmentService
{
    public DetectedMachineProfile Detect(string installDirectory, InstalledState? existingState, bool internetReachable)
    {
        var gpuName = "None";
        var gpuVendor = "none";
        var gpuVramGb = 0d;
        var ramGb = 0d;
        var cpuLogical = Environment.ProcessorCount;

        try
        {
            using (var searcher = new ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController"))
            using (var results = searcher.Get())
            {
                foreach (ManagementObject item in results)
                {
                    gpuName = Convert.ToString(item["Name"]) ?? gpuName;
                    var ramBytes = item["AdapterRAM"] != null ? Convert.ToDouble(item["AdapterRAM"]) : 0d;
                    gpuVramGb = Math.Max(gpuVramGb, ramBytes / 1024d / 1024d / 1024d);
                    gpuVendor = DetectVendor(gpuName);
                    if (gpuVendor != "none")
                    {
                        break;
                    }
                }
            }
        }
        catch
        {
        }

        try
        {
            using (var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"))
            using (var results = searcher.Get())
            {
                foreach (ManagementObject item in results)
                {
                    var bytes = item["TotalPhysicalMemory"] != null ? Convert.ToDouble(item["TotalPhysicalMemory"]) : 0d;
                    ramGb = bytes / 1024d / 1024d / 1024d;
                    break;
                }
            }
        }
        catch
        {
        }

        var dotNetInstalled = HasDesktopRuntime7();
        var internetAvailable = internetReachable || HasInternetFallback();
        return new DetectedMachineProfile
        {
            OperatingSystem = Environment.OSVersion.VersionString,
            Is64Bit = Environment.Is64BitOperatingSystem,
            DotNetDesktopRuntimeInstalled = dotNetInstalled,
            DotNetStatus = dotNetInstalled ? ".NET Desktop Runtime 7.x detected" : ".NET Desktop Runtime 7.x missing",
            GpuVendor = gpuVendor,
            GpuName = gpuName,
            RamGb = ramGb,
            GpuVramGb = gpuVramGb,
            CpuLogical = cpuLogical,
            InternetAvailable = internetAvailable,
            InternetStatus = internetAvailable ? "Release feed reachable" : "Release feed not reachable",
            AvailableInstallDriveBytes = GetAvailableBytes(installDirectory),
            DiskStatus = FormatDiskStatus(installDirectory),
            ExistingInstallDetected = existingState != null,
            ExistingInstallStatus = existingState == null ? "No existing install detected" : "Existing install detected (" + existingState.CompletionState + ")"
        };
    }

    public async Task<bool> EnsureDotNetDesktopRuntimeAsync(SetupConfig config, string downloadDir, CancellationToken ct)
    {
        if (HasDesktopRuntime7())
        {
            return true;
        }

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        Directory.CreateDirectory(downloadDir);
        var installerPath = Path.Combine(downloadDir, "windowsdesktop-runtime-7-installer.exe");

        using (var http = new HttpClient())
        using (var response = await http.GetAsync(config.DotNetRuntimeUrl, HttpCompletionOption.ResponseHeadersRead, ct))
        {
            response.EnsureSuccessStatusCode();
            using (var input = await response.Content.ReadAsStreamAsync())
            using (var output = File.Create(installerPath))
            {
                await input.CopyToAsync(output);
            }
        }

        var process = Process.Start(new ProcessStartInfo
        {
            FileName = installerPath,
            Arguments = "/install /quiet /norestart",
            UseShellExecute = true,
            Verb = "runas"
        });

        if (process == null)
        {
            return false;
        }

        await Task.Run(() => process.WaitForExit(), ct);
        return process.ExitCode == 0 && HasDesktopRuntime7();
    }

    public async Task<bool> CanReachReleaseFeedAsync(SetupConfig config, CancellationToken ct)
    {
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (var http = new HttpClient())
            using (var response = await http.GetAsync(config.BuildReleaseUri(config.ManifestFileName), HttpCompletionOption.ResponseHeadersRead, ct))
            {
                return response.IsSuccessStatusCode;
            }
        }
        catch
        {
            return false;
        }
    }

    private static string DetectVendor(string gpuName)
    {
        var value = (gpuName ?? string.Empty).ToLowerInvariant();
        if (value.Contains("nvidia"))
        {
            return "nvidia";
        }

        if (value.Contains("amd") || value.Contains("radeon") || value.Contains("ati"))
        {
            return "amd";
        }

        if (value.Contains("intel"))
        {
            return "intel";
        }

        return "none";
    }

    private static bool HasDesktopRuntime7()
    {
        return HasDesktopRuntimeInRegistry() || HasDesktopRuntimeOnDisk() || HasDesktopRuntimeViaDotNet();
    }

    private static bool HasDesktopRuntimeInRegistry()
    {
        try
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var key = baseKey.OpenSubKey(@"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App"))
                {
                    if (key == null)
                    {
                        continue;
                    }

                    foreach (var name in key.GetSubKeyNames())
                    {
                        if (name.StartsWith("7.", StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
        }
        catch
        {
        }

        return false;
    }

    private static bool HasDesktopRuntimeOnDisk()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (string.IsNullOrWhiteSpace(programFiles))
        {
            return false;
        }

        var root = Path.Combine(programFiles, "dotnet", "shared", "Microsoft.WindowsDesktop.App");
        if (!Directory.Exists(root))
        {
            return false;
        }

        return Directory.EnumerateDirectories(root, "7.*").Any();
    }

    private static bool HasDesktopRuntimeViaDotNet()
    {
        try
        {
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = "--list-runtimes",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            });

            if (process == null)
            {
                return false;
            }

            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit(5000);
            return output.IndexOf("Microsoft.WindowsDesktop.App 7.", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool HasInternetFallback()
    {
        try
        {
            Dns.GetHostEntry("github.com");
            return true;
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
        var gb = GetAvailableBytes(installDirectory) / 1024d / 1024d / 1024d;
        return string.Format("{0:F1} GB free on target drive", gb);
    }
}
