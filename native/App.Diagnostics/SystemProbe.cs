using System.Diagnostics;
using System.IO;
using System.Management;
using App.Core.Runtime;

namespace App.Diagnostics;

public static class SystemProbe
{
    public static SystemProfile Detect()
    {
        var profile = new SystemProfile
        {
            CpuLogical = Environment.ProcessorCount,
            RamGb = Math.Round(GC.GetGCMemoryInfo().TotalAvailableMemoryBytes / 1024d / 1024d / 1024d, 1)
        };

        var (gpuName, vramGb) = TryGetGpuFromNvidiaSmi();
        if (string.Equals(gpuName, "None", StringComparison.OrdinalIgnoreCase))
        {
            (gpuName, vramGb) = TryGetGpuFromWmi();
        }
        var gpuVendor = DetectGpuVendor(gpuName);
        var hasNvidiaGpu = !string.Equals(gpuName, "None", StringComparison.OrdinalIgnoreCase) &&
                           gpuName.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase);
        var cuda = DiagnoseCuda();
        return new SystemProfile
        {
            CpuLogical = profile.CpuLogical,
            RamGb = profile.RamGb,
            GpuName = gpuName,
            GpuVendor = gpuVendor,
            GpuVramGb = vramGb,
            CudaAvailable = cuda.Available,
            Fp16Available = cuda.Available && hasNvidiaGpu,
            CudaStatus = cuda.Status,
            CudaInstallHint = cuda.InstallHint
        };
    }

    private static (bool Available, string Status, string InstallHint) DiagnoseCuda()
    {
        var hasDriver = HasCudaDriverDll();
        var hasSmi = ProbeCudaViaNvidiaSmi();
        var (hasRuntime, runtimeSource) = HasCudaRuntimeDlls();

        if (hasSmi && (hasRuntime || hasDriver))
        {
            var source = string.IsNullOrWhiteSpace(runtimeSource) ? "driver/runtime detected" : runtimeSource;
            return (true, $"CUDA ready ({source})", string.Empty);
        }

        if (hasDriver && hasRuntime)
        {
            return (true, $"CUDA runtime detected ({runtimeSource})", string.Empty);
        }

        var status = hasDriver
            ? "NVIDIA driver found, but CUDA runtime DLLs were not found in PATH/toolkit/torch runtime locations."
            : "NVIDIA CUDA driver (nvcuda.dll) not detected.";
        var hint =
            "Install NVIDIA driver and CUDA 12 runtime/toolkit: https://www.nvidia.com/Download/index.aspx  |  https://developer.nvidia.com/cuda-downloads";
        return (false, status, hint);
    }

    private static (string Name, double VramGb) TryGetGpuFromNvidiaSmi()
    {
        try
        {
            using var proc = new Process();
            proc.StartInfo.FileName = "nvidia-smi";
            proc.StartInfo.Arguments = "--query-gpu=name,memory.total --format=csv,noheader,nounits";
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.Start();

            var line = proc.StandardOutput.ReadLine();
            proc.WaitForExit(2000);
            if (string.IsNullOrWhiteSpace(line))
            {
                return ("None", 0.0);
            }

            var split = line.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (split.Length == 0)
            {
                return ("None", 0.0);
            }

            var name = split[0];
            var vram = 0.0;
            if (split.Length > 1 && double.TryParse(split[1], out var memMb))
            {
                vram = Math.Round(memMb / 1024d, 1);
            }

            return (name, vram);
        }
        catch
        {
            return ("None", 0.0);
        }
    }

    private static (string Name, double VramGb) TryGetGpuFromWmi()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController");
            var candidates = new List<(string Name, double VramGb, int Priority)>();
            foreach (var obj in searcher.Get())
            {
                var name = obj["Name"]?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var ramBytes = 0.0;
                if (double.TryParse(obj["AdapterRAM"]?.ToString(), out var v))
                {
                    ramBytes = v;
                }
                var vram = ramBytes > 0 ? Math.Round(ramBytes / 1024d / 1024d / 1024d, 1) : 0.0;
                candidates.Add((name, vram, GetVendorPriority(name)));
            }

            var selected = candidates
                .OrderByDescending(x => x.Priority)
                .ThenByDescending(x => x.VramGb)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(selected.Name))
            {
                return (selected.Name, selected.VramGb);
            }
        }
        catch
        {
            // Ignore probe failure.
        }

        return ("None", 0.0);
    }

    private static int GetVendorPriority(string gpuName)
    {
        var vendor = DetectGpuVendor(gpuName);
        return vendor switch
        {
            "nvidia" => 4,
            "amd" => 3,
            "intel" => 2,
            "other" => 1,
            _ => 0
        };
    }

    private static string DetectGpuVendor(string gpuName)
    {
        if (string.IsNullOrWhiteSpace(gpuName) || string.Equals(gpuName, "None", StringComparison.OrdinalIgnoreCase))
        {
            return "none";
        }

        if (gpuName.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase))
        {
            return "nvidia";
        }

        if (gpuName.Contains("AMD", StringComparison.OrdinalIgnoreCase) ||
            gpuName.Contains("Radeon", StringComparison.OrdinalIgnoreCase) ||
            gpuName.Contains("ATI", StringComparison.OrdinalIgnoreCase))
        {
            return "amd";
        }

        if (gpuName.Contains("Intel", StringComparison.OrdinalIgnoreCase))
        {
            return "intel";
        }

        return "other";
    }

    private static bool HasCudaDriverDll()
    {
        try
        {
            var sys = Environment.GetFolderPath(Environment.SpecialFolder.System);
            if (!string.IsNullOrWhiteSpace(sys))
            {
                var dll = Path.Combine(sys, "nvcuda.dll");
                if (File.Exists(dll))
                {
                    return true;
                }
            }

            var windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrWhiteSpace(windir))
            {
                var sys32Dll = Path.Combine(windir, "System32", "nvcuda.dll");
                if (File.Exists(sys32Dll))
                {
                    return true;
                }

                var wow64Dll = Path.Combine(windir, "SysWOW64", "nvcuda.dll");
                if (File.Exists(wow64Dll))
                {
                    return true;
                }
            }

            var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var dir in path.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                try
                {
                    if (File.Exists(Path.Combine(dir, "nvcuda.dll")))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Continue scanning.
                }
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static (bool Found, string Source) HasCudaRuntimeDlls()
    {
        var candidates = BuildCudaPathCandidates();
        foreach (var dir in candidates)
        {
            var cublas = Path.Combine(dir, "cublasLt64_12.dll");
            var cudart = Path.Combine(dir, "cudart64_12.dll");
            if (File.Exists(cublas) && File.Exists(cudart))
            {
                return (true, dir);
            }
        }

        return (false, string.Empty);
    }

    private static IReadOnlyList<string> BuildCudaPathCandidates()
    {
        var candidates = new List<string>();

        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var dir in path.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            candidates.Add(dir);
        }

        var cudaPath = Environment.GetEnvironmentVariable("CUDA_PATH");
        if (!string.IsNullOrWhiteSpace(cudaPath))
        {
            candidates.Add(Path.Combine(cudaPath, "bin"));
        }

        var appRoot = RuntimePathResolver.AppRoot;
        candidates.Add(Path.Combine(appRoot, ".venv", "Lib", "site-packages", "torch", "lib"));

        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var pythonRoot = Path.Combine(localAppData, "Programs", "Python");
        if (Directory.Exists(pythonRoot))
        {
            foreach (var pyDir in Directory.GetDirectories(pythonRoot))
            {
                candidates.Add(Path.Combine(pyDir, "Lib", "site-packages", "torch", "lib"));
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

        return candidates
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(NormalizePathSafe)
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizePathSafe(string path)
    {
        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
    }

    private static bool ProbeCudaViaNvidiaSmi()
    {
        try
        {
            using var proc = new Process();
            proc.StartInfo.FileName = "nvidia-smi";
            proc.StartInfo.Arguments = "--query-gpu=driver_version --format=csv,noheader";
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.Start();

            var line = proc.StandardOutput.ReadLine();
            proc.WaitForExit(2000);
            return proc.ExitCode == 0 && !string.IsNullOrWhiteSpace(line);
        }
        catch
        {
            return false;
        }
    }
}
