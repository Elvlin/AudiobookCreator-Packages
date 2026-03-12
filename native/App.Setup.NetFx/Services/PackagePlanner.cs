using System;
using System.Collections.Generic;
using System.Linq;
using AudiobookCreator.SetupNetFx.Models;

namespace AudiobookCreator.SetupNetFx.Services;

public sealed class PackagePlanner
{
    public ResolvedInstallPlan Resolve(PackageManifest manifest, SetupSelection selection, DetectedMachineProfile machine)
    {
        var packageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<string>();

        if (selection.ApiOnly)
        {
            AddWithDependencies(packageIds, manifest, "app_core_no_audiox");
        }
        else if (string.Equals(selection.InstallMode, "Auto", StringComparison.OrdinalIgnoreCase))
        {
            ResolveAuto(manifest, selection, machine, packageIds, warnings);
        }
        else
        {
            ResolveCustom(manifest, selection, machine, packageIds, warnings);
        }

        var totalBytes = packageIds
            .Select(id => manifest.Packages.First(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var downloadBytes = totalBytes.Sum(GetDownloadBytes);
        var finalInstallBytes = totalBytes.Sum(GetInstallBytes);
        var workingBytesRequired = downloadBytes + totalBytes.Sum(GetArchiveBytes) + finalInstallBytes;
        if (machine.ExistingInstallDetected)
        {
            workingBytesRequired += finalInstallBytes;
        }

        return new ResolvedInstallPlan
        {
            Mode = selection.InstallMode,
            PackageIds = packageIds.OrderBy(x => x).ToList(),
            DownloadBytes = downloadBytes,
            FinalInstallBytes = finalInstallBytes,
            WorkingBytesRequired = workingBytesRequired,
            Warnings = warnings,
            Summary = BuildSummary(selection, machine, packageIds.Count, downloadBytes, finalInstallBytes, workingBytesRequired)
        };
    }

    private static void ResolveAuto(
        PackageManifest manifest,
        SetupSelection selection,
        DetectedMachineProfile machine,
        ISet<string> packageIds,
        List<string> warnings)
    {
        AddWithDependencies(packageIds, manifest, "app_core_no_audiox");

        if (machine.RamGb < 8 || machine.AvailableInstallDriveBytes < 4L * 1024 * 1024 * 1024)
        {
            warnings.Add("PC is below the recommended spec for local models. Auto mode will install the app shell only.");
            selection.ApiOnly = true;
            return;
        }

        var vendor = machine.GpuVendor;
        if (string.Equals(vendor, "nvidia", StringComparison.OrdinalIgnoreCase))
        {
            AddWithDependencies(packageIds, manifest, "model_chatterbox_onnx");
            AddWithDependencies(packageIds, manifest, "model_kitten");
            AddWithDependencies(packageIds, manifest, "tools_basic");
            selection.ChatterboxBackend = "onnx";
            selection.KittenBackend = "onnx";
        }
        else if (string.Equals(vendor, "amd", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(vendor, "intel", StringComparison.OrdinalIgnoreCase))
        {
            AddWithDependencies(packageIds, manifest, "model_chatterbox_python");
            AddWithDependencies(packageIds, manifest, "python_chatterbox");
            AddWithDependencies(packageIds, manifest, "model_kitten");
            AddWithDependencies(packageIds, manifest, "python_kitten");
            AddWithDependencies(packageIds, manifest, "tools_basic");
            selection.ChatterboxBackend = "python";
            selection.KittenBackend = "python";
        }
        else
        {
            warnings.Add("No supported GPU vendor detected. Auto mode will install the app shell only.");
            selection.ApiOnly = true;
        }
    }

    private static void ResolveCustom(
        PackageManifest manifest,
        SetupSelection selection,
        DetectedMachineProfile machine,
        ISet<string> packageIds,
        List<string> warnings)
    {
        AddWithDependencies(packageIds, manifest, "app_core_no_audiox");

        ResolveChatterbox(manifest, selection.ChatterboxBackend, machine, packageIds, warnings);
        ResolveKitten(manifest, selection.KittenBackend, machine, packageIds, warnings);

        if (selection.IncludeQwen)
        {
            if (machine.RamGb < 24)
            {
                warnings.Add("Qwen is likely not suitable on this PC. Recommended RAM is 24 GB or higher.");
            }

            AddWithDependencies(packageIds, manifest, "python_qwen");
            AddWithDependencies(packageIds, manifest, "tools_qwen_worker");
            AddWithDependencies(packageIds, manifest, "model_qwen_onnx");
            AddWithDependencies(packageIds, manifest, "model_qwen_dll");
            AddWithDependencies(packageIds, manifest, "model_qwen_python_cache");
        }

        if (selection.IncludeLocalLlm)
        {
            if (machine.RamGb < 16)
            {
                warnings.Add("Local LLM prep is likely not suitable on this PC. Recommended RAM is 16 GB or higher.");
            }

            AddWithDependencies(packageIds, manifest, "tools_llm");
            AddWithDependencies(packageIds, manifest, "model_llm");
        }
    }

    private static void ResolveChatterbox(
        PackageManifest manifest,
        string backend,
        DetectedMachineProfile machine,
        ISet<string> packageIds,
        List<string> warnings)
    {
        if (string.Equals(backend, "none", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (string.Equals(backend, "onnx", StringComparison.OrdinalIgnoreCase))
        {
            if (!string.Equals(machine.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase))
            {
                warnings.Add("Chatterbox ONNX is likely not suitable on this PC. NVIDIA is the recommended path.");
            }

            AddWithDependencies(packageIds, manifest, "model_chatterbox_onnx");
            return;
        }

        if (machine.RamGb < 12)
        {
            warnings.Add("Chatterbox Python is likely not suitable on this PC. Recommended RAM is 12 GB or higher.");
        }

        AddWithDependencies(packageIds, manifest, "python_chatterbox");
        AddWithDependencies(packageIds, manifest, "model_chatterbox_python");
    }

    private static void ResolveKitten(
        PackageManifest manifest,
        string backend,
        DetectedMachineProfile machine,
        ISet<string> packageIds,
        List<string> warnings)
    {
        if (string.Equals(backend, "none", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        AddWithDependencies(packageIds, manifest, "tools_basic");

        if (string.Equals(backend, "onnx", StringComparison.OrdinalIgnoreCase))
        {
            if (!string.Equals(machine.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase))
            {
                warnings.Add("Kitten ONNX is likely not suitable on this PC. NVIDIA is the recommended path.");
            }

            AddWithDependencies(packageIds, manifest, "model_kitten");
            return;
        }

        if (machine.RamGb < 8)
        {
            warnings.Add("Kitten Python is likely not suitable on this PC. Recommended RAM is 8 GB or higher.");
        }

        AddWithDependencies(packageIds, manifest, "python_kitten");
        AddWithDependencies(packageIds, manifest, "model_kitten");
    }

    private static string BuildSummary(SetupSelection selection, DetectedMachineProfile machine, int count, long downloadBytes, long finalInstallBytes, long workingBytesRequired)
    {
        var downloadGb = downloadBytes / 1024d / 1024d / 1024d;
        var installGb = finalInstallBytes / 1024d / 1024d / 1024d;
        var workingGb = workingBytesRequired / 1024d / 1024d / 1024d;
        return $"{selection.InstallMode} mode on {machine.GpuVendor.ToUpperInvariant()} detected PC. {count} packages selected, {downloadGb:F2} GB download, {installGb:F2} GB installed footprint, {workingGb:F2} GB temporary working space required.";
    }

    private static long GetDownloadBytes(PackageEntry package)
    {
        return package.IsMultipart
            ? package.Parts.Sum(part => part.SizeBytes)
            : package.SizeBytes;
    }

    private static long GetArchiveBytes(PackageEntry package)
    {
        return package.OriginalSizeBytes ?? package.SizeBytes;
    }

    private static long GetInstallBytes(PackageEntry package)
    {
        if (package.InstallBytes > 0)
        {
            return package.InstallBytes;
        }

        return package.OriginalSizeBytes ?? package.SizeBytes;
    }

    private static void AddWithDependencies(ISet<string> target, PackageManifest manifest, string id)
    {
        var package = manifest.Packages.FirstOrDefault(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
        if (package is null || !target.Add(package.Id))
        {
            return;
        }

        foreach (var dependency in package.Dependencies)
        {
            AddWithDependencies(target, manifest, dependency);
        }
    }
}
