using System;

namespace AudiobookCreator.SetupNetFx.Models;

public sealed class DetectedMachineProfile
{
    public string OperatingSystem { get; set; } = Environment.OSVersion.VersionString;
    public bool Is64Bit { get; set; } = Environment.Is64BitOperatingSystem;
    public bool DotNetDesktopRuntimeInstalled { get; set; }
    public string DotNetStatus { get; set; } = "Unknown";
    public string GpuVendor { get; set; } = "none";
    public string GpuName { get; set; } = "None";
    public double RamGb { get; set; }
    public double GpuVramGb { get; set; }
    public int CpuLogical { get; set; }
    public bool InternetAvailable { get; set; }
    public string InternetStatus { get; set; } = "Unknown";
    public string DiskStatus { get; set; } = "Unknown";
    public long AvailableInstallDriveBytes { get; set; }
    public bool ExistingInstallDetected { get; set; }
    public string ExistingInstallStatus { get; set; } = "No existing install detected";
}
