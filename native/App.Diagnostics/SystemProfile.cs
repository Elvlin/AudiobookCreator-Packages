namespace App.Diagnostics;

public sealed class SystemProfile
{
    public int CpuLogical { get; init; }
    public double RamGb { get; init; }
    public string GpuName { get; init; } = "None";
    public string GpuVendor { get; init; } = "none";
    public double GpuVramGb { get; init; }
    public bool CudaAvailable { get; init; }
    public bool Fp16Available { get; init; }
    public string CudaStatus { get; init; } = "Unknown";
    public string CudaInstallHint { get; init; } = string.Empty;
}
