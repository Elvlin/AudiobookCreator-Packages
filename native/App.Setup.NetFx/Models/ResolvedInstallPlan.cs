using System.Collections.Generic;

namespace AudiobookCreator.SetupNetFx.Models;

public sealed class ResolvedInstallPlan
{
    public string Mode { get; set; } = "Auto";
    public List<string> PackageIds { get; set; } = new List<string>();
    public long DownloadBytes { get; set; }
    public long FinalInstallBytes { get; set; }
    public long WorkingBytesRequired { get; set; }
    public List<string> Warnings { get; set; } = new List<string>();
    public string Summary { get; set; } = string.Empty;
}
