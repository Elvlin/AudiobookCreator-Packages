using System.Collections.Generic;

namespace AudiobookCreator.UI.SetupFlow.Models;

public sealed class ResolvedInstallPlan
{
    public string Mode { get; set; } = "Auto";
    public string Summary { get; set; } = string.Empty;
    public List<string> PackageIds { get; set; } = new();
    public long DownloadBytes { get; set; }
    public long FinalInstallBytes { get; set; }
    public long WorkingBytesRequired { get; set; }
    public List<string> Warnings { get; set; } = new();
}
