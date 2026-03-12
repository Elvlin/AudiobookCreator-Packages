using System.Collections.Generic;

namespace AudiobookCreator.SetupNetFx.Models;

public sealed class PackageManifest
{
    public string Version { get; set; } = string.Empty;
    public string GeneratedUtc { get; set; } = string.Empty;
    public string SourceRelease { get; set; } = string.Empty;
    public List<PackageEntry> Packages { get; set; } = new List<PackageEntry>();
}

public sealed class PackageEntry
{
    public string Id { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public string File { get; set; } = string.Empty;
    public string RelativeUrl { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public long InstallBytes { get; set; }
    public string Sha256 { get; set; } = string.Empty;
    public string InstallTarget { get; set; } = "app_root";
    public List<string> Dependencies { get; set; } = new List<string>();
    public string OriginalFile { get; set; } = string.Empty;
    public long? OriginalSizeBytes { get; set; }
    public List<PackagePart> Parts { get; set; } = new List<PackagePart>();

    public bool IsMultipart => Parts.Count > 0;
}

public sealed class PackagePart
{
    public int Index { get; set; }
    public string File { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string Sha256 { get; set; } = string.Empty;
}
