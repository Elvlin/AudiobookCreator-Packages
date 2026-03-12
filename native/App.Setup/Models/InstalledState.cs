namespace AudiobookCreator.Setup.Models;

public sealed class InstalledState
{
    public string ProductName { get; set; } = "Audiobook Creator";
    public string InstallPath { get; set; } = string.Empty;
    public string CompletionState { get; set; } = "unknown";
    public DateTimeOffset LastUpdatedUtc { get; set; } = DateTimeOffset.UtcNow;
    public List<InstalledPackage> Packages { get; set; } = new();
    public string SetupVersion { get; set; } = "0.1.0";
}

public sealed class InstalledPackage
{
    public string Id { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public DateTimeOffset InstalledUtc { get; set; } = DateTimeOffset.UtcNow;
}
