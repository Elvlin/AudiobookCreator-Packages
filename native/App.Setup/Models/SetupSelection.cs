namespace AudiobookCreator.Setup.Models;

public sealed class SetupSelection
{
    public string InstallMode { get; set; } = "Auto";
    public string ChatterboxBackend { get; set; } = "auto";
    public string KittenBackend { get; set; } = "auto";
    public bool IncludeQwen { get; set; }
    public bool IncludeLocalLlm { get; set; }
    public bool ApiOnly { get; set; }
    public bool EnterApiKeyNow { get; set; }
    public string ApiProvider { get; set; } = "openai";
    public string ApiKey { get; set; } = string.Empty;
    public string InstallDirectory { get; set; } = string.Empty;
}
