namespace AudiobookCreator.SetupNetFx.Models;

public sealed class InstallProgress
{
    public double Percent { get; set; }
    public string Message { get; set; } = string.Empty;
}
