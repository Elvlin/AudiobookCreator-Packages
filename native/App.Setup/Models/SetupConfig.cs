using System.IO;

namespace AudiobookCreator.Setup.Models;

public sealed class SetupConfig
{
    public string ProductName { get; init; } = "Audiobook Creator";
    public string ReleaseTag { get; init; } = "0.1.0";
    public string ReleaseBaseUrl { get; init; } = "https://github.com/Elvlin/AudiobookCreator-Packages/releases/download/0.1.0/";
    public string ManifestFileName { get; init; } = "packages.json";
    public string ManifestSignatureFileName { get; init; } = "packages.sig";
    public string ShaFileName { get; init; } = "SHA256SUMS.txt";
    public string AppId { get; init; } = "AudiobookCreator";
    public string DefaultInstallDir { get; init; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Audiobook Creator");
    public string DotNetRuntimeUrl { get; init; } = "https://aka.ms/dotnet/7.0/windowsdesktop-runtime-win-x64.exe";
    public string ManifestPublicKeyPemPath { get; init; } = string.Empty;

    public Uri BuildReleaseUri(string relativePath)
    {
        return new Uri(new Uri(ReleaseBaseUrl), relativePath);
    }
}
