using System;
using System.IO;

namespace AudiobookCreator.SetupNetFx.Models;

public sealed class SetupConfig
{
    public string ProductName { get; set; } = "Audiobook Creator";
    public string ReleaseTag { get; set; } = "0.1.0";
    public string ReleaseBaseUrl { get; set; } = "https://github.com/Elvlin/AudiobookCreator-Packages/releases/download/0.1.0/";
    public string ManifestFileName { get; set; } = "packages.json";
    public string ManifestSignatureFileName { get; set; } = "packages.sig";
    public string ShaFileName { get; set; } = "SHA256SUMS.txt";
    public string AppId { get; set; } = "AudiobookCreator";
    public string DefaultInstallDir { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Audiobook Creator");
    public string DotNetRuntimeUrl { get; set; } = "https://aka.ms/dotnet/7.0/windowsdesktop-runtime-win-x64.exe";
    public string ManifestPublicKeyPemPath { get; set; } = string.Empty;

    public Uri BuildReleaseUri(string relativePath)
    {
        return new Uri(new Uri(ReleaseBaseUrl), relativePath);
    }
}
