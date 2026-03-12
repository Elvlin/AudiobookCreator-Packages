using System;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AudiobookCreator.UI.SetupFlow.Models;

namespace AudiobookCreator.UI.SetupFlow.Services;

public sealed class ManifestService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly HttpClient _http;
    private readonly SetupConfig _config;

    public ManifestService(HttpClient http, SetupConfig config)
    {
        _http = http;
        _config = config;
    }

    public async Task<PackageManifest> LoadManifestAsync(CancellationToken ct)
    {
        var uri = _config.BuildReleaseUri(_config.ManifestFileName);
        var manifestBytes = await _http.GetByteArrayAsync(uri, ct);
        await VerifySignatureIfConfiguredAsync(manifestBytes, ct);
        using var stream = new MemoryStream(manifestBytes);
        var manifest = await JsonSerializer.DeserializeAsync<PackageManifest>(stream, JsonOptions, ct);
        return manifest ?? throw new InvalidOperationException("Failed to load package manifest.");
    }

    private async Task VerifySignatureIfConfiguredAsync(byte[] manifestBytes, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(_config.ManifestPublicKeyPemPath) || !File.Exists(_config.ManifestPublicKeyPemPath))
        {
            return;
        }

        var signatureUri = _config.BuildReleaseUri(_config.ManifestSignatureFileName);
        byte[] signatureBytes;
        try
        {
            signatureBytes = await _http.GetByteArrayAsync(signatureUri, ct);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Manifest signature is required but could not be downloaded.", ex);
        }

        var publicKeyPem = await File.ReadAllTextAsync(_config.ManifestPublicKeyPemPath, ct);
        using var rsa = RSA.Create();
        rsa.ImportFromPem(publicKeyPem);
        if (!rsa.VerifyData(manifestBytes, signatureBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
        {
            throw new InvalidOperationException("Package manifest signature verification failed.");
        }
    }
}
