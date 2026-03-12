using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using AudiobookCreator.SetupNetFx.Models;

namespace AudiobookCreator.SetupNetFx.Services;

public sealed class ManifestService
{
    private readonly HttpClient _http;
    private readonly SetupConfig _config;
    private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer
    {
        MaxJsonLength = int.MaxValue
    };

    public ManifestService(HttpClient http, SetupConfig config)
    {
        _http = http;
        _config = config;
    }

    public async Task<PackageManifest> LoadManifestAsync(CancellationToken ct)
    {
        var uri = _config.BuildReleaseUri(_config.ManifestFileName);
        using (var response = await _http.GetAsync(uri, ct))
        {
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            var manifest = _serializer.Deserialize<PackageManifest>(json);
            if (manifest == null)
            {
                throw new InvalidOperationException("Failed to load package manifest.");
            }

            return manifest;
        }
    }
}
