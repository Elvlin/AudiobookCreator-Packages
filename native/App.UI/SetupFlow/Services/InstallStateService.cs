using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AudiobookCreator.UI.SetupFlow.Models;

namespace AudiobookCreator.UI.SetupFlow.Services;

public sealed class InstallStateService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public async Task SaveAsync(InstalledState state, string installDirectory, CancellationToken ct)
    {
        var defaultsDir = Path.Combine(installDirectory, "defaults");
        Directory.CreateDirectory(defaultsDir);
        var path = Path.Combine(defaultsDir, "install_state.json");
        await using var stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, state, JsonOptions, ct);
    }

    public async Task<InstalledState?> LoadAsync(string installDirectory, CancellationToken ct)
    {
        var path = Path.Combine(installDirectory, "defaults", "install_state.json");
        if (!File.Exists(path))
        {
            return null;
        }

        await using var stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<InstalledState>(stream, JsonOptions, ct);
    }
}
