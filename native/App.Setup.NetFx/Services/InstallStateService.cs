using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using AudiobookCreator.SetupNetFx.Models;

namespace AudiobookCreator.SetupNetFx.Services;

public sealed class InstallStateService
{
    private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer
    {
        MaxJsonLength = int.MaxValue
    };

    public Task SaveAsync(InstalledState state, string installDirectory, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        var defaultsDir = Path.Combine(installDirectory, "defaults");
        Directory.CreateDirectory(defaultsDir);
        var path = Path.Combine(defaultsDir, "install_state.json");
        File.WriteAllText(path, _serializer.Serialize(state));
        return Task.CompletedTask;
    }

    public Task<InstalledState?> LoadAsync(string installDirectory, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        var path = Path.Combine(installDirectory, "defaults", "install_state.json");
        if (!File.Exists(path))
        {
            return Task.FromResult<InstalledState?>(null);
        }

        return Task.FromResult<InstalledState?>(_serializer.Deserialize<InstalledState>(File.ReadAllText(path)));
    }
}
