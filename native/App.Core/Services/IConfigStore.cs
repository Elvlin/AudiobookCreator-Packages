using App.Core.Models;

namespace App.Core.Services;

public interface IConfigStore
{
    Task<AppConfig> LoadAsync(CancellationToken ct = default);
    Task SaveAsync(AppConfig config, CancellationToken ct = default);
}
