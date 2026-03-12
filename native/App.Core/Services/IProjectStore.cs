using App.Core.Models;

namespace App.Core.Services;

public interface IProjectStore
{
    Task<ProjectDocument> LoadAsync(string path, CancellationToken ct = default);
    Task SaveAsync(ProjectDocument project, string path, CancellationToken ct = default);
}
