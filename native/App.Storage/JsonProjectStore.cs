using System.Text.Json;
using App.Core.Models;
using App.Core.Services;

namespace App.Storage;

public sealed class JsonProjectStore : IProjectStore
{
    public const string Extension = ".abproj";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public async Task<ProjectDocument> LoadAsync(string path, CancellationToken ct = default)
    {
        await using var fs = File.OpenRead(path);
        var project = await JsonSerializer.DeserializeAsync<ProjectDocument>(fs, JsonOptions, ct);
        return project ?? new ProjectDocument();
    }

    public async Task SaveAsync(ProjectDocument project, string path, CancellationToken ct = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await using var fs = File.Create(path);
        await JsonSerializer.SerializeAsync(fs, project, JsonOptions, ct);
    }
}
