using System.IO;
using System.Text.Json;
using App.Core.Models;
using App.Storage;
using AudiobookCreator.Setup.Models;

namespace AudiobookCreator.Setup.Services;

public sealed class AppConfigWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public async Task ApplyAsync(string installDirectory, SetupSelection selection, CancellationToken ct)
    {
        var defaultsDir = Path.Combine(installDirectory, "defaults");
        Directory.CreateDirectory(defaultsDir);
        var configPath = Path.Combine(defaultsDir, "app_config.json");
        var secretStore = new ApiSecretStore(installDirectory);

        AppConfig config;
        if (File.Exists(configPath))
        {
            await using var read = File.OpenRead(configPath);
            config = await JsonSerializer.DeserializeAsync<AppConfig>(read, JsonOptions, ct) ?? new AppConfig();
        }
        else
        {
            config = new AppConfig();
        }

        config.PreferDevice = "auto";

        if (selection.ApiOnly)
        {
            config.BackendMode = "api";
        }
        else
        {
            config.BackendMode = "local";
            config.LocalModelPreset = ResolveLocalModelPreset(selection);
            config.LocalModelBackend = ResolveLocalModelBackend(selection);
        }

        if (selection.EnterApiKeyNow && !string.IsNullOrWhiteSpace(selection.ApiKey))
        {
            config.ApiProvider = selection.ApiProvider;
            var openAiKey = string.Equals(selection.ApiProvider, "openai", StringComparison.OrdinalIgnoreCase)
                ? selection.ApiKey.Trim()
                : string.Empty;
            var alibabaKey = string.Equals(selection.ApiProvider, "alibaba", StringComparison.OrdinalIgnoreCase)
                ? selection.ApiKey.Trim()
                : string.Empty;
            secretStore.SaveKeys(openAiKey, alibabaKey);
        }

        config.ApiKey = string.Empty;
        config.ApiKeyAlibaba = string.Empty;
        config.ApiKeyOpenAi = string.Empty;

        await using var write = File.Create(configPath);
        await JsonSerializer.SerializeAsync(write, config, JsonOptions, ct);
    }

    private static string ResolveLocalModelPreset(SetupSelection selection)
    {
        if (string.Equals(selection.ChatterboxBackend, "onnx", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(selection.ChatterboxBackend, "python", StringComparison.OrdinalIgnoreCase))
        {
            return "chatterbox_onnx";
        }

        if (string.Equals(selection.KittenBackend, "onnx", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(selection.KittenBackend, "python", StringComparison.OrdinalIgnoreCase))
        {
            return "kitten_tts";
        }

        if (selection.IncludeQwen)
        {
            return "qwen3_tts";
        }

        return "chatterbox_onnx";
    }

    private static string ResolveLocalModelBackend(SetupSelection selection)
    {
        if (string.Equals(selection.ChatterboxBackend, "python", StringComparison.OrdinalIgnoreCase))
        {
            return "python";
        }

        if (string.Equals(selection.KittenBackend, "python", StringComparison.OrdinalIgnoreCase))
        {
            return "python";
        }

        if (selection.IncludeQwen)
        {
            return "python";
        }

        return string.Empty;
    }
}
