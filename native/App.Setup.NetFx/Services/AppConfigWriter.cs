using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using AudiobookCreator.SetupNetFx.Models;

namespace AudiobookCreator.SetupNetFx.Services;

public sealed class AppConfigWriter
{
    private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer
    {
        MaxJsonLength = int.MaxValue
    };

    public Task ApplyAsync(string installDirectory, SetupSelection selection, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        var defaultsDir = Path.Combine(installDirectory, "defaults");
        Directory.CreateDirectory(defaultsDir);
        var configPath = Path.Combine(defaultsDir, "app_config.json");

        var config = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        if (File.Exists(configPath))
        {
            var existing = _serializer.DeserializeObject(File.ReadAllText(configPath)) as Dictionary<string, object>;
            if (existing != null)
            {
                config = existing;
            }
        }

        config["preferDevice"] = "auto";
        if (selection.ApiOnly)
        {
            config["backendMode"] = "api";
        }
        else
        {
            config["backendMode"] = "local";
            config["localModelPreset"] = ResolveLocalModelPreset(selection);
            config["localModelBackend"] = ResolveLocalModelBackend(selection);
        }

        if (selection.EnterApiKeyNow && !string.IsNullOrWhiteSpace(selection.ApiKey))
        {
            config["apiProvider"] = selection.ApiProvider;
            SaveSecrets(installDirectory, selection);
        }

        config["apiKey"] = string.Empty;
        config["apiKeyOpenAi"] = string.Empty;
        config["apiKeyAlibaba"] = string.Empty;

        File.WriteAllText(configPath, _serializer.Serialize(config));
        return Task.CompletedTask;
    }

    private static void SaveSecrets(string installDirectory, SetupSelection selection)
    {
        var payload = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["openAiKey"] = string.Equals(selection.ApiProvider, "openai", StringComparison.OrdinalIgnoreCase) ? selection.ApiKey.Trim() : string.Empty,
            ["alibabaKey"] = string.Equals(selection.ApiProvider, "alibaba", StringComparison.OrdinalIgnoreCase) ? selection.ApiKey.Trim() : string.Empty
        };

        var json = new JavaScriptSerializer().Serialize(payload);
        var plainBytes = Encoding.UTF8.GetBytes(json);
        var entropy = Encoding.UTF8.GetBytes(installDirectory);
        var protectedBytes = ProtectedData.Protect(plainBytes, entropy, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(Path.Combine(installDirectory, "defaults", "api_secrets.dat"), protectedBytes);
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
        if (string.Equals(selection.ChatterboxBackend, "python", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(selection.KittenBackend, "python", StringComparison.OrdinalIgnoreCase) ||
            selection.IncludeQwen)
        {
            return "python";
        }

        return string.Empty;
    }
}
