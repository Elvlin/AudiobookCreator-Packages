using System.Security.Cryptography;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;

namespace App.Storage;

[SupportedOSPlatform("windows")]
public sealed class ApiSecretStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    private readonly string _appRoot;

    public ApiSecretStore(string appRoot)
    {
        _appRoot = appRoot;
    }

    private string SecretsPath => Path.Combine(_appRoot, "defaults", "api_secrets.dat");

    public bool HasAnySecrets()
    {
        return File.Exists(SecretsPath);
    }

    public string GetOpenAiKey()
    {
        var payload = LoadPayload();
        return payload.OpenAiKey;
    }

    public string GetAlibabaKey()
    {
        var payload = LoadPayload();
        return payload.AlibabaKey;
    }

    public void SaveKeys(string? openAiKey, string? alibabaKey)
    {
        var payload = new SecretPayload
        {
            OpenAiKey = (openAiKey ?? string.Empty).Trim(),
            AlibabaKey = (alibabaKey ?? string.Empty).Trim()
        };

        if (string.IsNullOrWhiteSpace(payload.OpenAiKey) && string.IsNullOrWhiteSpace(payload.AlibabaKey))
        {
            DeleteSecrets();
            return;
        }

        var defaultsDir = Path.Combine(_appRoot, "defaults");
        Directory.CreateDirectory(defaultsDir);
        var json = JsonSerializer.Serialize(payload, JsonOptions);
        var plainBytes = Encoding.UTF8.GetBytes(json);
        var protectedBytes = ProtectedData.Protect(plainBytes, BuildEntropy(), DataProtectionScope.CurrentUser);
        File.WriteAllBytes(SecretsPath, protectedBytes);
    }

    public void DeleteSecrets()
    {
        if (File.Exists(SecretsPath))
        {
            File.Delete(SecretsPath);
        }
    }

    public bool MigrateLegacySecrets(App.Core.Models.AppConfig config)
    {
        var legacyOpenAi = (config.ApiKeyOpenAi ?? string.Empty).Trim();
        var legacyAlibaba = (config.ApiKeyAlibaba ?? string.Empty).Trim();
        var legacyGeneric = (config.ApiKey ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(legacyOpenAi) &&
            string.IsNullOrWhiteSpace(legacyAlibaba) &&
            string.IsNullOrWhiteSpace(legacyGeneric))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(legacyOpenAi) && !string.IsNullOrWhiteSpace(legacyGeneric))
        {
            var provider = (config.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
            if (provider == "alibaba")
            {
                legacyAlibaba = legacyGeneric;
            }
            else
            {
                legacyOpenAi = legacyGeneric;
            }
        }

        SaveKeys(legacyOpenAi, legacyAlibaba);
        config.ApiKey = string.Empty;
        config.ApiKeyOpenAi = string.Empty;
        config.ApiKeyAlibaba = string.Empty;
        return true;
    }

    private SecretPayload LoadPayload()
    {
        if (!File.Exists(SecretsPath))
        {
            return new SecretPayload();
        }

        try
        {
            var protectedBytes = File.ReadAllBytes(SecretsPath);
            var plainBytes = ProtectedData.Unprotect(protectedBytes, BuildEntropy(), DataProtectionScope.CurrentUser);
            var payload = JsonSerializer.Deserialize<SecretPayload>(plainBytes, JsonOptions);
            return payload ?? new SecretPayload();
        }
        catch
        {
            return new SecretPayload();
        }
    }

    private byte[] BuildEntropy()
    {
        return Encoding.UTF8.GetBytes(_appRoot);
    }

    private sealed class SecretPayload
    {
        public string OpenAiKey { get; set; } = string.Empty;
        public string AlibabaKey { get; set; } = string.Empty;
    }
}
