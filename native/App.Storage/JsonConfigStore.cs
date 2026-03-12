using System.Text.Json;
using System.Runtime.Versioning;
using App.Core.Models;
using App.Core.Services;

namespace App.Storage;

[SupportedOSPlatform("windows")]
public sealed class JsonConfigStore : IConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public async Task<AppConfig> LoadAsync(CancellationToken ct = default)
    {
        var path = RuntimePaths.ConfigPath;
        var secretStore = new ApiSecretStore(RuntimePaths.AppRoot);
        if (!File.Exists(path))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var cfg = new AppConfig();
            await SaveAsync(cfg, ct);
            return cfg;
        }

        await using var fs = File.OpenRead(path);
        var config = await JsonSerializer.DeserializeAsync<AppConfig>(fs, JsonOptions, ct);
        config ??= new AppConfig();

        var migrated = secretStore.MigrateLegacySecrets(config);
        config.ApiKeyOpenAi = secretStore.GetOpenAiKey();
        config.ApiKeyAlibaba = secretStore.GetAlibabaKey();
        config.ApiKey = string.Empty;

        if (migrated)
        {
            await SaveAsync(config, ct);
        }

        return config;
    }

    public async Task SaveAsync(AppConfig config, CancellationToken ct = default)
    {
        var path = RuntimePaths.ConfigPath;
        var secretStore = new ApiSecretStore(RuntimePaths.AppRoot);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        secretStore.SaveKeys(config.ApiKeyOpenAi, config.ApiKeyAlibaba);
        var persisted = CloneWithoutSecrets(config);
        await using var fs = File.Create(path);
        await JsonSerializer.SerializeAsync(fs, persisted, JsonOptions, ct);
    }

    private static AppConfig CloneWithoutSecrets(AppConfig config)
    {
        return new AppConfig
        {
            DefaultOutputDir = config.DefaultOutputDir,
            ModelCacheDir = config.ModelCacheDir,
            PreferDevice = config.PreferDevice,
            OfflineMode = config.OfflineMode,
            BackendMode = config.BackendMode,
            LocalModelPreset = config.LocalModelPreset,
            LocalModelBackend = config.LocalModelBackend,
            ModelRepoId = config.ModelRepoId,
            AdditionalModelRepoId = config.AdditionalModelRepoId,
            AutoDownloadModel = config.AutoDownloadModel,
            AutoRemoveCompletedInputFiles = config.AutoRemoveCompletedInputFiles,
            GenerateSrtSubtitles = config.GenerateSrtSubtitles,
            GenerateAssSubtitles = config.GenerateAssSubtitles,
            ApiProvider = config.ApiProvider,
            ApiPreset = config.ApiPreset,
            ApiKey = string.Empty,
            ApiKeyAlibaba = string.Empty,
            ApiKeyOpenAi = string.Empty,
            ApiModelId = config.ApiModelId,
            ApiVoice = config.ApiVoice,
            ApiBaseUrl = config.ApiBaseUrl,
            ApiLanguageType = config.ApiLanguageType,
            ApiVoiceDesignTargetModel = config.ApiVoiceDesignTargetModel,
            CachedOpenAiLlmModels = new(config.CachedOpenAiLlmModels),
            CachedAlibabaLlmModels = new(config.CachedAlibabaLlmModels),
            CachedOpenAiTtsModels = new(config.CachedOpenAiTtsModels),
            CachedAlibabaTtsModels = new(config.CachedAlibabaTtsModels),
            AudioEnhanceEnabledByDefault = config.AudioEnhanceEnabledByDefault,
            AudioEnhanceAutoRun = config.AudioEnhanceAutoRun,
            AudioEnhanceMode = config.AudioEnhanceMode,
            AudioEnhanceProvider = config.AudioEnhanceProvider,
            AudioEnhanceVariant = config.AudioEnhanceVariant,
            AudioEnhanceRuntimePath = config.AudioEnhanceRuntimePath,
            AudioEnhanceModelDir = config.AudioEnhanceModelDir,
            AudioEnhanceModelRepoId = config.AudioEnhanceModelRepoId,
            AudioEnhanceUseLlmCueDetection = config.AudioEnhanceUseLlmCueDetection,
            AudioEnhanceUseLlmCueRefine = config.AudioEnhanceUseLlmCueRefine,
            AudioEnhanceAmbienceDb = config.AudioEnhanceAmbienceDb,
            AudioEnhanceOneShotDb = config.AudioEnhanceOneShotDb,
            AudioEnhanceDuckingDb = config.AudioEnhanceDuckingDb,
            AudioEnhanceCueMaxPerMinute = config.AudioEnhanceCueMaxPerMinute,
            AudioEnhanceExportNarrationOnly = config.AudioEnhanceExportNarrationOnly,
            AudioEnhanceExportStems = config.AudioEnhanceExportStems,
            LlmPrepProvider = config.LlmPrepProvider,
            LlmPrepSplitModel = config.LlmPrepSplitModel,
            LlmPrepInstructionModel = config.LlmPrepInstructionModel,
            LlmPrepUseSeparateModels = config.LlmPrepUseSeparateModels,
            LlmLocalRuntimePath = config.LlmLocalRuntimePath,
            LlmLocalModelDir = config.LlmLocalModelDir,
            LlmLocalSplitModelFile = config.LlmLocalSplitModelFile,
            LlmLocalInstructionModelFile = config.LlmLocalInstructionModelFile,
            LlmPrepTemperatureSplit = config.LlmPrepTemperatureSplit,
            LlmPrepTemperatureInstruction = config.LlmPrepTemperatureInstruction,
            LlmPrepMaxTokensSplit = config.LlmPrepMaxTokensSplit,
            LlmPrepMaxTokensInstruction = config.LlmPrepMaxTokensInstruction,
            ModelProfiles = new(config.ModelProfiles),
            LastOpenDir = config.LastOpenDir,
            RecentProjects = new(config.RecentProjects)
        };
    }
}
