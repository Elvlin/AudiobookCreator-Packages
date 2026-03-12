namespace App.Core.Models;

public sealed class AppConfig
{
    public string DefaultOutputDir { get; set; } = "output";
    public string ModelCacheDir { get; set; } = "models";
    public string PreferDevice { get; set; } = "auto";
    public bool OfflineMode { get; set; }
    public string BackendMode { get; set; } = "local";
    public string LocalModelPreset { get; set; } = "chatterbox_onnx";
    public string LocalModelBackend { get; set; } = string.Empty;
    public string ModelRepoId { get; set; } = "onnx-community/chatterbox-ONNX";
    public string AdditionalModelRepoId { get; set; } = string.Empty;
    public bool AutoDownloadModel { get; set; } = true;
    public bool AutoRemoveCompletedInputFiles { get; set; }
    public bool GenerateSrtSubtitles { get; set; }
    public bool GenerateAssSubtitles { get; set; }
    public string ApiProvider { get; set; } = "openai";
    public string ApiPreset { get; set; } = "openai_default";
    public string ApiKey { get; set; } = string.Empty;
    public string ApiKeyAlibaba { get; set; } = string.Empty;
    public string ApiKeyOpenAi { get; set; } = string.Empty;
    public string ApiModelId { get; set; } = "gpt-4o-mini-tts";
    public string ApiVoice { get; set; } = string.Empty;
    public string ApiBaseUrl { get; set; } = "https://dashscope-intl.aliyuncs.com/api/v1";
    public string ApiLanguageType { get; set; } = "Auto";
    public string ApiVoiceDesignTargetModel { get; set; } = "qwen3-tts-vd-2026-01-26";
    public List<string> CachedOpenAiLlmModels { get; set; } = new();
    public List<string> CachedAlibabaLlmModels { get; set; } = new();
    public List<string> CachedOpenAiTtsModels { get; set; } = new();
    public List<string> CachedAlibabaTtsModels { get; set; } = new();
    public bool AudioEnhanceEnabledByDefault { get; set; }
    public bool AudioEnhanceAutoRun { get; set; } = true;
    public string AudioEnhanceMode { get; set; } = "sfx_ambience";
    public string AudioEnhanceProvider { get; set; } = "local_audiox";
    public string AudioEnhanceVariant { get; set; } = "audiox_base";
    public string AudioEnhanceRuntimePath { get; set; } = string.Empty;
    public string AudioEnhanceModelDir { get; set; } = "models/audiox";
    public string AudioEnhanceModelRepoId { get; set; } = "HKUSTAudio/AudioX";
    public bool AudioEnhanceUseLlmCueDetection { get; set; }
    public bool AudioEnhanceUseLlmCueRefine { get; set; } = true;
    public double AudioEnhanceAmbienceDb { get; set; } = -24.0;
    public double AudioEnhanceOneShotDb { get; set; } = -15.0;
    public double AudioEnhanceDuckingDb { get; set; } = -8.0;
    public int AudioEnhanceCueMaxPerMinute { get; set; } = 10;
    public bool AudioEnhanceExportNarrationOnly { get; set; } = true;
    public bool AudioEnhanceExportStems { get; set; } = true;
    public string LlmPrepProvider { get; set; } = "local";
    public string LlmPrepSplitModel { get; set; } = "qwen2.5-7b-instruct";
    public string LlmPrepInstructionModel { get; set; } = "qwen2.5-14b-instruct";
    public bool LlmPrepUseSeparateModels { get; set; } = true;
    public string LlmLocalRuntimePath { get; set; } = string.Empty;
    public string LlmLocalModelDir { get; set; } = "models/llm";
    public string LlmLocalSplitModelFile { get; set; } = "qwen2.5-7b-instruct.gguf";
    public string LlmLocalInstructionModelFile { get; set; } = "qwen2.5-14b-instruct.gguf";
    public double LlmPrepTemperatureSplit { get; set; } = 0.2;
    public double LlmPrepTemperatureInstruction { get; set; } = 0.6;
    public int LlmPrepMaxTokensSplit { get; set; } = 1024;
    public int LlmPrepMaxTokensInstruction { get; set; } = 512;
    public Dictionary<string, SynthesisSettings> ModelProfiles { get; set; } = new();
    public string LastOpenDir { get; set; } = string.Empty;
    public List<string> RecentProjects { get; set; } = new();
}
