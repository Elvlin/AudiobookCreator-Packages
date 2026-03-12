namespace App.Inference;

public sealed class InferenceOptions
{
    public string Provider { get; init; } = "openai";
    public string ApiKey { get; init; } = string.Empty;
    public string ModelId { get; init; } = "gpt-4o-mini-tts";
    public string Voice { get; init; } = string.Empty;
    public string BaseUrl { get; init; } = "https://dashscope-intl.aliyuncs.com/api/v1";
    public string LanguageType { get; init; } = "Auto";
    public string VoiceDesignTargetModel { get; init; } = "qwen3-tts-vd-2026-01-26";
}
