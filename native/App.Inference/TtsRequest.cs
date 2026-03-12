namespace App.Inference;

public sealed class TtsRequest
{
    public string Text { get; init; } = string.Empty;
    public string VoicePath { get; init; } = string.Empty;
    /// <summary>Optional transcript of the voice sample (ref_text). Used by Qwen3 for better voice clone when set.</summary>
    public string RefText { get; init; } = string.Empty;
    public string OutputPath { get; init; } = string.Empty;
    public float Speed { get; init; } = 1.0f;
    public string StylePresetKey { get; init; } = string.Empty;
    public float? ChatterboxExaggeration { get; init; }
    public bool? QwenDoSample { get; init; }
    public float? QwenTemperature { get; init; }
    public int? QwenTopK { get; init; }
    public float? QwenTopP { get; init; }
    public float? QwenRepetitionPenalty { get; init; }
    public bool QwenAutoRetryBadChunks { get; init; }
    public int QwenBadChunkRetryCount { get; init; }
    public bool? QwenXVectorOnlyMode { get; init; }
    public int ChunkPauseMs { get; init; }
    public int ParagraphPauseMs { get; init; }
    public int ClausePauseMinMs { get; init; }
    public int ClausePauseMaxMs { get; init; }
    public int SentencePauseMinMs { get; init; }
    public int SentencePauseMaxMs { get; init; }
    public int EllipsisPauseMinMs { get; init; }
    public int EllipsisPauseMaxMs { get; init; }
    public int ParagraphPauseMinMs { get; init; }
    public int ParagraphPauseMaxMs { get; init; }
}
