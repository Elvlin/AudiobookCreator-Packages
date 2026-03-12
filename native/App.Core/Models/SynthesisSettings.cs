namespace App.Core.Models;

public sealed class SynthesisSettings
{
    public string OutputFormat { get; set; } = "wav";
    public string ChunkMode { get; set; } = "auto";
    public int ManualMaxChars { get; set; } = 600;
    public int MinChars { get; set; } = 200;
    public int MaxChars { get; set; } = 420;
    public double NarrationTargetSec { get; set; } = 4.4;
    public double DialogueTargetSec { get; set; } = 3.4;
    public double DialogueOverflow { get; set; } = 1.8;
    public int ChunkPauseMs { get; set; } = 30;
    public int ParagraphPauseMs { get; set; } = 180;
    public int ClausePauseMinMs { get; set; }
    public int ClausePauseMaxMs { get; set; }
    public int SentencePauseMinMs { get; set; }
    public int SentencePauseMaxMs { get; set; }
    public int EllipsisPauseMinMs { get; set; }
    public int EllipsisPauseMaxMs { get; set; }
    public int ParagraphPauseMinMs { get; set; }
    public int ParagraphPauseMaxMs { get; set; }
    public double Atempo { get; set; } = 1.06;
    public bool ReadSmallNumbersAsWords { get; set; } = true;
    public bool ProtectBracketDirectives { get; set; } = true;
    public string LocalInstructionHint { get; set; } = string.Empty;
    public string StylePresetKey { get; set; } = "standard";
    public bool QwenStableAudiobookPreset { get; set; } = true;
    public bool QwenUseRefText { get; set; } = false;
    public bool QwenDoSample { get; set; } = false;
    public double QwenTemperature { get; set; } = 1.0;
    public int QwenTopK { get; set; } = 50;
    public double QwenTopP { get; set; } = 1.0;
    public double QwenRepetitionPenalty { get; set; } = 1.0;
    public bool QwenAutoRetryBadChunks { get; set; } = true;
    public int QwenBadChunkRetryCount { get; set; } = 3;
}
