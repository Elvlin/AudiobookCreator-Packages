namespace App.Core.Models;

public sealed class ProjectDocument
{
    public string Name { get; set; } = "New Audiobook Project";
    public string SourceTextPath { get; set; } = string.Empty;
    public List<string> SourceTextFiles { get; set; } = new();
    public string VoicePath { get; set; } = string.Empty;
    public string OutputDir { get; set; } = "output";
    public string OutputBasename { get; set; } = "audiobook";
    public string InlineText { get; set; } = string.Empty;
    public SynthesisSettings Settings { get; set; } = new();
    public List<ProjectQueueRowState> QueueRows { get; set; } = new();
    public List<ProjectHistoryRowState> HistoryRows { get; set; } = new();
    public Dictionary<string, PreparedScriptDocument> PreparedScriptsBySourcePath { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);
    public string VoiceDesignDraftName { get; set; } = string.Empty;
    public string VoiceDesignDraftText { get; set; } = string.Empty;
    public string VoiceDesignDraftPrompt { get; set; } = string.Empty;
    public string VoiceDesignDraftLanguage { get; set; } = "auto";
    public bool AudioEnhanceEnabled { get; set; }
    public bool AudioEnhanceUseLlmCueRefine { get; set; } = true;
    public string AudioEnhanceProfile { get; set; } = "balanced";
    public bool AudioEnhanceUseAllInputFiles { get; set; }
    public string AudioEnhanceOutputDir { get; set; } = string.Empty;
    public string AudioEnhanceSceneNotes { get; set; } = string.Empty;
    public List<ProjectAudioEnhanceInputFileState> AudioEnhanceInputFiles { get; set; } = new();
    public List<EnhanceCueOverride> CueOverrides { get; set; } = new();
    public Dictionary<string, List<EnhancePreparedSegment>> AudioEnhancePreparedSegmentsBySourcePath { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);
}

public sealed class ProjectAudioEnhanceInputFileState
{
    public string FullPath { get; set; } = string.Empty;
    public bool IsSelected { get; set; }
}

public sealed class PreparedScriptDocument
{
    public string SourcePath { get; set; } = string.Empty;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.UtcNow;
    public string Version { get; set; } = "v1";
    public string SourceSignature { get; set; } = string.Empty;
    public List<PreparedScriptPart> Parts { get; set; } = new();
}

public sealed class PreparedScriptPart
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public int Order { get; set; }
    public string Text { get; set; } = string.Empty;
    public string Instruction { get; set; } = string.Empty;
    public string? VoicePath { get; set; }
    public string SpeakerTag { get; set; } = "Narrator";
    public bool Locked { get; set; }
    public bool DisableAutoSfxDetection { get; set; }
    public string SoundEffectPrompt { get; set; } = string.Empty;
}

public sealed class ProjectQueueRowState
{
    public string FileName { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string Status { get; set; } = "Queued";
    public string ProgressLabel { get; set; } = "0%";
    public double ProgressValue { get; set; }
    public string ChunkInfo { get; set; } = "Queued";
    public string Eta { get; set; } = "Queued";
    public string OutputPath { get; set; } = string.Empty;
}

public sealed class ProjectHistoryRowState
{
    public string FileName { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public string ModelLabel { get; set; } = string.Empty;
    public string DeviceLabel { get; set; } = string.Empty;
    public string VoiceLabel { get; set; } = string.Empty;
    public string CompletedAt { get; set; } = string.Empty;
    public string DurationLabel { get; set; } = string.Empty;
}
