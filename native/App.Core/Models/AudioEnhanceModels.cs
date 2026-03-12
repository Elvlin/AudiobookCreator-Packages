namespace App.Core.Models;

public sealed class EnhanceCue
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public double StartSeconds { get; set; }
    public double EndSeconds { get; set; }
    public string CueType { get; set; } = "ambience";
    public string Prompt { get; set; } = string.Empty;
    public double Intensity { get; set; } = 0.6;
    public string Source { get; set; } = "rules";
}

public sealed class EnhanceCueOverride
{
    public string CueId { get; set; } = string.Empty;
    public bool Disabled { get; set; }
    public string Prompt { get; set; } = string.Empty;
    public double? StartSeconds { get; set; }
    public double? EndSeconds { get; set; }
    public double? Intensity { get; set; }
}

public sealed class EnhanceTimelineEvent
{
    public string CueId { get; set; } = string.Empty;
    public string CueType { get; set; } = "ambience";
    public string Prompt { get; set; } = string.Empty;
    public string AudioPath { get; set; } = string.Empty;
    public double StartSeconds { get; set; }
    public double EndSeconds { get; set; }
    public double Intensity { get; set; } = 0.6;
}

public sealed class EnhanceStemPaths
{
    public string VoicePath { get; set; } = string.Empty;
    public string AmbiencePath { get; set; } = string.Empty;
    public string SfxPath { get; set; } = string.Empty;
}

public sealed class EnhanceResult
{
    public bool Success { get; set; }
    public bool UsedEstimatedTiming { get; set; }
    public bool UsedLlmRefine { get; set; }
    public string NarrationOnlyPath { get; set; } = string.Empty;
    public string FinalOutputPath { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public EnhanceStemPaths Stems { get; set; } = new();
    public List<EnhanceCue> Cues { get; set; } = new();
}

public sealed class EnhanceCueExtractionSettings
{
    public int CueMaxPerMinute { get; set; } = 10;
    public bool UseLlmRefine { get; set; } = true;
    public double MinAmbienceHoldSeconds { get; set; } = 6.0;
    public double OneShotCooldownSeconds { get; set; } = 2.5;
}

public sealed class EnhancePreparedSegment
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public int Order { get; set; }
    public double StartSeconds { get; set; }
    public double EndSeconds { get; set; }
    public string Text { get; set; } = string.Empty;
    public string AmbiencePrompt { get; set; } = string.Empty;
    public string OneShotPrompt { get; set; } = string.Empty;
    public double OneShotSeconds { get; set; }
    public double Intensity { get; set; } = 0.6;
    public bool Enabled { get; set; } = true;
}
