using System;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using App.Core.Models;

namespace AudiobookCreator.UI;

internal static class ScriptPrepEngine
{
    private static readonly string[] DialogueVerbs =
    {
        "said", "asked", "replied", "whispered", "shouted", "murmured", "yelled", "cried", "called"
    };

    public static PreparedScriptDocument BuildPreparedScript(
        string sourcePath,
        string sourceText,
        string sourceSignature)
    {
        var parts = SplitIntoParts(sourceText)
            .Select((p, idx) => new PreparedScriptPart
            {
                Id = Guid.NewGuid().ToString("N"),
                Order = idx,
                Text = p.Text,
                SpeakerTag = p.SpeakerTag,
                Instruction = BuildInstruction(p.Text, p.SpeakerTag)
            })
            .ToList();

        return new PreparedScriptDocument
        {
            SourcePath = sourcePath,
            SourceSignature = sourceSignature,
            UpdatedAt = DateTimeOffset.UtcNow,
            Version = "v1",
            Parts = parts
        };
    }

    public static List<PreparedScriptPart> ReSplit(string sourceText)
    {
        return SplitIntoParts(sourceText)
            .Select((p, idx) => new PreparedScriptPart
            {
                Id = Guid.NewGuid().ToString("N"),
                Order = idx,
                Text = p.Text,
                SpeakerTag = p.SpeakerTag,
                Instruction = BuildInstruction(p.Text, p.SpeakerTag)
            })
            .ToList();
    }

    public static void RegenerateInstructions(IList<PreparedScriptPart> parts)
    {
        for (var i = 0; i < parts.Count; i++)
        {
            if (parts[i].Locked)
            {
                continue;
            }

            parts[i].Instruction = BuildInstruction(parts[i].Text, parts[i].SpeakerTag);
            parts[i].Order = i;
        }
    }

    private static IReadOnlyList<(string Text, string SpeakerTag)> SplitIntoParts(string text)
    {
        var normalized = (text ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r", "\n", StringComparison.Ordinal)
            .Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return Array.Empty<(string Text, string SpeakerTag)>();
        }

        var blocks = Regex.Split(normalized, @"\n\s*\n")
            .Select(b => Regex.Replace(b, @"[ \t]+", " ").Trim())
            .Where(b => !string.IsNullOrWhiteSpace(b))
            .ToList();

        var result = new List<(string Text, string SpeakerTag)>();
        var speakerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var speakerCounter = 0;

        foreach (var block in blocks)
        {
            // Match straight quotes + smart quotes and keep quote marks in stored dialogue text.
            var dialogueHits = Regex.Matches(block, "(?:\"[^\"]+\")|(?:\u201C[^\u201D]+\u201D)|(?:â€œ[^â€]+â€)");
            if (dialogueHits.Count == 0)
            {
                result.Add((block, "Narrator"));
                continue;
            }

            var cursor = 0;
            foreach (Match hit in dialogueHits)
            {
                if (hit.Index > cursor)
                {
                    var narration = block[cursor..hit.Index].Trim();
                    if (!string.IsNullOrWhiteSpace(narration))
                    {
                        result.Add((narration, "Narrator"));
                    }
                }

                var quote = hit.Value.Trim();
                if (!string.IsNullOrWhiteSpace(quote))
                {
                    var speaker = ResolveSpeakerTag(block, hit.Index, speakerMap, ref speakerCounter);
                    result.Add((quote, speaker));
                }

                cursor = hit.Index + hit.Length;
            }

            if (cursor < block.Length)
            {
                var tail = block[cursor..].Trim();
                if (!string.IsNullOrWhiteSpace(tail))
                {
                    result.Add((tail, "Narrator"));
                }
            }
        }

        return result;
    }

    private static string ResolveSpeakerTag(
        string block,
        int quoteIndex,
        Dictionary<string, string> speakerMap,
        ref int counter)
    {
        var windowStart = Math.Max(0, quoteIndex - 80);
        var prefix = block[windowStart..Math.Min(block.Length, quoteIndex + 80)];
        var verbs = string.Join("|", DialogueVerbs);

        // Pattern A: "Alex said ..."
        var match = Regex.Match(prefix, $@"\b([A-Z][a-z]+)\s+({verbs})\b", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            // Pattern B: "... said Alex"
            match = Regex.Match(prefix, $@"\b({verbs})\s+([A-Z][a-z]+)\b", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                var nameAfterVerb = match.Groups[2].Value.Trim();
                if (!speakerMap.TryGetValue(nameAfterVerb, out var mapped))
                {
                    counter++;
                    mapped = $"Speaker_{counter:00}";
                    speakerMap[nameAfterVerb] = mapped;
                }
                return mapped;
            }

            // Unknown speaker: create a new speaker id (do not reuse the same one).
            counter++;
            return $"Speaker_{counter:00}";
        }

        var name = match.Groups[1].Value.Trim();
        if (!speakerMap.TryGetValue(name, out var tag))
        {
            counter++;
            tag = $"Speaker_{counter:00}";
            speakerMap[name] = tag;
        }

        return tag;
    }

    private static string BuildInstruction(string text, string speakerTag)
    {
        var t = (text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(t))
        {
            return "neutral narration, clear diction, steady pacing";
        }

        var lower = t.ToLowerInvariant();
        var emotional = lower.Contains("!") || lower.Contains("please") || lower.Contains("no!") ||
                        lower.Contains("help") || lower.Contains("cry") || lower.Contains("angry");
        var question = t.EndsWith("?", StringComparison.Ordinal);
        var whisper = lower.Contains("whisper") || lower.Contains("quietly");

        if (!string.Equals(speakerTag, "Narrator", StringComparison.OrdinalIgnoreCase))
        {
            if (whisper)
            {
                return "dialogue voice, soft whispering tone, low volume, intimate delivery";
            }
            if (question)
            {
                return "dialogue voice, questioning tone, clear phrasing, slight upward inflection";
            }
            if (emotional)
            {
                return "dialogue voice, emotional intensity, expressive delivery, controlled pacing";
            }

            return "dialogue voice, natural conversational tone, clean pronunciation";
        }

        if (whisper)
        {
            return "narration, soft low-volume tone, calm pacing";
        }
        if (emotional)
        {
            return "narration, expressive emotional tone, strong emphasis on key words";
        }
        if (question)
        {
            return "narration, inquisitive tone, smooth rhythm";
        }

        return "narration, neutral warm tone, clear diction, steady pacing";
    }
}
