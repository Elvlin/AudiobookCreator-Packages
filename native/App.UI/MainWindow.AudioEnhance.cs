using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using App.Core.Models;
using App.Storage;

namespace AudiobookCreator.UI;

public partial class MainWindow
{
    private static readonly (string[] Keywords, string Prompt)[] AmbienceCueRules =
    {
        (new[] { "rain", "storm", "drizzle", "thunderstorm" }, "steady rain ambience, distant weather, natural outdoor bed"),
        (new[] { "wind", "gust", "blizzard" }, "soft windy ambience, airy movement, cinematic atmosphere"),
        (new[] { "crowd", "market", "street", "city", "traffic" }, "distant urban ambience, crowd murmur, soft city bed"),
        (new[] { "forest", "woods", "jungle", "birds" }, "forest ambience, soft leaves, distant birds, natural space"),
        (new[] { "ocean", "sea", "shore", "waves" }, "ocean ambience, gentle waves, low coastal air"),
        (new[] { "cave", "tunnel", "underground" }, "hollow cave ambience, subtle low reverb, dark air"),
        (new[] { "silence", "quiet", "tense", "tension" }, "very low tense room tone, subtle cinematic suspense bed"),
        (new[] { "room", "hall", "hallway", "corridor", "chamber", "office", "library" }, "soft indoor room tone, light natural reverb, subtle interior air"),
        (new[] { "tavern", "inn", "bar", "restaurant", "cafe" }, "soft indoor public ambience, low voices, room tone"),
        (new[] { "battlefield", "battle", "war", "camp" }, "distant battle ambience, low movement, cinematic battlefield bed"),
        (new[] { "fire", "flame", "fireplace", "hearth", "campfire" }, "gentle fire crackle ambience, warm low bed"),
        (new[] { "river", "stream", "waterfall" }, "steady running water ambience, natural outdoor bed"),
        (new[] { "harbor", "port", "dock", "ship" }, "distant harbor ambience, water movement, creaks, low activity"),
        (new[] { "school", "classroom" }, "light indoor school ambience, room tone, faint distant movement"),
        (new[] { "train", "station", "platform" }, "subtle station ambience, low movement, distant mechanical bed"),
        (new[] { "engine", "generator", "machine room", "machinery", "factory" }, "low mechanical ambience, steady machine hum, industrial bed"),
        (new[] { "church", "cathedral", "temple", "sanctuary" }, "large reverberant indoor ambience, quiet ceremonial hall tone"),
        (new[] { "hospital", "clinic", "ward" }, "clean indoor medical ambience, faint room tone, distant equipment"),
        (new[] { "prison", "cell", "dungeon" }, "cold enclosed ambience, low room tone, restrained echo"),
        (new[] { "snow", "snowstorm", "icy wind" }, "cold winter ambience, soft wind, sparse outdoor atmosphere")
    };

    private static readonly (string[] Keywords, string Prompt)[] OneShotCueRules =
    {
        (new[] { "door slam", "slammed the door", "door crashed" }, "heavy wooden door slam impact"),
        (new[] { "door opened", "door open", "opened the door", "door creaked", "door creak", "door closed", "closed the door", "shut the door" }, "wooden door movement one-shot"),
        (new[] { "knock", "knocked", "knocking" }, "firm wooden door knock one-shot"),
        (new[] { "footstep", "footsteps", "steps", "boots on stone", "boots on wood", "boots on the floor", "ran across", "sprinted across" }, "short footsteps on hard floor"),
        (new[] { "glass shatter", "glass broke", "glass break" }, "sharp glass shatter one-shot"),
        (new[] { "gunshot", "shot fired", "fired a shot" }, "single distant gunshot impact"),
        (new[] { "thunder", "lightning strike" }, "single thunder crack impact"),
        (new[] { "sword clash", "metal clash", "blade hit", "steel rang", "ring of steel", "metal scraped", "steel scraped", "drew his sword", "drew her sword", "blade scraped", "sword scraped" }, "metallic sword clash impact"),
        (new[] { "slash", "slashed", "sliced", "cut through", "carved through", "stabbed", "pierced", "lunged", "blade tore" }, "sharp blade slash one-shot"),
        (new[] { "impact", "hit", "crash", "exploded", "smashed into", "slammed into", "collided with", "rammed into", "struck the wall", "hit the wall", "hit the ground", "landed hard" }, "cinematic impact one-shot"),
        (new[] { "scream", "screamed", "shout", "shouted", "yell", "yelled" }, "single human shout one-shot"),
        (new[] { "punch", "punched", "slap", "slapped", "kick", "kicked", "elbowed", "kneed", "socked him", "socked her", "decked him", "decked her" }, "body hit impact one-shot"),
        (new[] { "explosion", "blast", "detonation" }, "explosion impact one-shot"),
        (new[] { "drop", "dropped", "fell to the floor", "fell on the floor", "body fell", "collapsed to the floor", "crashed to the ground", "dropped to one knee" }, "object drop impact one-shot"),
        (new[] { "boom", "bam", "bang", "pow", "kaboom" }, "heavy cinematic impact one-shot"),
        (new[] { "thud", "thump", "wham", "smack" }, "blunt impact hit one-shot"),
        (new[] { "clang", "clang!", "clank", "clink" }, "metal hit one-shot"),
        (new[] { "crack", "crack!", "snap", "snap!" }, "sharp crack one-shot"),
        (new[] { "whoosh", "woosh", "swoosh", "swish" }, "fast whoosh pass-by one-shot"),
        (new[] { "creak", "creaked" }, "wooden creak one-shot"),
        (new[] { "click", "clicked", "tick", "ticked", "tap", "tapped" }, "small mechanical click one-shot"),
        (new[] { "beep", "buzz", "ring", "rang" }, "short device alert one-shot"),
        (new[] { "roar", "growl", "hiss", "snarl", "howled" }, "animal or creature vocal one-shot"),
        (new[] { "floorboards creaked", "wood groaned", "boards groaned", "chair scraped", "scraped against the floor", "dragged across the floor" }, "wood scrape or creak one-shot"),
        (new[] { "armor clattered", "chain rattled", "coins clinked", "keys jingled" }, "small metallic movement one-shot"),
        (new[] { "breath caught", "gasped", "sharp inhale" }, "short startled breath one-shot")
    };

    private static readonly (Regex Pattern, string Prompt)[] OneShotCueRegexRules =
    {
        (new Regex(@"\b(boom|bam|bang|pow|kaboom|thud|thump|wham|smack)\b[!?.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "heavy cinematic impact one-shot"),
        (new Regex(@"\b(clang|clank|clink)\b[!?.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "metal hit one-shot"),
        (new Regex(@"\b(crack|snap)\b[!?.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "sharp crack one-shot"),
        (new Regex(@"\b(whoosh|woosh|swoosh|swish)\b[!?.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "fast whoosh pass-by one-shot"),
        (new Regex(@"\b(creak|click|clack|tap|beep|buzz|ring)\b[!?.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "short mechanical or object one-shot"),
        (new Regex(@"\b(slash(?:ed|es|ing)?|slice(?:d|s|ing)?|stab(?:bed|s|bing)?|pierce(?:d|s|ing)?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "sharp blade slash one-shot"),
        (new Regex(@"\b(punch(?:ed|es|ing)?|kick(?:ed|s|ing)?|slap(?:ped|s|ping)?|smash(?:ed|es|ing)? into|slam(?:med|s|ming)? into)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "body hit impact one-shot"),
        (new Regex(@"\b(scrape(?:d|s|ing)?|skid(?:ded|s|ding)?|scuff(?:ed|s|ing)?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant), "hard scrape one-shot")
    };

    private sealed record TimedTextUnit(string Text, double StartSeconds, double EndSeconds);

    private sealed class AudioEnhanceProgress
    {
        public string Stage { get; init; } = string.Empty;
        public int Completed { get; init; }
        public int Total { get; init; }
    }

    private void ApplyAudioEnhanceUiFromProject()
    {
        if (EnhanceAudioCheckBox is null)
        {
            return;
        }

        EnhanceAudioCheckBox.IsChecked = _project.AudioEnhanceEnabled;
    }

    private void EnhanceAudioCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        _project.AudioEnhanceEnabled = EnhanceAudioCheckBox?.IsChecked == true;
        _ = AutoSaveCurrentProjectAsync();
    }

    private void EnhanceSettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        SettingsButton_OnClick(sender, e);
    }

    private bool IsEnhanceEnabledForRun()
    {
        if (!AppFeatureFlags.AudioEnhancementEnabled)
        {
            return false;
        }

        if (!_config.AudioEnhanceAutoRun)
        {
            return false;
        }

        return EnhanceAudioCheckBox?.IsChecked == true;
    }

    private async Task<EnhanceResult> RunAudioEnhanceForChapterAsync(
        string sourceText,
        string outputPath,
        IReadOnlyList<SubtitleChunkTiming>? chunkTimings,
        PreparedScriptDocument? preparedScript,
        CancellationToken ct,
        IProgress<AudioEnhanceProgress>? progress = null)
    {
        var result = new EnhanceResult
        {
            Success = false,
            FinalOutputPath = outputPath
        };

        var outputDir = Path.GetDirectoryName(outputPath) ?? RuntimePaths.AppRoot;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath).Trim().TrimStart('.').ToLowerInvariant();

        var workRoot = Path.Combine(outputDir, ".audio_enhance_tmp", SanitizeFileName(baseName));
        Directory.CreateDirectory(workRoot);

        var narrationWorkWav = Path.Combine(workRoot, "narration_input.wav");
        var mixedWorkWav = Path.Combine(workRoot, "mixed_output.wav");
        var narrationOnlyPath = Path.Combine(outputDir, $"{baseName}.narration_only.wav");

        try
        {
            progress?.Report(new AudioEnhanceProgress { Stage = "Preparing narration...", Completed = 0, Total = 1 });
            await EnsureWavInputAsync(outputPath, narrationWorkWav, ct);
            result.NarrationOnlyPath = narrationOnlyPath;

            if (_config.AudioEnhanceExportNarrationOnly)
            {
                File.Copy(narrationWorkWav, narrationOnlyPath, overwrite: true);
            }

            if (!TryGetAudioDurationSeconds(narrationWorkWav, out var totalSeconds) || totalSeconds < 0.2)
            {
                throw new InvalidOperationException("Narration audio duration is invalid for enhancement.");
            }

            progress?.Report(new AudioEnhanceProgress { Stage = "Extracting cues...", Completed = 0, Total = 1 });
            var extractionSettings = new EnhanceCueExtractionSettings
            {
                CueMaxPerMinute = Math.Clamp(_config.AudioEnhanceCueMaxPerMinute, 1, 60),
                UseLlmRefine = _project.AudioEnhanceUseLlmCueRefine && _config.AudioEnhanceUseLlmCueRefine,
                MinAmbienceHoldSeconds = 6.0,
                OneShotCooldownSeconds = 2.5
            };

            var units = BuildEnhanceTimedUnits(sourceText, chunkTimings, totalSeconds, out var usedEstimatedTiming);
            var cues = ExtractRuleCues(units, totalSeconds, extractionSettings);
            result.UsedEstimatedTiming = usedEstimatedTiming;
            List<EnhanceCue> manualPreparedCues = new();
            HashSet<int>? excludedUnitIndexes = null;

            if (preparedScript is not null && chunkTimings is { Count: > 0 })
            {
                manualPreparedCues = BuildPreparedScriptManualSfxCues(preparedScript, chunkTimings, totalSeconds, out excludedUnitIndexes);
                if (excludedUnitIndexes is { Count: > 0 })
                {
                    var autoUnits = units
                        .Where((_, idx) => !excludedUnitIndexes.Contains(idx))
                        .ToList();
                    cues = ExtractRuleCues(autoUnits, totalSeconds, extractionSettings);
                }
            }

            if (_config.AudioEnhanceUseLlmCueDetection)
            {
                progress?.Report(new AudioEnhanceProgress { Stage = "Detecting cues with LLM...", Completed = 0, Total = 1 });
                var llmUnits = excludedUnitIndexes is { Count: > 0 }
                    ? units.Where((_, idx) => !excludedUnitIndexes.Contains(idx)).ToList()
                    : units;
                var detected = await TryDetectCuesWithLlmAsync(llmUnits, totalSeconds, extractionSettings, ct);
                if (detected.Count > 0)
                {
                    cues = detected;
                }
            }

            if (manualPreparedCues.Count > 0)
            {
                cues.AddRange(manualPreparedCues);
            }

            ApplyCueOverrides(cues);

            if (extractionSettings.UseLlmRefine && cues.Count > 0)
            {
                var refined = await TryRefineCuesAsync(cues, units, totalSeconds, ct);
                if (refined.Count > 0)
                {
                    cues = refined;
                    result.UsedLlmRefine = true;
                }
            }

            cues = cues
                .Where(c => string.Equals(c.CueType, "oneshot", StringComparison.OrdinalIgnoreCase))
                .OrderBy(c => c.StartSeconds)
                .ToList();

            result.Cues = cues.ToList();
            if (cues.Count == 0)
            {
                result.Success = true;
                result.Message = "No sound-effect cues detected, kept narration-only output.";
                await WriteEnhanceOutputAsync(narrationWorkWav, mixedWorkWav, outputPath, ext, ct);
                return result;
            }

            var cacheRoot = Path.Combine(outputDir, ".audiox_cache");
            Directory.CreateDirectory(cacheRoot);

            var timeline = new List<EnhanceTimelineEvent>(cues.Count);
            for (var i = 0; i < cues.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var cue = cues[i];
                progress?.Report(new AudioEnhanceProgress
                {
                    Stage = "Generating SFX...",
                    Completed = i,
                    Total = cues.Count
                });

                var duration = Math.Clamp(cue.EndSeconds - cue.StartSeconds, 1.0, cue.CueType == "ambience" ? 8.0 : 3.0);
                var clipPath = await EnsureCueClipAsync(cue, duration, cacheRoot, ct);
                if (string.IsNullOrWhiteSpace(clipPath) || !File.Exists(clipPath))
                {
                    continue;
                }

                timeline.Add(new EnhanceTimelineEvent
                {
                    CueId = cue.Id,
                    CueType = cue.CueType,
                    Prompt = cue.Prompt,
                    AudioPath = clipPath,
                    StartSeconds = cue.StartSeconds,
                    EndSeconds = cue.EndSeconds,
                    Intensity = cue.Intensity
                });
            }

            progress?.Report(new AudioEnhanceProgress { Stage = "Mixing...", Completed = 0, Total = 1 });
            var stems = await MixEnhanceTracksAsync(narrationWorkWav, timeline, units, mixedWorkWav, outputDir, baseName, ct);
            result.Stems = stems;

            await WriteEnhanceOutputAsync(mixedWorkWav, mixedWorkWav, outputPath, ext, ct);
            result.Success = true;
            result.FinalOutputPath = outputPath;
            result.Message = "Enhancement completed.";
            return result;
        }
        catch (Exception ex)
        {
            try
            {
                if (File.Exists(narrationWorkWav))
                {
                    await WriteEnhanceOutputAsync(narrationWorkWav, narrationWorkWav, outputPath, ext, ct);
                }
            }
            catch
            {
                // Keep original generation result; enhance remains fail-soft.
            }
            result.Success = false;
            result.Message = $"Enhance warning: {ex.Message}";
            return result;
        }
        finally
        {
            try
            {
                if (Directory.Exists(workRoot))
                {
                    Directory.Delete(workRoot, recursive: true);
                }
            }
            catch
            {
                // Non-fatal temp cleanup.
            }
        }
    }

    private static List<EnhanceCue> BuildPreparedScriptManualSfxCues(
        PreparedScriptDocument preparedScript,
        IReadOnlyList<SubtitleChunkTiming> chunkTimings,
        double totalSeconds,
        out HashSet<int> excludedUnitIndexes)
    {
        excludedUnitIndexes = new HashSet<int>();
        var cues = new List<EnhanceCue>();
        var validParts = (preparedScript.Parts ?? new List<PreparedScriptPart>())
            .Where(p => !string.IsNullOrWhiteSpace((p.Text ?? string.Empty).Trim()))
            .OrderBy(p => p.Order)
            .ToList();

        var count = Math.Min(validParts.Count, chunkTimings.Count);
        for (var i = 0; i < count; i++)
        {
            var part = validParts[i];
            if (!part.DisableAutoSfxDetection)
            {
                continue;
            }

            excludedUnitIndexes.Add(i);
            if (string.IsNullOrWhiteSpace(part.SoundEffectPrompt))
            {
                continue;
            }

            var timing = chunkTimings[i];
            var span = Math.Max(0.35, timing.EndSeconds - timing.StartSeconds);
            var start = Math.Max(0, timing.StartSeconds + (span * 0.35));
            var end = Math.Min(totalSeconds, start + Math.Min(1.4, Math.Max(0.45, span * 0.35)));
            cues.Add(new EnhanceCue
            {
                Id = $"manualsfx_{HashStable($"{part.Id}|{start:0.###}|{part.SoundEffectPrompt}")}",
                StartSeconds = start,
                EndSeconds = Math.Max(start + 0.25, end),
                CueType = "oneshot",
                Prompt = part.SoundEffectPrompt.Trim(),
                Intensity = 0.85,
                Source = "prepared-manual"
            });
        }

        return cues;
    }

    private async Task<EnhanceResult> RunAudioEnhanceForExternalAudioAsync(
        string inputAudioPath,
        string outputPath,
        string cueText,
        IReadOnlyList<EnhanceCue>? explicitCues,
        IReadOnlyList<TimedTextUnit>? explicitUnits,
        CancellationToken ct,
        IProgress<AudioEnhanceProgress>? progress = null)
    {
        var result = new EnhanceResult
        {
            Success = false,
            FinalOutputPath = outputPath
        };

        var outputDir = Path.GetDirectoryName(outputPath) ?? RuntimePaths.AppRoot;
        var baseName = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath).Trim().TrimStart('.').ToLowerInvariant();

        var workRoot = Path.Combine(outputDir, ".audio_enhance_tmp", $"ext_{SanitizeFileName(baseName)}");
        Directory.CreateDirectory(workRoot);

        var narrationWorkWav = Path.Combine(workRoot, "input_audio.wav");
        var mixedWorkWav = Path.Combine(workRoot, "mixed_output.wav");
        var narrationOnlyPath = Path.Combine(outputDir, $"{baseName}.narration_only.wav");

        try
        {
            progress?.Report(new AudioEnhanceProgress { Stage = "Preparing narration...", Completed = 0, Total = 1 });
            await EnsureWavInputAsync(inputAudioPath, narrationWorkWav, ct);
            result.NarrationOnlyPath = narrationOnlyPath;

            if (_config.AudioEnhanceExportNarrationOnly)
            {
                File.Copy(narrationWorkWav, narrationOnlyPath, overwrite: true);
            }

            if (!TryGetAudioDurationSeconds(narrationWorkWav, out var totalSeconds) || totalSeconds < 0.2)
            {
                throw new InvalidOperationException("Input audio duration is invalid for enhancement.");
            }

            var sourceText = string.IsNullOrWhiteSpace(cueText)
                ? Path.GetFileNameWithoutExtension(inputAudioPath).Replace('_', ' ').Replace('-', ' ')
                : cueText;

            progress?.Report(new AudioEnhanceProgress { Stage = "Extracting cues...", Completed = 0, Total = 1 });
            var extractionSettings = new EnhanceCueExtractionSettings
            {
                CueMaxPerMinute = Math.Clamp(_config.AudioEnhanceCueMaxPerMinute, 1, 60),
                UseLlmRefine = _config.AudioEnhanceUseLlmCueRefine,
                MinAmbienceHoldSeconds = 6.0,
                OneShotCooldownSeconds = 2.5
            };

            var units = explicitUnits?.Where(u => u.EndSeconds > u.StartSeconds)
                .OrderBy(u => u.StartSeconds)
                .ToList()
                ?? BuildEnhanceTimedUnits(sourceText, null, totalSeconds, out _);
            result.UsedEstimatedTiming = explicitUnits is null || explicitUnits.Count == 0;

            List<EnhanceCue> cues;
            if (explicitCues is { Count: > 0 })
            {
                cues = explicitCues
                    .Select(c => new EnhanceCue
                    {
                        Id = string.IsNullOrWhiteSpace(c.Id) ? Guid.NewGuid().ToString("N") : c.Id,
                        StartSeconds = Math.Clamp(c.StartSeconds, 0, totalSeconds),
                        EndSeconds = Math.Clamp(c.EndSeconds, 0, totalSeconds),
                        CueType = string.Equals(c.CueType, "oneshot", StringComparison.OrdinalIgnoreCase) ? "oneshot" : "ambience",
                        Prompt = c.Prompt ?? string.Empty,
                        Intensity = Math.Clamp(c.Intensity <= 0 ? 0.6 : c.Intensity, 0.1, 1.0),
                        Source = string.IsNullOrWhiteSpace(c.Source) ? "prepared" : c.Source
                    })
                    .Where(c => !string.IsNullOrWhiteSpace(c.Prompt))
                    .Select(c =>
                    {
                        if (c.EndSeconds <= c.StartSeconds)
                        {
                            c.EndSeconds = Math.Min(totalSeconds, c.StartSeconds + (c.CueType == "oneshot" ? 1.2 : 3.0));
                        }
                        return c;
                    })
                    .Where(c => c.EndSeconds > c.StartSeconds)
                    .OrderBy(c => c.StartSeconds)
                    .ToList();
                result.UsedEstimatedTiming = false;
            }
            else
            {
                cues = ExtractRuleCues(units, totalSeconds, extractionSettings);
                if (_config.AudioEnhanceUseLlmCueDetection)
                {
                    progress?.Report(new AudioEnhanceProgress { Stage = "Detecting cues with LLM...", Completed = 0, Total = 1 });
                    var detected = await TryDetectCuesWithLlmAsync(units, totalSeconds, extractionSettings, ct);
                    if (detected.Count > 0)
                    {
                        cues = detected;
                    }
                }

                if (cues.Count == 0)
                {
                    cues.Add(new EnhanceCue
                    {
                        Id = $"amb_{HashStable($"{sourceText}|{totalSeconds:0.###}")}",
                        StartSeconds = 0,
                        EndSeconds = totalSeconds,
                        CueType = "ambience",
                        Prompt = "subtle cinematic ambience bed, low noise, natural room tone",
                        Intensity = 0.50,
                        Source = "fallback"
                    });
                }

                if (extractionSettings.UseLlmRefine && cues.Count > 0)
                {
                    var refined = await TryRefineCuesAsync(cues, units, totalSeconds, ct);
                    if (refined.Count > 0)
                    {
                        cues = refined;
                        result.UsedLlmRefine = true;
                    }
                }
            }

            result.Cues = cues.ToList();

            var cacheRoot = Path.Combine(outputDir, ".audiox_cache");
            Directory.CreateDirectory(cacheRoot);
            var timeline = new List<EnhanceTimelineEvent>(cues.Count);
            for (var i = 0; i < cues.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var cue = cues[i];
                progress?.Report(new AudioEnhanceProgress
                {
                    Stage = "Generating SFX...",
                    Completed = i,
                    Total = cues.Count
                });

                var duration = Math.Clamp(cue.EndSeconds - cue.StartSeconds, 1.0, cue.CueType == "ambience" ? 8.0 : 3.0);
                var clipPath = await EnsureCueClipAsync(cue, duration, cacheRoot, ct);
                if (string.IsNullOrWhiteSpace(clipPath) || !File.Exists(clipPath))
                {
                    continue;
                }

                timeline.Add(new EnhanceTimelineEvent
                {
                    CueId = cue.Id,
                    CueType = cue.CueType,
                    Prompt = cue.Prompt,
                    AudioPath = clipPath,
                    StartSeconds = cue.StartSeconds,
                    EndSeconds = cue.EndSeconds,
                    Intensity = cue.Intensity
                });
            }

            progress?.Report(new AudioEnhanceProgress { Stage = "Mixing...", Completed = 0, Total = 1 });
            var stems = await MixEnhanceTracksAsync(narrationWorkWav, timeline, units, mixedWorkWav, outputDir, baseName, ct);
            result.Stems = stems;

            await WriteEnhanceOutputAsync(mixedWorkWav, mixedWorkWav, outputPath, ext, ct);
            result.Success = true;
            result.FinalOutputPath = outputPath;
            result.Message = "External audio enhancement completed.";
            return result;
        }
        catch (Exception ex)
        {
            try
            {
                if (File.Exists(narrationWorkWav))
                {
                    await WriteEnhanceOutputAsync(narrationWorkWav, narrationWorkWav, outputPath, ext, ct);
                }
            }
            catch
            {
                // Keep original source untouched on fallback failure.
            }

            result.Success = false;
            result.Message = $"Enhance warning: {ex.Message}";
            return result;
        }
        finally
        {
            try
            {
                if (Directory.Exists(workRoot))
                {
                    Directory.Delete(workRoot, recursive: true);
                }
            }
            catch
            {
                // Non-fatal temp cleanup.
            }
        }
    }

    private async Task WriteEnhanceOutputAsync(
        string sourceWavPath,
        string mixedWavPath,
        string outputPath,
        string outputExt,
        CancellationToken ct)
    {
        var input = File.Exists(mixedWavPath) ? mixedWavPath : sourceWavPath;
        if (string.Equals(outputExt, "wav", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(outputExt))
        {
            File.Copy(input, outputPath, overwrite: true);
            return;
        }

        var ffmpeg = ResolveFfmpegExecutable();
        if (string.IsNullOrWhiteSpace(ffmpeg))
        {
            File.Copy(input, outputPath, overwrite: true);
            return;
        }

        var psi = new ProcessStartInfo
        {
            FileName = ffmpeg,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-y");
        psi.ArgumentList.Add("-i");
        psi.ArgumentList.Add(input);
        psi.ArgumentList.Add("-vn");
        if (string.Equals(outputExt, "mp3", StringComparison.OrdinalIgnoreCase))
        {
            psi.ArgumentList.Add("-c:a");
            psi.ArgumentList.Add("libmp3lame");
            psi.ArgumentList.Add("-q:a");
            psi.ArgumentList.Add("2");
        }
        psi.ArgumentList.Add(outputPath);

        using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start ffmpeg process.");
        using var reg = ct.Register(() =>
        {
            try
            {
                if (!proc.HasExited)
                {
                    proc.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // Ignore cancellation race.
            }
        });
        var stderr = await proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync(ct);
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException($"ffmpeg output conversion failed: {stderr}");
        }
    }

    private async Task EnsureWavInputAsync(string sourcePath, string targetWavPath, CancellationToken ct)
    {
        if (string.Equals(Path.GetExtension(sourcePath), ".wav", StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(sourcePath, targetWavPath, overwrite: true);
            return;
        }

        var ffmpeg = ResolveFfmpegExecutable();
        if (string.IsNullOrWhiteSpace(ffmpeg))
        {
            throw new InvalidOperationException("ffmpeg is required for non-WAV enhance input.");
        }

        var psi = new ProcessStartInfo
        {
            FileName = ffmpeg,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-y");
        psi.ArgumentList.Add("-i");
        psi.ArgumentList.Add(sourcePath);
        psi.ArgumentList.Add("-acodec");
        psi.ArgumentList.Add("pcm_s16le");
        psi.ArgumentList.Add(targetWavPath);

        using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start ffmpeg process.");
        using var reg = ct.Register(() =>
        {
            try
            {
                if (!proc.HasExited)
                {
                    proc.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // Ignore cancellation race.
            }
        });
        var stderr = await proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync(ct);
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException($"ffmpeg input conversion failed: {stderr}");
        }
    }

    private List<TimedTextUnit> BuildEnhanceTimedUnits(
        string sourceText,
        IReadOnlyList<SubtitleChunkTiming>? chunkTimings,
        double totalSeconds,
        out bool usedEstimatedTiming)
    {
        var subtitleCues = BuildSubtitleCues(sourceText, chunkTimings, totalSeconds);
        usedEstimatedTiming = chunkTimings is null || chunkTimings.Count == 0;
        return subtitleCues
            .Where(c => !string.IsNullOrWhiteSpace(c.Text))
            .Select(c => new TimedTextUnit(c.Text, c.StartSeconds, c.EndSeconds))
            .ToList();
    }

    private List<EnhanceCue> ExtractRuleCues(
        IReadOnlyList<TimedTextUnit> units,
        double totalSeconds,
        EnhanceCueExtractionSettings settings)
    {
        var cues = new List<EnhanceCue>();
        if (units.Count == 0)
        {
            return cues;
        }

        var ambienceCandidates = new List<(double Start, string Prompt, double Intensity)>();
        var oneShotLastAt = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        foreach (var unit in units)
        {
            var lower = (unit.Text ?? string.Empty).ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(lower))
            {
                continue;
            }

            foreach (var rule in AmbienceCueRules)
            {
                if (rule.Keywords.Any(k => ContainsCueKeyword(lower, k)))
                {
                    var refinedPrompt = RefineAmbiencePrompt(rule.Prompt, lower);
                    ambienceCandidates.Add((unit.StartSeconds, refinedPrompt, 0.55));
                    break;
                }
            }

            foreach (var rule in OneShotCueRules)
            {
                if (!rule.Keywords.Any(k => ContainsCueKeyword(lower, k)))
                {
                    continue;
                }

                var refinedPrompt = RefineOneShotPrompt(rule.Prompt, lower);
                var start = Math.Max(0, unit.StartSeconds + ((unit.EndSeconds - unit.StartSeconds) * 0.35));
                var end = Math.Min(totalSeconds, start + 1.2);
                if (oneShotLastAt.TryGetValue(refinedPrompt, out var last) &&
                    start - last < settings.OneShotCooldownSeconds)
                {
                    continue;
                }

                oneShotLastAt[refinedPrompt] = start;
                cues.Add(new EnhanceCue
                {
                    Id = $"oneshot_{HashStable($"{refinedPrompt}|{start:0.###}")}",
                    StartSeconds = start,
                    EndSeconds = Math.Max(start + 0.25, end),
                    CueType = "oneshot",
                    Prompt = refinedPrompt,
                    Intensity = 0.8,
                    Source = "rules"
                });
            }

            foreach (var regexRule in OneShotCueRegexRules)
            {
                if (!regexRule.Pattern.IsMatch(unit.Text ?? string.Empty))
                {
                    continue;
                }

                var refinedPrompt = RefineOneShotPrompt(regexRule.Prompt, lower);
                var start = Math.Max(0, unit.StartSeconds + ((unit.EndSeconds - unit.StartSeconds) * 0.28));
                var end = Math.Min(totalSeconds, start + 1.0);
                if (oneShotLastAt.TryGetValue(refinedPrompt, out var last) &&
                    start - last < settings.OneShotCooldownSeconds)
                {
                    continue;
                }

                oneShotLastAt[refinedPrompt] = start;
                cues.Add(new EnhanceCue
                {
                    Id = $"oneshot_{HashStable($"{refinedPrompt}|{start:0.###}")}",
                    StartSeconds = start,
                    EndSeconds = Math.Max(start + 0.25, end),
                    CueType = "oneshot",
                    Prompt = refinedPrompt,
                    Intensity = 0.85,
                    Source = "rules"
                });
            }
        }

        ambienceCandidates = ambienceCandidates
            .OrderBy(a => a.Start)
            .ToList();

        var filteredAmbience = new List<(double Start, string Prompt, double Intensity)>();
        double lastAmbienceAt = double.NegativeInfinity;
        foreach (var candidate in ambienceCandidates)
        {
            if (candidate.Start - lastAmbienceAt < settings.MinAmbienceHoldSeconds)
            {
                continue;
            }
            filteredAmbience.Add(candidate);
            lastAmbienceAt = candidate.Start;
        }

        for (var i = 0; i < filteredAmbience.Count; i++)
        {
            var start = Math.Max(0, filteredAmbience[i].Start);
            var nextStart = i < filteredAmbience.Count - 1 ? filteredAmbience[i + 1].Start : totalSeconds;
            var end = Math.Min(totalSeconds, Math.Max(start + 1.0, nextStart));
            cues.Add(new EnhanceCue
            {
                Id = $"amb_{HashStable($"{filteredAmbience[i].Prompt}|{start:0.###}")}",
                StartSeconds = start,
                EndSeconds = end,
                CueType = "ambience",
                Prompt = filteredAmbience[i].Prompt,
                Intensity = filteredAmbience[i].Intensity,
                Source = "rules"
            });
        }

        var maxCues = Math.Max(1, (int)Math.Ceiling((Math.Max(totalSeconds, 1) / 60.0) * settings.CueMaxPerMinute));
        if (cues.Count > maxCues)
        {
            cues = cues
                .OrderBy(c => c.CueType == "ambience" ? 0 : 1)
                .ThenBy(c => c.StartSeconds)
                .Take(maxCues)
                .ToList();
        }

        return cues
            .OrderBy(c => c.StartSeconds)
            .ToList();
    }

    private static bool ContainsCueKeyword(string text, string keyword)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(keyword))
        {
            return false;
        }

        keyword = keyword.Trim().ToLowerInvariant();
        if (keyword.IndexOf(' ') >= 0)
        {
            return text.Contains(keyword, StringComparison.Ordinal);
        }

        return Regex.IsMatch(
            text,
            $@"\b{Regex.Escape(keyword)}\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static string RefineAmbiencePrompt(string basePrompt, string lowerText)
    {
        if (string.IsNullOrWhiteSpace(basePrompt))
        {
            return basePrompt;
        }

        if (basePrompt.Contains("urban ambience", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "market", "bazaar"))
            {
                return "busy market ambience, crowd murmur, distant stalls, soft city bed";
            }
            if (ContainsAnyCueKeyword(lowerText, "alley", "backstreet"))
            {
                return "quiet city alley ambience, distant urban hum, sparse movement";
            }
        }

        if (basePrompt.Contains("forest ambience", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "night", "midnight", "dark"))
            {
                return "night forest ambience, soft leaves, distant nocturnal insects, dark natural space";
            }
            if (ContainsAnyCueKeyword(lowerText, "river", "stream"))
            {
                return "forest ambience with nearby running water, soft leaves, natural outdoor bed";
            }
        }

        if (basePrompt.Contains("indoor room tone", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "library"))
            {
                return "quiet library ambience, paper movement, soft indoor room tone";
            }
            if (ContainsAnyCueKeyword(lowerText, "office"))
            {
                return "soft office room tone, restrained indoor ambience, faint interior air";
            }
        }

        return basePrompt;
    }

    private static string RefineOneShotPrompt(string basePrompt, string lowerText)
    {
        if (string.IsNullOrWhiteSpace(basePrompt))
        {
            return basePrompt;
        }

        if (basePrompt.Contains("footsteps", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "stone", "stone floor", "cobblestone"))
            {
                return "heavy boots on stone one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "wood", "wood floor", "floorboards"))
            {
                return "footsteps on wood floor one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "gravel", "pebbles"))
            {
                return "footsteps on gravel one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "mud", "wet ground"))
            {
                return "wet muddy footsteps one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "snow", "ice"))
            {
                return "boots crunching on snow one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "water", "puddle"))
            {
                return "footstep splash one-shot";
            }
        }

        if (basePrompt.Contains("door", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "metal", "iron"))
            {
                return basePrompt.Contains("slam", StringComparison.OrdinalIgnoreCase)
                    ? "heavy metal door slam one-shot"
                    : "metal door movement one-shot";
            }

            if (ContainsAnyCueKeyword(lowerText, "gate"))
            {
                return "heavy gate movement one-shot";
            }
        }

        if (basePrompt.Contains("blade", StringComparison.OrdinalIgnoreCase) ||
            basePrompt.Contains("sword clash", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "dagger", "knife"))
            {
                return "fast dagger slash one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "axe"))
            {
                return "heavy axe swing impact one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "spear", "lance"))
            {
                return "piercing spear thrust one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "draw", "drew", "sheath", "unsheathed"))
            {
                return "metal blade unsheathing one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "parried", "blocked", "deflected"))
            {
                return "sharp metal parry clash one-shot";
            }
        }

        if (basePrompt.Contains("body hit", StringComparison.OrdinalIgnoreCase) ||
            basePrompt.Contains("impact", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "wall", "pillar"))
            {
                return "body slammed into wall one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "ground", "floor"))
            {
                return "body hitting ground one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "table", "desk", "chair"))
            {
                return "body crashing into furniture one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "punch", "fist"))
            {
                return "heavy punch impact one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "kick", "kneed", "elbowed"))
            {
                return "hard body strike one-shot";
            }
        }

        if (basePrompt.Contains("mechanical", StringComparison.OrdinalIgnoreCase) ||
            basePrompt.Contains("device alert", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "lock"))
            {
                return "small lock click one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "phone", "terminal", "screen"))
            {
                return "short device notification one-shot";
            }
        }

        if (basePrompt.Contains("creature vocal", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "dragon"))
            {
                return "distant dragon roar one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "wolf"))
            {
                return "wolf growl one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "snake", "serpent"))
            {
                return "sharp snake hiss one-shot";
            }
        }

        if (basePrompt.Contains("wood scrape", StringComparison.OrdinalIgnoreCase))
        {
            if (ContainsAnyCueKeyword(lowerText, "chair"))
            {
                return "chair scraping floor one-shot";
            }
            if (ContainsAnyCueKeyword(lowerText, "floorboards", "boards"))
            {
                return "creaking floorboards one-shot";
            }
        }

        return basePrompt;
    }

    private static bool ContainsAnyCueKeyword(string text, params string[] keywords)
    {
        foreach (var keyword in keywords)
        {
            if (ContainsCueKeyword(text, keyword))
            {
                return true;
            }
        }

        return false;
    }

    private void ApplyCueOverrides(List<EnhanceCue> cues)
    {
        if (_project.CueOverrides is null || _project.CueOverrides.Count == 0 || cues.Count == 0)
        {
            return;
        }

        var byId = cues.ToDictionary(c => c.Id, StringComparer.OrdinalIgnoreCase);
        foreach (var ov in _project.CueOverrides)
        {
            if (string.IsNullOrWhiteSpace(ov.CueId) || !byId.TryGetValue(ov.CueId, out var cue))
            {
                continue;
            }

            if (ov.Disabled)
            {
                cue.EndSeconds = cue.StartSeconds;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(ov.Prompt))
            {
                cue.Prompt = ov.Prompt.Trim();
            }
            if (ov.StartSeconds.HasValue)
            {
                cue.StartSeconds = Math.Max(0, ov.StartSeconds.Value);
            }
            if (ov.EndSeconds.HasValue)
            {
                cue.EndSeconds = Math.Max(cue.StartSeconds + 0.1, ov.EndSeconds.Value);
            }
            if (ov.Intensity.HasValue)
            {
                cue.Intensity = Math.Clamp(ov.Intensity.Value, 0.1, 1.0);
            }
        }

        cues.RemoveAll(c => c.EndSeconds <= c.StartSeconds);
    }

    private async Task<List<EnhanceCue>> TryDetectCuesWithLlmAsync(
        IReadOnlyList<TimedTextUnit> units,
        double totalSeconds,
        EnhanceCueExtractionSettings settings,
        CancellationToken ct)
    {
        try
        {
            if (units.Count == 0)
            {
                return new List<EnhanceCue>();
            }

            var provider = (_config.LlmPrepProvider ?? string.Empty).Trim().ToLowerInvariant();
            if (provider == "openai" && string.IsNullOrWhiteSpace(_config.ApiKeyOpenAi))
            {
                return new List<EnhanceCue>();
            }
            if (provider == "alibaba" && string.IsNullOrWhiteSpace(_config.ApiKeyAlibaba))
            {
                return new List<EnhanceCue>();
            }

            var client = new ScriptPrepLlmClient(_config);
            var chunks = BuildEnhanceCueDetectionChunks(units, provider);
            var detected = new List<EnhanceCue>();
            foreach (var chunk in chunks)
            {
                ct.ThrowIfCancellationRequested();
                var prompt = BuildEnhanceCueDetectionPrompt(chunk, totalSeconds);
                var content = await client.RunCustomPromptAsync(
                    prompt,
                    useInstructionModel: true,
                    temperature: Math.Min(_config.LlmPrepTemperatureInstruction, 0.3),
                    maxTokens: Math.Clamp(_config.LlmPrepMaxTokensInstruction * 2, 256, 2200),
                    ct);

                var parsed = ParseRefinedCueLines(content, Array.Empty<EnhanceCue>());
                if (parsed.Count == 0)
                {
                    continue;
                }

                detected.AddRange(parsed);
            }

            if (detected.Count == 0)
            {
                return new List<EnhanceCue>();
            }

            return detected
                .Where(c => c.EndSeconds > c.StartSeconds && !string.IsNullOrWhiteSpace(c.Prompt))
                .GroupBy(c => $"{c.CueType}|{Math.Round(c.StartSeconds, 1):0.0}|{c.Prompt}", StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(c => c.StartSeconds)
                .Take(Math.Max(1, (int)Math.Ceiling((Math.Max(totalSeconds, 1) / 60.0) * settings.CueMaxPerMinute)))
                .ToList();
        }
        catch
        {
            return new List<EnhanceCue>();
        }
    }

    private async Task<List<EnhanceCue>> TryRefineCuesAsync(
        IReadOnlyList<EnhanceCue> cues,
        IReadOnlyList<TimedTextUnit> units,
        double totalSeconds,
        CancellationToken ct)
    {
        try
        {
            var provider = (_config.LlmPrepProvider ?? string.Empty).Trim().ToLowerInvariant();
            if (provider == "openai" && string.IsNullOrWhiteSpace(_config.ApiKeyOpenAi))
            {
                return cues.ToList();
            }
            if (provider == "alibaba" && string.IsNullOrWhiteSpace(_config.ApiKeyAlibaba))
            {
                return cues.ToList();
            }
            var client = new ScriptPrepLlmClient(_config);
            var prompt = BuildEnhanceCueRefinePrompt(cues, units, totalSeconds);
            var content = await client.RunCustomPromptAsync(
                prompt,
                useInstructionModel: true,
                temperature: Math.Min(_config.LlmPrepTemperatureInstruction, 0.25),
                maxTokens: Math.Clamp(_config.LlmPrepMaxTokensInstruction * 2, 256, 2200),
                ct);
            var refined = ParseRefinedCueLines(content, cues);
            return refined.Count == 0 ? cues.ToList() : refined;
        }
        catch
        {
            return cues.ToList();
        }
    }

    private static List<List<TimedTextUnit>> BuildEnhanceCueDetectionChunks(IReadOnlyList<TimedTextUnit> units, string provider)
    {
        var targetChars = string.Equals(provider, "local", StringComparison.OrdinalIgnoreCase) ? 2600 : 5200;
        var targetLines = string.Equals(provider, "local", StringComparison.OrdinalIgnoreCase) ? 10 : 18;
        var chunks = new List<List<TimedTextUnit>>();
        var current = new List<TimedTextUnit>();
        var chars = 0;

        foreach (var unit in units.Where(u => !string.IsNullOrWhiteSpace(u.Text)))
        {
            var lineLen = (unit.Text?.Length ?? 0) + 32;
            if (current.Count > 0 && (current.Count >= targetLines || chars + lineLen > targetChars))
            {
                chunks.Add(current);
                current = new List<TimedTextUnit>();
                chars = 0;
            }

            current.Add(unit);
            chars += lineLen;
        }

        if (current.Count > 0)
        {
            chunks.Add(current);
        }

        return chunks;
    }

    private static string BuildEnhanceCueDetectionPrompt(IReadOnlyList<TimedTextUnit> units, double totalSeconds)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Task: Detect audiobook background ambience and one-shot sound effects from timed text lines.");
        sb.AppendLine("Return plain text only. No markdown.");
        sb.AppendLine("Output one cue per line in this exact format:");
        sb.AppendLine("start_sec|end_sec|cue_type|prompt|intensity");
        sb.AppendLine("cue_type must be ambience or oneshot.");
        sb.AppendLine("If a line needs no sound, do not output anything for that line.");
        sb.AppendLine("Only create cues when the text clearly implies audible environment or a discrete sound event.");
        sb.AppendLine("Keep prompts short, practical, and generation-ready.");
        sb.AppendLine("Intensity range is 0.10 to 1.00.");
        sb.AppendLine($"Total duration: {totalSeconds:0.###}");
        sb.AppendLine();
        sb.AppendLine("Timed lines:");
        foreach (var unit in units)
        {
            sb.AppendLine($"{unit.StartSeconds:0.###}|{unit.EndSeconds:0.###}|{unit.Text}");
        }
        return sb.ToString();
    }

    private static string BuildEnhanceCueRefinePrompt(
        IReadOnlyList<EnhanceCue> cues,
        IReadOnlyList<TimedTextUnit> units,
        double totalSeconds)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Task: Refine these audiobook ambience and sound effect cues.");
        sb.AppendLine("Return plain text only, exact format per line:");
        sb.AppendLine("start_sec|end_sec|cue_type|prompt|intensity");
        sb.AppendLine("cue_type must be ambience or oneshot.");
        sb.AppendLine("Intensity range is 0.10 to 1.00.");
        sb.AppendLine("Do not return markdown.");
        sb.AppendLine($"Total duration: {totalSeconds:0.###}");
        sb.AppendLine();
        sb.AppendLine("Current cues:");
        foreach (var cue in cues.OrderBy(c => c.StartSeconds))
        {
            sb.AppendLine($"{cue.StartSeconds:0.###}|{cue.EndSeconds:0.###}|{cue.CueType}|{cue.Prompt}|{cue.Intensity:0.##}");
        }
        sb.AppendLine();
        sb.AppendLine("Context lines:");
        foreach (var unit in units.Take(40))
        {
            sb.AppendLine($"{unit.StartSeconds:0.###}-{unit.EndSeconds:0.###}: {unit.Text}");
        }
        return sb.ToString();
    }

    private static List<EnhanceCue> ParseRefinedCueLines(string content, IReadOnlyList<EnhanceCue> fallback)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return fallback.ToList();
        }

        var result = new List<EnhanceCue>();
        var lines = content
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim())
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();
        foreach (var line in lines)
        {
            var m = Regex.Match(
                line,
                @"^\s*(\d+(?:\.\d+)?)\|(\d+(?:\.\d+)?)\|(ambience|oneshot)\|(.+?)\|([01](?:\.\d+)?)\s*$",
                RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                continue;
            }

            var start = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var end = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            var type = m.Groups[3].Value.ToLowerInvariant();
            var prompt = m.Groups[4].Value.Trim();
            var intensity = Math.Clamp(double.Parse(m.Groups[5].Value, CultureInfo.InvariantCulture), 0.1, 1.0);
            if (end <= start || string.IsNullOrWhiteSpace(prompt))
            {
                continue;
            }

            result.Add(new EnhanceCue
            {
                Id = $"{type}_{HashStable($"{prompt}|{start:0.###}|{end:0.###}")}",
                StartSeconds = start,
                EndSeconds = end,
                CueType = type,
                Prompt = prompt,
                Intensity = intensity,
                Source = "llm"
            });
        }

        return result.OrderBy(c => c.StartSeconds).ToList();
    }

    private async Task<string?> EnsureCueClipAsync(EnhanceCue cue, double durationSeconds, string cacheRoot, CancellationToken ct)
    {
        var hash = HashStable($"{cue.CueType}|{cue.Prompt}|{durationSeconds:0.###}|{cue.Intensity:0.##}|{_config.AudioEnhanceVariant}");
        var path = Path.Combine(cacheRoot, $"{hash}.wav");
        if (File.Exists(path) && new FileInfo(path).Length > 512)
        {
            return path;
        }

        var python = ResolveAudioEnhancePython();
        if (string.IsNullOrWhiteSpace(python))
        {
            throw new InvalidOperationException("AudioX runtime not found. Configure Audio Enhancement runtime in Settings.");
        }

        var script = EnsureAudioXScript();
        var modelDir = ResolveAudioEnhanceModelDir();
        Directory.CreateDirectory(modelDir);

        var psi = new ProcessStartInfo
        {
            FileName = python,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add(script);
        psi.ArgumentList.Add("--model-dir");
        psi.ArgumentList.Add(modelDir);
        psi.ArgumentList.Add("--variant");
        psi.ArgumentList.Add(string.IsNullOrWhiteSpace(_config.AudioEnhanceVariant) ? "audiox_base" : _config.AudioEnhanceVariant);
        psi.ArgumentList.Add("--prompt");
        psi.ArgumentList.Add(cue.Prompt);
        psi.ArgumentList.Add("--seconds");
        psi.ArgumentList.Add(durationSeconds.ToString("0.###", CultureInfo.InvariantCulture));
        psi.ArgumentList.Add("--output");
        psi.ArgumentList.Add(path);

        using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start AudioX generation process.");
        using var reg = ct.Register(() =>
        {
            try
            {
                if (!proc.HasExited)
                {
                    proc.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // Ignore cancellation race.
            }
        });

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync(ct);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException($"AudioX cue generation failed: {TrimForError(stderr)} {TrimForError(stdout)}");
        }

        if (!File.Exists(path) || new FileInfo(path).Length < 512)
        {
            throw new InvalidOperationException("AudioX generated clip is missing or invalid.");
        }

        return path;
    }

    private async Task<EnhanceStemPaths> MixEnhanceTracksAsync(
        string narrationWavPath,
        IReadOnlyList<EnhanceTimelineEvent> timeline,
        IReadOnlyList<TimedTextUnit> units,
        string finalWavPath,
        string outputDir,
        string baseName,
        CancellationToken ct)
    {
        var stems = new EnhanceStemPaths();
        using var narrationReader = new AudioFileReader(narrationWavPath);
        var targetSampleRate = narrationReader.WaveFormat.SampleRate;
        var targetChannels = Math.Clamp(narrationReader.WaveFormat.Channels, 1, 2);
        var voice = ReadAudioAsFloats(narrationWavPath, targetSampleRate, targetChannels);

        var totalFrames = voice.Length / targetChannels;
        var ambience = new float[voice.Length];
        var sfx = new float[voice.Length];

        foreach (var ev in timeline.OrderBy(e => e.StartSeconds))
        {
            ct.ThrowIfCancellationRequested();
            var clip = ReadAudioAsFloats(ev.AudioPath, targetSampleRate, targetChannels);
            if (clip.Length == 0)
            {
                continue;
            }

            var startFrame = Math.Clamp((int)Math.Round(ev.StartSeconds * targetSampleRate), 0, totalFrames - 1);
            var endFrame = Math.Clamp((int)Math.Round(ev.EndSeconds * targetSampleRate), startFrame + 1, totalFrames);
            var gainDb = ev.CueType == "ambience" ? _config.AudioEnhanceAmbienceDb : _config.AudioEnhanceOneShotDb;
            var gain = (float)(Math.Pow(10.0, gainDb / 20.0) * Math.Clamp(ev.Intensity, 0.1, 1.0));

            if (ev.CueType == "ambience")
            {
                MixLoopedClip(
                    destination: ambience,
                    clip: clip,
                    channels: targetChannels,
                    startFrame: startFrame,
                    endFrame: endFrame,
                    gain: gain,
                    fadeInSeconds: 0.30,
                    fadeOutSeconds: 0.50,
                    sampleRate: targetSampleRate);
            }
            else
            {
                MixOneShotClip(
                    destination: sfx,
                    clip: clip,
                    channels: targetChannels,
                    startFrame: startFrame,
                    endFrame: endFrame,
                    gain: gain,
                    fadeInSeconds: 0.02,
                    fadeOutSeconds: 0.08,
                    sampleRate: targetSampleRate);
            }
        }

        ApplyDialogueDucking(units, ambience, targetChannels, targetSampleRate, _config.AudioEnhanceDuckingDb);

        var final = new float[voice.Length];
        var peak = 0f;
        for (var i = 0; i < final.Length; i++)
        {
            var v = voice[i] + ambience[i] + sfx[i];
            peak = Math.Max(peak, Math.Abs(v));
            final[i] = v;
        }

        if (peak > 0.98f)
        {
            var scale = 0.98f / peak;
            for (var i = 0; i < final.Length; i++)
            {
                final[i] *= scale;
            }
        }

        WritePcm16Wave(finalWavPath, final, targetSampleRate, targetChannels);

        if (_config.AudioEnhanceExportStems)
        {
            var stemsDir = Path.Combine(outputDir, $"{baseName}.stems");
            Directory.CreateDirectory(stemsDir);
            stems.VoicePath = Path.Combine(stemsDir, "voice.wav");
            stems.AmbiencePath = Path.Combine(stemsDir, "ambience.wav");
            stems.SfxPath = Path.Combine(stemsDir, "sfx.wav");
            WritePcm16Wave(stems.VoicePath, voice, targetSampleRate, targetChannels);
            WritePcm16Wave(stems.AmbiencePath, ambience, targetSampleRate, targetChannels);
            WritePcm16Wave(stems.SfxPath, sfx, targetSampleRate, targetChannels);
        }

        await Task.CompletedTask;
        return stems;
    }

    private static void ApplyDialogueDucking(
        IReadOnlyList<TimedTextUnit> units,
        float[] ambience,
        int channels,
        int sampleRate,
        double duckingDb)
    {
        if (ambience.Length == 0 || units.Count == 0)
        {
            return;
        }

        var totalFrames = ambience.Length / channels;
        var duckLinear = (float)Math.Pow(10.0, Math.Clamp(duckingDb, -30.0, 0.0) / 20.0);
        var mask = Enumerable.Repeat(1.0f, totalFrames).ToArray();

        foreach (var unit in units)
        {
            if (!LooksLikeDialogue(unit.Text))
            {
                continue;
            }

            var start = Math.Clamp((int)Math.Round(unit.StartSeconds * sampleRate), 0, totalFrames - 1);
            var end = Math.Clamp(start + (int)Math.Round(sampleRate * 1.2), start + 1, totalFrames);
            for (var f = start; f < end; f++)
            {
                mask[f] = Math.Min(mask[f], duckLinear);
            }
        }

        for (var f = 0; f < totalFrames; f++)
        {
            var g = mask[f];
            var offset = f * channels;
            for (var ch = 0; ch < channels; ch++)
            {
                ambience[offset + ch] *= g;
            }
        }
    }

    private static bool LooksLikeDialogue(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.Contains('"') || text.Contains('\'');
    }

    private static void MixLoopedClip(
        float[] destination,
        float[] clip,
        int channels,
        int startFrame,
        int endFrame,
        float gain,
        double fadeInSeconds,
        double fadeOutSeconds,
        int sampleRate)
    {
        var destFrames = destination.Length / channels;
        var clipFrames = Math.Max(1, clip.Length / channels);
        var clampedStart = Math.Clamp(startFrame, 0, destFrames - 1);
        var clampedEnd = Math.Clamp(endFrame, clampedStart + 1, destFrames);
        var totalFrames = clampedEnd - clampedStart;
        var fadeIn = Math.Clamp((int)Math.Round(fadeInSeconds * sampleRate), 0, totalFrames);
        var fadeOut = Math.Clamp((int)Math.Round(fadeOutSeconds * sampleRate), 0, totalFrames);

        for (var i = 0; i < totalFrames; i++)
        {
            var destFrame = clampedStart + i;
            var srcFrame = i % clipFrames;
            var envelope = 1.0f;
            if (fadeIn > 0 && i < fadeIn)
            {
                envelope = Math.Min(envelope, i / (float)fadeIn);
            }
            if (fadeOut > 0 && i > totalFrames - fadeOut)
            {
                envelope = Math.Min(envelope, (totalFrames - i) / (float)fadeOut);
            }

            var destOffset = destFrame * channels;
            var srcOffset = srcFrame * channels;
            for (var ch = 0; ch < channels; ch++)
            {
                destination[destOffset + ch] += clip[srcOffset + ch] * gain * envelope;
            }
        }
    }

    private static void MixOneShotClip(
        float[] destination,
        float[] clip,
        int channels,
        int startFrame,
        int endFrame,
        float gain,
        double fadeInSeconds,
        double fadeOutSeconds,
        int sampleRate)
    {
        var destFrames = destination.Length / channels;
        var clipFrames = Math.Max(1, clip.Length / channels);
        var clampedStart = Math.Clamp(startFrame, 0, destFrames - 1);
        var clampedEnd = Math.Clamp(endFrame, clampedStart + 1, destFrames);
        var totalFrames = Math.Min(clampedEnd - clampedStart, clipFrames);
        var fadeIn = Math.Clamp((int)Math.Round(fadeInSeconds * sampleRate), 0, totalFrames);
        var fadeOut = Math.Clamp((int)Math.Round(fadeOutSeconds * sampleRate), 0, totalFrames);

        for (var i = 0; i < totalFrames; i++)
        {
            var envelope = 1.0f;
            if (fadeIn > 0 && i < fadeIn)
            {
                envelope = Math.Min(envelope, i / (float)fadeIn);
            }
            if (fadeOut > 0 && i > totalFrames - fadeOut)
            {
                envelope = Math.Min(envelope, (totalFrames - i) / (float)fadeOut);
            }

            var destOffset = (clampedStart + i) * channels;
            var srcOffset = i * channels;
            for (var ch = 0; ch < channels; ch++)
            {
                destination[destOffset + ch] += clip[srcOffset + ch] * gain * envelope;
            }
        }
    }

    private static float[] ReadAudioAsFloats(string path, int targetSampleRate, int targetChannels)
    {
        using var reader = new AudioFileReader(path);
        ISampleProvider provider = reader;
        if (provider.WaveFormat.Channels != targetChannels)
        {
            provider = targetChannels == 1
                ? new StereoToMonoSampleProvider(provider)
                : new MonoToStereoSampleProvider(provider);
        }

        if (provider.WaveFormat.SampleRate != targetSampleRate)
        {
            provider = new WdlResamplingSampleProvider(provider, targetSampleRate);
        }

        var samples = new List<float>(targetSampleRate * targetChannels * 4);
        var buffer = new float[targetSampleRate * targetChannels];
        while (true)
        {
            var read = provider.Read(buffer, 0, buffer.Length);
            if (read <= 0)
            {
                break;
            }

            for (var i = 0; i < read; i++)
            {
                samples.Add(buffer[i]);
            }
        }

        return samples.ToArray();
    }

    private static void WritePcm16Wave(string path, float[] samples, int sampleRate, int channels)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? RuntimePaths.AppRoot);
        using var writer = new WaveFileWriter(path, new WaveFormat(sampleRate, 16, channels));
        var pcm = new byte[Math.Max(4096, samples.Length * 2)];
        var offset = 0;
        for (var i = 0; i < samples.Length; i++)
        {
            var clamped = Math.Clamp(samples[i], -1.0f, 1.0f);
            var s = (short)Math.Round(clamped * short.MaxValue);
            pcm[offset++] = (byte)(s & 0xFF);
            pcm[offset++] = (byte)((s >> 8) & 0xFF);
            if (offset >= pcm.Length - 2)
            {
                writer.Write(pcm, 0, offset);
                offset = 0;
            }
        }

        if (offset > 0)
        {
            writer.Write(pcm, 0, offset);
        }
    }

    private string ResolveAudioEnhancePython()
    {
        var configured = (_config.AudioEnhanceRuntimePath ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
        {
            return configured;
        }

        var embedded = Path.Combine(RuntimePaths.AppRoot, "python_qwen", "Scripts", "python.exe");
        if (File.Exists(embedded))
        {
            return embedded;
        }

        return "python";
    }

    private string ResolveAudioEnhanceModelDir()
    {
        var raw = string.IsNullOrWhiteSpace(_config.AudioEnhanceModelDir) ? "models/audiox" : _config.AudioEnhanceModelDir.Trim();
        return Path.IsPathRooted(raw) ? raw : Path.Combine(RuntimePaths.AppRoot, raw);
    }

    private string EnsureAudioXScript()
    {
        var scriptDir = Path.Combine(RuntimePaths.AppRoot, "tools", "audiox");
        Directory.CreateDirectory(scriptDir);
        var scriptPath = Path.Combine(scriptDir, "audiox_generate.py");
        var shouldWrite = true;
        if (File.Exists(scriptPath))
        {
            var existing = File.ReadAllText(scriptPath);
            shouldWrite = !string.Equals(existing, AudioXGenerateScript, StringComparison.Ordinal);
        }

        if (shouldWrite)
        {
            File.WriteAllText(scriptPath, AudioXGenerateScript, new UTF8Encoding(false));
        }
        return scriptPath;
    }

    private static string HashStable(string input)
    {
        var bytes = Encoding.UTF8.GetBytes(input ?? string.Empty);
        var hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant()[..16];
    }

    private static string TrimForError(string? text)
    {
        var s = (text ?? string.Empty).Trim();
        if (s.Length <= 240)
        {
            return s;
        }
        return s[..240] + "...";
    }

    private static string? TryReadString(string json, params object[] path)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            JsonElement current = doc.RootElement;
            foreach (var p in path)
            {
                if (p is string key)
                {
                    if (!current.TryGetProperty(key, out var child))
                    {
                        return null;
                    }
                    current = child;
                }
                else if (p is int idx)
                {
                    if (current.ValueKind != JsonValueKind.Array || idx < 0 || idx >= current.GetArrayLength())
                    {
                        return null;
                    }
                    current = current[idx];
                }
            }

            return current.ValueKind == JsonValueKind.String ? current.GetString() : current.ToString();
        }
        catch
        {
            return null;
        }
    }

    private const string AudioXGenerateScript = """
import argparse
import os
import sys

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model-dir", required=False, default="")
    parser.add_argument("--variant", required=False, default="audiox_base")
    parser.add_argument("--prompt", required=True)
    parser.add_argument("--seconds", required=True, type=float)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    model_dir = args.model_dir or ""
    if model_dir:
        os.makedirs(model_dir, exist_ok=True)
        os.environ.setdefault("HF_HOME", model_dir)
        os.environ.setdefault("HUGGINGFACE_HUB_CACHE", model_dir)
        os.environ.setdefault("TRANSFORMERS_CACHE", model_dir)

    try:
        import torch
        import torchaudio
    except Exception as ex:
        print(f"ERR: missing torch/torchaudio: {ex}", file=sys.stderr)
        return 2

    model_loader = None
    generate_fn = None
    try:
        from audiox import get_pretrained_model
        from audiox.inference.generation import generate_diffusion_cond
        model_loader = get_pretrained_model
        generate_fn = generate_diffusion_cond
    except Exception:
        try:
            from stable_audio_tools import get_pretrained_model
            from stable_audio_tools.inference.generation import generate_diffusion_cond
            model_loader = get_pretrained_model
            generate_fn = generate_diffusion_cond
        except Exception as ex:
            print(f"ERR: AudioX runtime import failed: {ex}", file=sys.stderr)
            return 3

    model_name = {
        "audiox_base": "HKUSTAudio/AudioX",
        "audiox_maf": "HKUSTAudio/AudioX-MAF",
        "audiox_maf_mmdit": "HKUSTAudio/AudioX-MAF-MMDiT"
    }.get((args.variant or "audiox_base").lower(), "HKUSTAudio/AudioX")

    try:
        model, model_config = model_loader(model_name)
    except Exception as ex:
        print(f"ERR: failed loading model '{model_name}': {ex}", file=sys.stderr)
        return 4

    model.eval()
    if torch.cuda.is_available():
        model = model.to("cuda")
    else:
        model = model.to("cpu")

    requested_seconds = max(1.0, min(8.0, float(args.seconds)))
    model_seconds = 10.0
    sample_rate = int(getattr(model, "sample_rate", 44100))
    target_fps = 5
    if isinstance(_model_config, dict):
        try:
            target_fps = int(_model_config.get("video_fps", target_fps))
        except Exception:
            target_fps = 5

    video_tensors = torch.zeros((int(target_fps * model_seconds), 3, 224, 224), dtype=torch.float32, device=device)
    video_sync_frames = torch.zeros((1, 240, 768), dtype=torch.float32, device=device)
    audio_prompt = torch.zeros((1, 2, int(sample_rate * model_seconds)), dtype=torch.float32, device=device)
    conditioning = [{
        "video_prompt": {
            "video_tensors": video_tensors.unsqueeze(0),
            "video_sync_frames": video_sync_frames
        },
        "text_prompt": args.prompt or "",
        "audio_prompt": audio_prompt,
        "seconds_start": 0,
        "seconds_total": model_seconds
    }]

    try:
        output = generate_fn(
            model,
            steps=40,
            cfg_scale=7,
            conditioning=conditioning,
            sample_size=int(sample_rate * model_seconds),
            sigma_min=0.3,
            sigma_max=500,
            sampler_type="dpmpp-3m-sde",
            device=device
        )
    except Exception as ex:
        print(f"ERR: generation failed: {ex}", file=sys.stderr)
        return 5

    if output is None:
        print("ERR: generation returned no output", file=sys.stderr)
        return 6

    if not os.path.isdir(os.path.dirname(os.path.abspath(args.output))):
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)

    try:
        audio = output[0].to(torch.float32).detach().cpu()
        target_len = int(sample_rate * requested_seconds)
        if audio.shape[-1] > target_len:
            audio = audio[..., :target_len]
        elif audio.shape[-1] < target_len:
            pad = target_len - audio.shape[-1]
            audio = torch.nn.functional.pad(audio, (0, pad))
        if audio.dim() == 1:
            audio = audio.unsqueeze(0)
        torchaudio.save(args.output, audio, sample_rate)
    except Exception as ex:
        print(f"ERR: failed writing wav: {ex}", file=sys.stderr)
        return 7

    print("OK")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
""";
}
