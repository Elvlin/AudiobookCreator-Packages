using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text;

namespace App.Inference;

public sealed class ApiBackend : ITtsBackend
{
    public string Name => "api";

    private readonly InferenceOptions _options;
    private readonly HttpClient _httpClient;
    private string? _resolvedAlibabaVcVoiceId;
    private string? _resolvedAlibabaVcVoiceModel;
    private static readonly object AlibabaVcCacheLock = new();
    private static readonly HashSet<string> AlibabaSystemVoices = new(StringComparer.OrdinalIgnoreCase)
    {
        // Common system voices documented for Qwen TTS families.
        "Cherry", "Serena", "Ethan", "Chelsie"
    };

    public ApiBackend(InferenceOptions options, HttpClient? httpClient = null)
    {
        _options = options;
        _httpClient = httpClient ?? new HttpClient();
    }

    public async Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(_options.ApiKey))
        {
            throw new InvalidOperationException("API key is required for API backend.");
        }
        if (string.IsNullOrWhiteSpace(request.OutputPath))
        {
            throw new ArgumentException("OutputPath is required.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(request.OutputPath)!);

        var provider = (_options.Provider ?? "openai").Trim().ToLowerInvariant();
        if (provider == "openai")
        {
            await SynthesizeWithOpenAiAsync(request, ct);
            return;
        }

        if (provider == "alibaba")
        {
            await SynthesizeWithAlibabaAsync(request, ct);
            return;
        }

        throw new InvalidOperationException($"Unsupported API provider: {provider}. Supported providers: openai, alibaba.");
    }

    public async Task GenerateAlibabaVoiceDesignAsync(
        string voiceName,
        string previewText,
        string stylePrompt,
        string languageType,
        string outputPath,
        IProgress<QwenWorkerProgress>? progress = null,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(_options.ApiKey))
        {
            throw new InvalidOperationException("API key is required for Alibaba VoiceDesign.");
        }

        if (string.IsNullOrWhiteSpace(previewText))
        {
            throw new InvalidOperationException("Voice text is required for Alibaba VoiceDesign.");
        }

        if (string.IsNullOrWhiteSpace(stylePrompt))
        {
            throw new InvalidOperationException("Style prompt is required for Alibaba VoiceDesign.");
        }

        var baseUrl = ResolveBaseUrl();
        var targetModel = string.IsNullOrWhiteSpace(_options.VoiceDesignTargetModel)
            ? "qwen3-tts-vd-2026-01-26"
            : _options.VoiceDesignTargetModel.Trim();
        var languageCode = MapAlibabaVoiceDesignLanguageCode(languageType);

        progress?.Report(new QwenWorkerProgress("Preparing model...", null, 1, 4));

        var input = new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["target_model"] = targetModel,
            ["voice_prompt"] = stylePrompt,
            ["preview_text"] = previewText,
            ["preferred_name"] = string.IsNullOrWhiteSpace(voiceName) ? "voice_design" : voiceName.Trim()
        };
        if (!string.IsNullOrWhiteSpace(languageCode))
        {
            input["language"] = languageCode;
        }

        using var createReq = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/services/audio/tts/customization");
        createReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.ApiKey);
        createReq.Content = JsonContent.Create(new
        {
            model = "qwen-voice-design",
            input,
            parameters = new
            {
                sample_rate = 24000,
                response_format = "wav"
            }
        });

        using var createRes = await _httpClient.SendAsync(createReq, HttpCompletionOption.ResponseHeadersRead, ct);
        var createPayload = await createRes.Content.ReadAsStringAsync(ct);
        if (!createRes.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Alibaba VoiceDesign API error: {(int)createRes.StatusCode} {createRes.ReasonPhrase}");
        }

        var designedVoice = TryReadString(createPayload, "output", "voice");
        if (string.IsNullOrWhiteSpace(designedVoice))
        {
            throw new InvalidOperationException("Alibaba VoiceDesign did not return generated voice ID.");
        }

        progress?.Report(new QwenWorkerProgress("Generating voice audio...", null, 3, 4));

        var request = new TtsRequest
        {
            Text = previewText,
            OutputPath = outputPath
        };

        await SynthesizeWithAlibabaAsync(
            request,
            ct,
            overrideModel: targetModel,
            overrideVoice: designedVoice,
            overrideLanguageType: languageType);

        progress?.Report(new QwenWorkerProgress("Finalizing output...", null, 4, 4));
    }

    private async Task SynthesizeWithOpenAiAsync(TtsRequest request, CancellationToken ct)
    {
        var model = string.IsNullOrWhiteSpace(_options.ModelId) ? "gpt-4o-mini-tts" : _options.ModelId;
        var voice = string.IsNullOrWhiteSpace(_options.Voice) ? "alloy" : _options.Voice;
        var text = (request.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new InvalidOperationException("Input text is empty.");
        }
        var parts = SplitApiInputText(text, 3800);
        if (parts.Count <= 1)
        {
            var singleWav = await FetchOpenAiWavChunkAsync(parts[0].Text, model, voice, request.Speed, ct);
            await File.WriteAllBytesAsync(request.OutputPath, singleWav, ct);
            return;
        }

        var chunks = new List<byte[]>(parts.Count);
        var pausesMs = new List<int>(Math.Max(0, parts.Count - 1));
        for (var i = 0; i < parts.Count; i++)
        {
            var wav = await FetchOpenAiWavChunkAsync(parts[i].Text, model, voice, request.Speed, ct);
            chunks.Add(wav);
            if (i + 1 < parts.Count)
            {
                pausesMs.Add(ResolveApiBoundaryPauseMs(parts[i], request));
            }
        }

        var merged = MergeWaveAudio(chunks, pausesMs, "OpenAI");
        await File.WriteAllBytesAsync(request.OutputPath, merged, ct);
    }

    private async Task SynthesizeWithAlibabaAsync(
        TtsRequest request,
        CancellationToken ct,
        string? overrideModel = null,
        string? overrideVoice = null,
        string? overrideLanguageType = null)
    {
        var text = (request.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new InvalidOperationException("Input text is empty.");
        }

        var model = string.IsNullOrWhiteSpace(overrideModel)
            ? (string.IsNullOrWhiteSpace(_options.ModelId) ? "qwen3-tts-flash" : _options.ModelId.Trim())
            : overrideModel.Trim();
        var configuredVoice = string.IsNullOrWhiteSpace(overrideVoice) ? _options.Voice : overrideVoice;
        var requiresClonedVoice = model.Contains("qwen3-tts-vc", StringComparison.OrdinalIgnoreCase);
        var configuredVoiceLooksLikeSystemVoice = AlibabaSystemVoices.Contains((configuredVoice ?? string.Empty).Trim());
        if (requiresClonedVoice &&
            (string.IsNullOrWhiteSpace(configuredVoice) || LooksLikeLocalAudioReference(configuredVoice) || configuredVoiceLooksLikeSystemVoice))
        {
            configuredVoice = await ResolveOrCreateAlibabaVcVoiceAsync(request.VoicePath, model, ct);
        }
        if (requiresClonedVoice && string.IsNullOrWhiteSpace(configuredVoice))
        {
            throw new InvalidOperationException("Alibaba Qwen3-TTS-VC requires a cloned voice ID.");
        }
        var voice = requiresClonedVoice
            ? configuredVoice!.Trim()
            : NormalizeAlibabaSystemVoice(configuredVoice);

        var languageType = NormalizeAlibabaLanguageType(
            string.IsNullOrWhiteSpace(overrideLanguageType) ? _options.LanguageType : overrideLanguageType);
        var parts = SplitApiInputText(text, 560);
        if (parts.Count <= 1)
        {
            var singleAudio = await FetchAlibabaAudioBytesAsync(parts[0].Text, model, voice, languageType, ct);
            await File.WriteAllBytesAsync(request.OutputPath, singleAudio, ct);
            return;
        }

        var chunkAudio = new List<byte[]>(parts.Count);
        var pausesMs = new List<int>(Math.Max(0, parts.Count - 1));
        for (var i = 0; i < parts.Count; i++)
        {
            var bytes = await FetchAlibabaAudioBytesAsync(parts[i].Text, model, voice, languageType, ct);
            chunkAudio.Add(bytes);
            if (i + 1 < parts.Count)
            {
                pausesMs.Add(ResolveApiBoundaryPauseMs(parts[i], request));
            }
        }

        var merged = MergeWaveAudio(chunkAudio, pausesMs, "Alibaba");
        await File.WriteAllBytesAsync(request.OutputPath, merged, ct);
    }

    private async Task<byte[]> FetchAlibabaAudioBytesAsync(
        string inputText,
        string model,
        string voice,
        string languageType,
        CancellationToken ct)
    {
        var baseUrl = ResolveBaseUrl();
        using var req = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/services/aigc/multimodal-generation/generation");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.ApiKey);
        req.Content = JsonContent.Create(new
        {
            model,
            input = new
            {
                text = inputText,
                voice,
                language_type = languageType
            }
        });

        using var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        var payload = await res.Content.ReadAsStringAsync(ct);
        if (!res.IsSuccessStatusCode)
        {
            var errorSnippet = string.IsNullOrWhiteSpace(payload)
                ? string.Empty
                : $" | {TrimForError(payload)}";
            throw new InvalidOperationException($"Alibaba API error: {(int)res.StatusCode} {res.ReasonPhrase}{errorSnippet}");
        }

        var audioUrl = TryReadString(payload, "output", "audio", "url");
        if (!string.IsNullOrWhiteSpace(audioUrl))
        {
            using var audioRes = await _httpClient.GetAsync(audioUrl, HttpCompletionOption.ResponseHeadersRead, ct);
            audioRes.EnsureSuccessStatusCode();
            var bytes = await audioRes.Content.ReadAsByteArrayAsync(ct);
            return bytes;
        }

        var audioData = TryReadString(payload, "output", "audio", "data");
        if (!string.IsNullOrWhiteSpace(audioData))
        {
            try
            {
                var bytes = Convert.FromBase64String(audioData);
                return bytes;
            }
            catch (FormatException)
            {
                throw new InvalidOperationException("Alibaba API returned invalid audio payload.");
            }
        }

        throw new InvalidOperationException("Alibaba API did not return audio URL/data.");
    }

    private static List<ApiTextPart> SplitApiInputText(string text, int maxChars)
    {
        var clean = (text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(clean))
        {
            return new List<ApiTextPart>();
        }

        var result = new List<ApiTextPart>();
        var paragraphs = Regex.Split(clean.Replace("\r\n", "\n"), @"\n\s*\n")
            .Select(p => Regex.Replace(p, @"\s+", " ").Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToList();
        if (paragraphs.Count == 0)
        {
            paragraphs.Add(Regex.Replace(clean, @"\s+", " ").Trim());
        }

        for (var p = 0; p < paragraphs.Count; p++)
        {
            var paragraph = paragraphs[p];
            var cursor = 0;
            while (cursor < paragraph.Length)
            {
                var remaining = paragraph.Length - cursor;
                if (remaining <= maxChars)
                {
                    var tail = paragraph[cursor..].Trim();
                    if (!string.IsNullOrWhiteSpace(tail))
                    {
                        result.Add(BuildApiTextPart(tail, paragraphEnd: true));
                    }
                    break;
                }

                var window = paragraph.Substring(cursor, maxChars);
                var cut = window.LastIndexOfAny(new[] { '.', '!', '?', ';', ':', ',', ' ' });
                if (cut < maxChars / 2)
                {
                    cut = maxChars;
                }

                var piece = paragraph.Substring(cursor, cut).Trim();
                if (!string.IsNullOrWhiteSpace(piece))
                {
                    result.Add(BuildApiTextPart(piece, paragraphEnd: false));
                }
                cursor += cut;
                while (cursor < paragraph.Length && char.IsWhiteSpace(paragraph[cursor]))
                {
                    cursor++;
                }
            }
        }

        return result;
    }

    private static byte[] MergeWaveAudio(IReadOnlyList<byte[]> waveFiles, IReadOnlyList<int>? pauseMs = null, string? providerName = null)
    {
        var provider = string.IsNullOrWhiteSpace(providerName) ? "API" : providerName.Trim();
        if (waveFiles.Count == 0)
        {
            throw new InvalidOperationException("No audio chunks to merge.");
        }

        if (waveFiles.Count == 1)
        {
            return waveFiles[0];
        }

        var parts = waveFiles.Select(bytes => ParseWave(bytes, provider)).ToList();
        var first = parts[0];
        foreach (var p in parts.Skip(1))
        {
            if (p.AudioFormat != first.AudioFormat ||
                p.Channels != first.Channels ||
                p.SampleRate != first.SampleRate ||
                p.BitsPerSample != first.BitsPerSample)
            {
                throw new InvalidOperationException($"{provider} API returned incompatible WAV chunks.");
            }
        }

        var totalData = parts.Sum(p => p.Data.Length);
        if (pauseMs is not null && pauseMs.Count > 0)
        {
            var bytesPerSecond = first.SampleRate * first.Channels * Math.Max(1, first.BitsPerSample / 8);
            for (var i = 0; i < Math.Min(pauseMs.Count, parts.Count - 1); i++)
            {
                var silenceBytes = (int)Math.Round(bytesPerSecond * (Math.Max(0, pauseMs[i]) / 1000.0));
                totalData += silenceBytes;
            }
        }
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms, Encoding.ASCII, leaveOpen: true);
        bw.Write(Encoding.ASCII.GetBytes("RIFF"));
        bw.Write(36 + totalData);
        bw.Write(Encoding.ASCII.GetBytes("WAVE"));
        bw.Write(Encoding.ASCII.GetBytes("fmt "));
        bw.Write(16);
        bw.Write(first.AudioFormat);
        bw.Write(first.Channels);
        bw.Write(first.SampleRate);
        var byteRate = first.SampleRate * first.Channels * (first.BitsPerSample / 8);
        bw.Write(byteRate);
        var blockAlign = (short)(first.Channels * (first.BitsPerSample / 8));
        bw.Write(blockAlign);
        bw.Write(first.BitsPerSample);
        bw.Write(Encoding.ASCII.GetBytes("data"));
        bw.Write(totalData);
        for (var i = 0; i < parts.Count; i++)
        {
            bw.Write(parts[i].Data);
            if (pauseMs is not null && i < parts.Count - 1 && i < pauseMs.Count)
            {
                var bytesPerSecond = first.SampleRate * first.Channels * Math.Max(1, first.BitsPerSample / 8);
                var silenceBytes = (int)Math.Round(bytesPerSecond * (Math.Max(0, pauseMs[i]) / 1000.0));
                if (silenceBytes > 0)
                {
                    bw.Write(new byte[silenceBytes]);
                }
            }
        }
        bw.Flush();
        return ms.ToArray();
    }

    private static WavePart ParseWave(byte[] bytes, string providerName)
    {
        using var ms = new MemoryStream(bytes);
        using var br = new BinaryReader(ms, Encoding.ASCII, leaveOpen: true);
        if (ms.Length < 44)
        {
            throw new InvalidOperationException($"{providerName} API returned invalid WAV (too short).");
        }

        var riff = Encoding.ASCII.GetString(br.ReadBytes(4));
        br.ReadInt32();
        var wave = Encoding.ASCII.GetString(br.ReadBytes(4));
        if (!string.Equals(riff, "RIFF", StringComparison.Ordinal) ||
            !string.Equals(wave, "WAVE", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"{providerName} API returned non-WAV audio.");
        }

        ushort audioFormat = 1;
        ushort channels = 1;
        int sampleRate = 24000;
        ushort bitsPerSample = 16;
        byte[]? data = null;
        while (ms.Position + 8 <= ms.Length)
        {
            var chunkId = Encoding.ASCII.GetString(br.ReadBytes(4));
            var chunkSizeRaw = br.ReadUInt32();
            long chunkSize = chunkSizeRaw;

            // Handle unknown/placeholder chunk lengths from some WAV emitters.
            if (chunkSizeRaw == uint.MaxValue)
            {
                chunkSize = Math.Max(0, ms.Length - ms.Position);
            }

            if (ms.Position + chunkSize > ms.Length)
            {
                if (chunkId == "data")
                {
                    // Clamp trailing data chunk to available bytes.
                    chunkSize = Math.Max(0, ms.Length - ms.Position);
                }
                else
                {
                    throw new InvalidOperationException($"{providerName} API returned invalid WAV chunk ({chunkId}).");
                }
            }

            if (chunkId == "fmt ")
            {
                audioFormat = br.ReadUInt16();
                channels = br.ReadUInt16();
                sampleRate = br.ReadInt32();
                br.ReadInt32();
                br.ReadUInt16();
                bitsPerSample = br.ReadUInt16();
                var left = (int)Math.Max(0, chunkSize - 16);
                if (left > 0)
                {
                    br.ReadBytes(left);
                }
            }
            else if (chunkId == "data")
            {
                data = br.ReadBytes((int)chunkSize);
            }
            else
            {
                br.ReadBytes((int)chunkSize);
            }

            if ((chunkSize & 1) == 1 && ms.Position < ms.Length)
            {
                br.ReadByte();
            }
        }

        if (data is null || data.Length == 0)
        {
            throw new InvalidOperationException($"{providerName} API WAV payload missing data chunk.");
        }

        return new WavePart(audioFormat, channels, sampleRate, bitsPerSample, data);
    }

    private sealed record WavePart(ushort AudioFormat, ushort Channels, int SampleRate, ushort BitsPerSample, byte[] Data);

    private async Task<string> ResolveOrCreateAlibabaVcVoiceAsync(string voicePath, string targetModel, CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(_resolvedAlibabaVcVoiceId) &&
            string.Equals(_resolvedAlibabaVcVoiceModel, targetModel, StringComparison.OrdinalIgnoreCase))
        {
            return _resolvedAlibabaVcVoiceId;
        }

        if (string.IsNullOrWhiteSpace(voicePath) || !File.Exists(voicePath))
        {
            throw new InvalidOperationException(
                "Qwen3-TTS-VC requires a local voice sample file. Select a Narrator Voice first, then generate.");
        }

        var fullPath = Path.GetFullPath(voicePath);
        var cacheKey = BuildAlibabaVcCacheKey(fullPath, targetModel);
        var persistent = LoadAlibabaVcCache();
        if (persistent.TryGetValue(cacheKey, out var cachedVoice) && !string.IsNullOrWhiteSpace(cachedVoice))
        {
            _resolvedAlibabaVcVoiceId = cachedVoice;
            _resolvedAlibabaVcVoiceModel = targetModel;
            return cachedVoice;
        }

        var created = await CreateAlibabaVcVoiceAsync(fullPath, targetModel, ct);
        _resolvedAlibabaVcVoiceId = created;
        _resolvedAlibabaVcVoiceModel = targetModel;
        persistent[cacheKey] = created;
        SaveAlibabaVcCache(persistent);
        return created;
    }

    private async Task<string> CreateAlibabaVcVoiceAsync(string samplePath, string targetModel, CancellationToken ct)
    {
        var dataUri = BuildAudioDataUri(samplePath);
        var preferredName = BuildAlibabaPreferredName(samplePath);
        var baseUrl = ResolveBaseUrl();
        using var req = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/services/audio/tts/customization");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.ApiKey);
        req.Content = JsonContent.Create(new
        {
            model = "qwen-voice-enrollment",
            input = new
            {
                action = "create",
                target_model = targetModel,
                preferred_name = preferredName,
                audio = new
                {
                    data = dataUri
                }
            }
        });

        using var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        var payload = await res.Content.ReadAsStringAsync(ct);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Alibaba voice enrollment error: {(int)res.StatusCode} {res.ReasonPhrase}");
        }

        var voiceId = TryReadString(payload, "output", "voice")
                      ?? TryReadString(payload, "output", "voice_id");
        if (string.IsNullOrWhiteSpace(voiceId))
        {
            throw new InvalidOperationException("Alibaba voice enrollment did not return voice ID.");
        }

        return voiceId.Trim();
    }

    private static string BuildAlibabaPreferredName(string path)
    {
        var stem = Path.GetFileNameWithoutExtension(path);
        if (string.IsNullOrWhiteSpace(stem))
        {
            stem = "voice";
        }

        var safe = Regex.Replace(stem, @"[^A-Za-z0-9_]", "_");
        if (safe.Length > 16)
        {
            safe = safe[..16];
        }
        if (string.IsNullOrWhiteSpace(safe))
        {
            safe = "voice";
        }
        return safe;
    }

    private static string BuildAudioDataUri(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var mime = ext switch
        {
            ".wav" => "audio/wav",
            ".mp3" => "audio/mpeg",
            ".m4a" => "audio/mp4",
            ".flac" => "audio/flac",
            _ => "application/octet-stream"
        };
        var bytes = File.ReadAllBytes(filePath);
        var b64 = Convert.ToBase64String(bytes);
        return $"data:{mime};base64,{b64}";
    }

    private static string BuildAlibabaVcCacheKey(string fullPath, string targetModel)
    {
        var fi = new FileInfo(fullPath);
        var stamp = $"{fi.Length}:{fi.LastWriteTimeUtc.Ticks}";
        return $"{fullPath.ToLowerInvariant()}|{targetModel.ToLowerInvariant()}|{stamp}";
    }

    private static string GetAlibabaVcCachePath()
    {
        var root = AppContext.BaseDirectory;
        return Path.Combine(root, "alibaba_vc_voice_cache.json");
    }

    private static Dictionary<string, string> LoadAlibabaVcCache()
    {
        lock (AlibabaVcCacheLock)
        {
            var path = GetAlibabaVcCachePath();
            if (!File.Exists(path))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            try
            {
                var json = File.ReadAllText(path);
                var parsed = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                return parsed is null
                    ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, string>(parsed, StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }
    }

    private static void SaveAlibabaVcCache(Dictionary<string, string> cache)
    {
        lock (AlibabaVcCacheLock)
        {
            var path = GetAlibabaVcCachePath();
            var json = JsonSerializer.Serialize(cache);
            File.WriteAllText(path, json);
        }
    }

    private string ResolveBaseUrl()
    {
        var baseUrl = string.IsNullOrWhiteSpace(_options.BaseUrl)
            ? "https://dashscope-intl.aliyuncs.com/api/v1"
            : _options.BaseUrl.Trim().TrimEnd('/');
        return baseUrl;
    }

    private static string NormalizeAlibabaLanguageType(string? value)
    {
        return (value ?? "Auto").Trim().ToLowerInvariant() switch
        {
            "auto" => "Auto",
            "english" => "English",
            "chinese" => "Chinese",
            "japanese" => "Japanese",
            "korean" => "Korean",
            "german" => "German",
            "french" => "French",
            "russian" => "Russian",
            "portuguese" => "Portuguese",
            "spanish" => "Spanish",
            "italian" => "Italian",
            _ => "Auto"
        };
    }

    private static string? MapAlibabaVoiceDesignLanguageCode(string? value)
    {
        return (value ?? "Auto").Trim().ToLowerInvariant() switch
        {
            "auto" => null,
            "english" => "en",
            "chinese" => "zh",
            "japanese" => "ja",
            "korean" => "ko",
            "german" => "de",
            "french" => "fr",
            "russian" => "ru",
            "portuguese" => "pt",
            "spanish" => "es",
            "italian" => "it",
            _ => null
        };
    }

    private static string NormalizeAlibabaSystemVoice(string? value)
    {
        var voice = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(voice))
        {
            return "Cherry";
        }

        // Local files and local clone names should never be sent to Alibaba TTS as voice IDs.
        if (LooksLikeLocalAudioReference(voice))
        {
            return "Cherry";
        }

        // Alibaba custom voices are returned as IDs; keep those.
        if (voice.StartsWith("qwen-", StringComparison.OrdinalIgnoreCase))
        {
            return voice;
        }

        // Keep documented system voices; fallback otherwise.
        return AlibabaSystemVoices.Contains(voice) ? voice : "Cherry";
    }

    private static bool LooksLikeLocalAudioReference(string? value)
    {
        var s = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(s))
        {
            return false;
        }

        if (s.Contains('\\') || s.Contains('/') || s.Contains(':'))
        {
            return true;
        }

        var ext = Path.GetExtension(s);
        if (string.IsNullOrWhiteSpace(ext))
        {
            return false;
        }

        return ext.Equals(".wav", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".mp3", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".m4a", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".flac", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".ogg", StringComparison.OrdinalIgnoreCase);
    }

    private static string TrimForError(string payload)
    {
        const int maxLen = 400;
        var text = payload.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return text.Length <= maxLen ? text : text[..maxLen] + "...";
    }

    private async Task<byte[]> FetchOpenAiWavChunkAsync(
        string inputText,
        string model,
        string voice,
        float speed,
        CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/audio/speech");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.ApiKey);
        req.Content = JsonContent.Create(new
        {
            model,
            voice,
            input = inputText,
            response_format = "wav",
            speed = Math.Clamp(speed <= 0 ? 1.0f : speed, 0.25f, 4.0f)
        });

        using var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        var bytes = await res.Content.ReadAsByteArrayAsync(ct);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"OpenAI API error: {(int)res.StatusCode} {res.ReasonPhrase}");
        }

        var mediaType = res.Content.Headers.ContentType?.MediaType ?? string.Empty;
        if (mediaType.Contains("json", StringComparison.OrdinalIgnoreCase))
        {
            var payload = Encoding.UTF8.GetString(bytes);
            throw new InvalidOperationException($"OpenAI API returned JSON instead of audio: {TrimForError(payload)}");
        }
        return bytes;
    }

    private static ApiTextPart BuildApiTextPart(string text, bool paragraphEnd)
    {
        var t = (text ?? string.Empty).Trim();
        var last = t.Length == 0 ? '\0' : t[^1];
        var ellipsis = t.EndsWith("...", StringComparison.Ordinal) || t.EndsWith("…", StringComparison.Ordinal);
        return new ApiTextPart(t, last, paragraphEnd, ellipsis);
    }

    private static int ResolveApiBoundaryPauseMs(ApiTextPart part, TtsRequest request)
    {
        var chunkPause = Math.Max(0, request.ChunkPauseMs);
        var paragraphPause = Math.Max(0, request.ParagraphPauseMs);
        var clause = ResolvePauseRange(request.ClausePauseMinMs, request.ClausePauseMaxMs, chunkPause);
        var sentence = ResolvePauseRange(request.SentencePauseMinMs, request.SentencePauseMaxMs, chunkPause);
        var ellipsis = ResolvePauseRange(request.EllipsisPauseMinMs, request.EllipsisPauseMaxMs, chunkPause);
        var paragraph = ResolvePauseRange(request.ParagraphPauseMinMs, request.ParagraphPauseMaxMs, paragraphPause);

        if (part.ParagraphEnd)
        {
            return Math.Max(paragraphPause, Mid(paragraph.Min, paragraph.Max));
        }
        if (part.EndsWithEllipsis)
        {
            return Math.Max(chunkPause, Mid(ellipsis.Min, ellipsis.Max));
        }

        return part.LastChar switch
        {
            '.' or '!' or '?' => Math.Max(chunkPause, Mid(sentence.Min, sentence.Max)),
            ',' or ';' or ':' => Math.Max(chunkPause, Mid(clause.Min, clause.Max)),
            _ => chunkPause
        };
    }

    private static (int Min, int Max) ResolvePauseRange(int configuredMin, int configuredMax, int fallback)
    {
        var min = configuredMin;
        var max = configuredMax;
        if (min <= 0 && max <= 0)
        {
            min = fallback;
            max = fallback;
        }
        if (min > max)
        {
            (min, max) = (max, min);
        }
        min = Math.Max(0, min);
        max = Math.Max(0, max);
        return (min, max);
    }

    private static int Mid(int min, int max) => min + ((max - min) / 2);

    private sealed record ApiTextPart(string Text, char LastChar, bool ParagraphEnd, bool EndsWithEllipsis);

    private static string? TryReadString(string json, params string[] path)
    {
        if (string.IsNullOrWhiteSpace(json) || path.Length == 0)
        {
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(json);
            JsonElement cursor = doc.RootElement;
            foreach (var segment in path)
            {
                if (cursor.ValueKind != JsonValueKind.Object || !cursor.TryGetProperty(segment, out cursor))
                {
                    return null;
                }
            }

            return cursor.ValueKind == JsonValueKind.String ? cursor.GetString() : null;
        }
        catch
        {
            return null;
        }
    }
}
