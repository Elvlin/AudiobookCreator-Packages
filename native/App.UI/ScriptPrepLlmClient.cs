using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Net.Sockets;
using App.Core.Models;
using App.Core.Runtime;

namespace AudiobookCreator.UI;

internal sealed class ScriptPrepLlmClient
{
    public sealed class ScriptPrepProgress
    {
        public int Completed { get; init; }
        public int Total { get; init; }
        public string Stage { get; init; } = string.Empty;
    }

    private readonly AppConfig _cfg;
    private readonly HttpClient _http = new();
    public bool LastSplitUsedFallback { get; private set; }
    public string LastSplitFallbackReason { get; private set; } = string.Empty;

    public ScriptPrepLlmClient(AppConfig cfg)
    {
        _cfg = cfg;
    }

    public bool IsLocalServerRunning => LocalLlmServerManager.IsRunning;

    public Task StartLocalServerAsync(CancellationToken ct)
    {
        var splitModel = ResolveLocalModelFilePath(isInstruction: false);
        return LocalLlmServerManager.EnsureRunningForModelAsync(_cfg, splitModel, ct);
    }

    public Task StopLocalServerAsync()
    {
        return LocalLlmServerManager.StopAsync();
    }

    public static void ForceStopAllLocalServers()
    {
        LocalLlmServerManager.ForceStopSync();
    }

    public async Task<List<PreparedScriptPart>> SplitAsync(string sourceText, CancellationToken ct, IProgress<ScriptPrepProgress>? progress = null)
    {
        LastSplitUsedFallback = false;
        LastSplitFallbackReason = string.Empty;
        var model = ResolveModelForTask(isInstruction: false);
        var temp = _cfg.LlmPrepTemperatureSplit;
        var maxTokens = _cfg.LlmPrepMaxTokensSplit;
        var provider = (_cfg.LlmPrepProvider ?? "local").Trim().ToLowerInvariant();

        // Local runtime is more context-limited; split the source first and merge parts.
        if (provider == "local")
        {
            var chunks = BuildLocalSegmentationChunks(sourceText, maxTokens);
            var merged = new List<PreparedScriptPart>();
            var nextOrder = 0;
            for (var i = 0; i < chunks.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var chunkText = chunks[i];
                progress?.Report(new ScriptPrepProgress
                {
                    Completed = i,
                    Total = chunks.Count,
                    Stage = $"Splitting chunk {i + 1}/{chunks.Count}..."
                });

                var prompt = BuildSegmentationPrompt(chunkText);
                var content = await RunPromptAsync(model, temp, maxTokens, prompt, "segmentation", expectJson: true, ct);
                var partList = ParseSegmentationWithFallback(chunkText, content);
                for (var p = 0; p < partList.Count; p++)
                {
                    var part = partList[p];
                    part.Order = nextOrder++;
                    if (string.IsNullOrWhiteSpace(part.Id))
                    {
                        part.Id = Guid.NewGuid().ToString("N");
                    }
                    merged.Add(part);
                }
            }

            progress?.Report(new ScriptPrepProgress
            {
                Completed = chunks.Count,
                Total = chunks.Count,
                Stage = "Parsing split result..."
            });
            return merged;
        }

        // Cloud providers process whole chapter in one request by design.
        progress?.Report(new ScriptPrepProgress { Completed = 0, Total = 1, Stage = "Building split request..." });
        var promptAll = BuildSegmentationPrompt(sourceText);
        progress?.Report(new ScriptPrepProgress { Completed = 0, Total = 1, Stage = "Running split model..." });
        var contentAll = await RunPromptAsync(model, temp, maxTokens, promptAll, "segmentation", expectJson: true, ct);
        progress?.Report(new ScriptPrepProgress { Completed = 1, Total = 1, Stage = "Parsing split result..." });
        return ParseSegmentationWithFallback(sourceText, contentAll);
    }

    private List<PreparedScriptPart> ParseSegmentationWithFallback(string sourceText, string content)
    {
        try
        {
            return ParseSegmentationResult(content);
        }
        catch (Exception ex)
        {
            // Fall back to deterministic local splitting when segmentation output is invalid.
            LastSplitUsedFallback = true;
            LastSplitFallbackReason = $"LLM split response was invalid or unparsable: {TrimForError(ex.Message)}";
            var fallback = ScriptPrepEngine.ReSplit(sourceText);
            if (fallback.Count > 0)
            {
                return fallback;
            }
            throw;
        }
    }

    private static List<string> BuildLocalSegmentationChunks(string sourceText, int maxTokens)
    {
        var normalized = (sourceText ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return new List<string> { string.Empty };
        }

        // Approximation for llama.cpp local prompt size; keep input chunk conservative.
        var maxChars = Math.Clamp(maxTokens * 6, 1800, 12000);
        if (normalized.Length <= maxChars)
        {
            return new List<string> { normalized };
        }

        var chunks = new List<string>();
        var blocks = Regex.Split(normalized, @"\n\s*\n")
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
        var current = new StringBuilder();
        foreach (var block in blocks)
        {
            if (block.Length >= maxChars)
            {
                if (current.Length > 0)
                {
                    chunks.Add(current.ToString().Trim());
                    current.Clear();
                }

                var start = 0;
                while (start < block.Length)
                {
                    var len = FindSafeCutLength(block, start, maxChars);
                    var piece = block.Substring(start, len).Trim();
                    if (!string.IsNullOrWhiteSpace(piece))
                    {
                        chunks.Add(piece);
                    }
                    start += len;
                    while (start < block.Length && char.IsWhiteSpace(block[start]))
                    {
                        start++;
                    }
                }
                continue;
            }

            if (current.Length == 0)
            {
                current.Append(block);
                continue;
            }

            if (current.Length + 2 + block.Length > maxChars)
            {
                chunks.Add(current.ToString().Trim());
                current.Clear();
                current.Append(block);
            }
            else
            {
                current.AppendLine();
                current.AppendLine();
                current.Append(block);
            }
        }

        if (current.Length > 0)
        {
            chunks.Add(current.ToString().Trim());
        }

        if (chunks.Count == 0)
        {
            chunks.Add(normalized);
        }

        return chunks;
    }

    private static int FindSafeCutLength(string text, int start, int maxChars)
    {
        var remaining = text.Length - start;
        if (remaining <= maxChars)
        {
            return remaining;
        }

        var hardEnd = start + maxChars;
        var softStart = Math.Max(start + Math.Max(200, (int)(maxChars * 0.55)), start + 1);
        if (softStart >= hardEnd)
        {
            softStart = start + 1;
        }

        for (var i = hardEnd; i >= softStart; i--)
        {
            var c = text[i - 1];
            if (IsStrongBoundary(c) && !IsInsideQuoteSpan(text, start, i))
            {
                return i - start;
            }
        }

        for (var i = hardEnd; i >= softStart; i--)
        {
            var c = text[i - 1];
            if ((c == ',' || c == ';' || c == ':') && !IsInsideQuoteSpan(text, start, i))
            {
                return i - start;
            }
        }

        return maxChars;
    }

    private static bool IsStrongBoundary(char c)
    {
        return c == '\n' || c == '.' || c == '!' || c == '?' || c == '"' || c == '”';
    }

    private static bool IsInsideQuoteSpan(string text, int start, int endExclusive)
    {
        var quoteCount = 0;
        for (var i = start; i < endExclusive; i++)
        {
            var c = text[i];
            if (c == '"' || c == '“' || c == '”')
            {
                quoteCount++;
            }
        }

        return (quoteCount % 2) != 0;
    }

    public async Task<List<PreparedScriptPart>> GenerateInstructionsAsync(IReadOnlyList<PreparedScriptPart> parts, CancellationToken ct, IProgress<ScriptPrepProgress>? progress = null)
    {
        var model = ResolveModelForTask(isInstruction: true);
        var temp = _cfg.LlmPrepTemperatureInstruction;
        var maxTokens = _cfg.LlmPrepMaxTokensInstruction;
        var map = new Dictionary<int, string>();
        var ordered = parts.OrderBy(p => p.Order).ToList();
        var unlocked = ordered.Where(p => !p.Locked).ToList();
        var total = unlocked.Count;
        if (ordered.Count == 0)
        {
            return new List<PreparedScriptPart>();
        }
        if (total == 0)
        {
            progress?.Report(new ScriptPrepProgress { Completed = 1, Total = 1, Stage = "All parts are locked. Nothing to regenerate." });
            return ordered.Select(p => new PreparedScriptPart
            {
                Id = p.Id,
                Order = p.Order,
                Text = p.Text,
                SpeakerTag = p.SpeakerTag,
                VoicePath = p.VoicePath,
                Locked = p.Locked,
                Instruction = p.Instruction
            }).ToList();
        }

        var batchSize = Math.Clamp(total <= 10 ? total : 8, 1, 8);
        var done = 0;
        progress?.Report(new ScriptPrepProgress { Completed = 0, Total = total, Stage = "Preparing instruction batches..." });

        for (var i = 0; i < unlocked.Count; i += batchSize)
        {
            ct.ThrowIfCancellationRequested();
            var batch = unlocked.Skip(i).Take(batchSize).ToList();
            progress?.Report(new ScriptPrepProgress
            {
                Completed = done,
                Total = total,
                Stage = $"Generating instructions for parts {done + 1}-{Math.Min(done + batch.Count, total)}..."
            });

            var prompt = BuildInstructionPrompt(batch);
            var content = await RunPromptAsync(model, temp, maxTokens, prompt, "instructions", expectJson: false, ct);
            var batchMap = ParseInstructionResult(batch, content);
            foreach (var kv in batchMap)
            {
                map[kv.Key] = kv.Value;
            }

            done += batch.Count;
            progress?.Report(new ScriptPrepProgress { Completed = done, Total = total, Stage = $"Applied {done}/{total} instruction parts..." });
        }

        var output = new List<PreparedScriptPart>(parts.Count);
        for (var i = 0; i < parts.Count; i++)
        {
            var part = parts[i];
            var key = part.Order + 1;
            map.TryGetValue(key, out var inst);
            output.Add(new PreparedScriptPart
            {
                Id = part.Id,
                Order = part.Order,
                Text = part.Text,
                SpeakerTag = part.SpeakerTag,
                VoicePath = part.VoicePath,
                Locked = part.Locked,
                Instruction = part.Locked ? part.Instruction : SanitizeInstruction(inst)
            });
        }
        return output;
    }

    public async Task<List<PreparedScriptPart>> AutoPrepareAsync(string sourceText, CancellationToken ct, IProgress<ScriptPrepProgress>? progress = null)
    {
        progress?.Report(new ScriptPrepProgress { Completed = 0, Total = 2, Stage = "Splitting source text..." });
        var split = await SplitAsync(sourceText, ct);
        progress?.Report(new ScriptPrepProgress { Completed = 1, Total = 2, Stage = "Generating instructions..." });
        var instructed = await GenerateInstructionsAsync(split, ct);
        progress?.Report(new ScriptPrepProgress { Completed = 2, Total = 2, Stage = "Auto-prepare done." });
        return instructed;
    }

    public Task<string> RunCustomPromptAsync(
        string prompt,
        bool useInstructionModel,
        double temperature,
        int maxTokens,
        CancellationToken ct)
    {
        var model = ResolveModelForTask(useInstructionModel);
        var taskName = useInstructionModel ? "instructions" : "segmentation";
        return RunPromptAsync(model, temperature, maxTokens, prompt, taskName, expectJson: false, ct);
    }

    private async Task<string> RunPromptAsync(
        string model,
        double temperature,
        int maxTokens,
        string prompt,
        string taskName,
        bool expectJson,
        CancellationToken ct)
    {
        var provider = (_cfg.LlmPrepProvider ?? "local").Trim().ToLowerInvariant();
        return provider switch
        {
            "local" => await RunLocalLlamaAsync(model, temperature, maxTokens, prompt, taskName, expectJson, ct),
            "openai" => await RunOpenAiChatAsync(model, temperature, maxTokens, prompt, expectJson, ct),
            "alibaba" => await RunAlibabaChatAsync(model, temperature, maxTokens, prompt, expectJson, ct),
            _ => throw new InvalidOperationException($"Unsupported LLM provider: {provider}")
        };
    }

    private async Task<string> RunLocalLlamaAsync(
        string modelName,
        double temperature,
        int maxTokens,
        string prompt,
        string taskName,
        bool expectJson,
        CancellationToken ct)
    {
        var modelFile = ResolveLocalModelFilePath(taskName == "instructions");
        if (LocalLlmServerManager.IsRunning)
        {
            await LocalLlmServerManager.EnsureRunningForModelAsync(_cfg, modelFile, ct);
            return await LocalLlmServerManager.RunChatCompletionAsync(
                temperature,
                maxTokens,
                prompt,
                expectJson,
                ct);
        }

        var exe = (_cfg.LlmLocalRuntimePath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(exe) || !File.Exists(exe))
        {
            throw new InvalidOperationException("Local LLM runtime not found. Set llama.cpp executable path in Settings > LLM Script Prep.");
        }
        if (!File.Exists(modelFile))
        {
            throw new InvalidOperationException($"Local LLM model missing: {modelFile}");
        }

        var tempPromptFile = Path.Combine(Path.GetTempPath(), $"scriptprep_prompt_{Guid.NewGuid():N}.txt");
        await File.WriteAllTextAsync(tempPromptFile, prompt, Encoding.UTF8, ct);
        try
        {
            var args =
                $"-m \"{modelFile}\" -n {Math.Clamp(maxTokens, 64, 8192)} --temp {temperature.ToString("0.###", CultureInfo.InvariantCulture)} " +
                $"--top-p 0.95 --top-k 40 -f \"{tempPromptFile}\"";
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start local LLM runtime process.");
            using var reg = ct.Register(() =>
            {
                try { if (!p.HasExited) p.Kill(entireProcessTree: true); } catch { }
            });
            var stdoutTask = p.StandardOutput.ReadToEndAsync();
            var stderrTask = p.StandardError.ReadToEndAsync();
            await p.WaitForExitAsync(ct);
            var stdout = await stdoutTask;
            var stderr = await stderrTask;
            if (p.ExitCode != 0)
            {
                throw new InvalidOperationException($"Local LLM process failed ({p.ExitCode}): {TrimForError(stderr)}");
            }

            var content = expectJson ? ExtractJsonPayload(stdout) : stdout?.Trim();
            if (string.IsNullOrWhiteSpace(content))
            {
                throw new InvalidOperationException(expectJson
                    ? "Local LLM returned no JSON content."
                    : "Local LLM returned empty content.");
            }

            return content;
        }
        finally
        {
            try { File.Delete(tempPromptFile); } catch { }
        }
    }

    private async Task<string> RunOpenAiChatAsync(
        string model,
        double temperature,
        int maxTokens,
        string prompt,
        bool expectJson,
        CancellationToken ct)
    {
        var apiKey = (_cfg.ApiKeyOpenAi ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("OpenAI API key missing. Add it in Settings > API Settings.");
        }

        var modelId = string.IsNullOrWhiteSpace(model) ? "gpt-4.1-mini" : model;
        var firstTryUseMaxCompletionTokens = modelId.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase);

        var (ok, payload, statusCode, reasonPhrase) = await SendOpenAiChatCompletionAsync(
            apiKey,
            modelId,
            temperature,
            maxTokens,
            prompt,
            expectJson,
            useMaxCompletionTokens: firstTryUseMaxCompletionTokens,
            ct);

        if (!ok && IsUnsupportedOpenAiTokenParameter(payload))
        {
            (ok, payload, statusCode, reasonPhrase) = await SendOpenAiChatCompletionAsync(
                apiKey,
                modelId,
                temperature,
                maxTokens,
                prompt,
                expectJson,
                useMaxCompletionTokens: !firstTryUseMaxCompletionTokens,
                ct);
        }

        if (!ok)
        {
            throw new InvalidOperationException($"OpenAI LLM error: {statusCode} {reasonPhrase} | {TrimForError(payload)}");
        }

        var content = TryReadString(payload, "choices", 0, "message", "content");
        if (string.IsNullOrWhiteSpace(content))
        {
            throw new InvalidOperationException("OpenAI LLM response missing content.");
        }
        return expectJson ? (ExtractJsonPayload(content) ?? content) : content;
    }

    private async Task<(bool Ok, string Payload, int StatusCode, string ReasonPhrase)> SendOpenAiChatCompletionAsync(
        string apiKey,
        string model,
        double temperature,
        int maxTokens,
        string prompt,
        bool expectJson,
        bool useMaxCompletionTokens,
        CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        var body = new Dictionary<string, object?>
        {
            ["model"] = model,
            ["temperature"] = Math.Clamp(temperature, 0.0, 1.5),
            [useMaxCompletionTokens ? "max_completion_tokens" : "max_tokens"] = Math.Clamp(maxTokens, 64, 8192),
            ["messages"] = new object[]
            {
                new
                {
                    role = "system",
                    content = expectJson
                        ? "Return valid JSON only."
                        : "Follow the exact output format requested by the user prompt. No markdown."
                },
                new { role = "user", content = prompt }
            }
        };

        if (expectJson)
        {
            body["response_format"] = new { type = "json_object" };
        }

        req.Content = JsonContent.Create(body);
        using var res = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        var payload = await res.Content.ReadAsStringAsync(ct);
        return (res.IsSuccessStatusCode, payload, (int)res.StatusCode, res.ReasonPhrase ?? string.Empty);
    }

    private static bool IsUnsupportedOpenAiTokenParameter(string payload)
    {
        var p = (payload ?? string.Empty).ToLowerInvariant();
        if (!p.Contains("unsupported_parameter", StringComparison.Ordinal))
        {
            return false;
        }

        return p.Contains("max_tokens", StringComparison.Ordinal) ||
               p.Contains("max_completion_tokens", StringComparison.Ordinal);
    }

    private async Task<string> RunAlibabaChatAsync(
        string model,
        double temperature,
        int maxTokens,
        string prompt,
        bool expectJson,
        CancellationToken ct)
    {
        var apiKey = (_cfg.ApiKeyAlibaba ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("Alibaba API key missing. Add it in Settings > API Settings.");
        }

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        req.Content = JsonContent.Create(new
        {
            model = string.IsNullOrWhiteSpace(model) ? "qwen-plus" : model,
            temperature = Math.Clamp(temperature, 0.0, 1.5),
            max_tokens = Math.Clamp(maxTokens, 64, 8192),
            messages = new object[]
            {
                new { role = "system", content = expectJson ? "Return valid JSON only." : "Follow the exact output format requested by the user prompt. No markdown." },
                new { role = "user", content = prompt }
            }
        });

        using var res = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        var payload = await res.Content.ReadAsStringAsync(ct);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Alibaba LLM error: {(int)res.StatusCode} {res.ReasonPhrase} | {TrimForError(payload)}");
        }

        var content = TryReadString(payload, "choices", 0, "message", "content");
        if (string.IsNullOrWhiteSpace(content))
        {
            throw new InvalidOperationException("Alibaba LLM response missing content.");
        }
        return expectJson ? (ExtractJsonPayload(content) ?? content) : content;
    }

    private string ResolveModelForTask(bool isInstruction)
    {
        if (_cfg.LlmPrepUseSeparateModels)
        {
            return isInstruction
                ? (_cfg.LlmPrepInstructionModel ?? string.Empty).Trim()
                : (_cfg.LlmPrepSplitModel ?? string.Empty).Trim();
        }

        return (_cfg.LlmPrepSplitModel ?? string.Empty).Trim();
    }

    private string ResolveLocalModelFilePath(bool isInstruction)
    {
        var root = RuntimePathResolver.AppRoot;
        var modelDir = (_cfg.LlmLocalModelDir ?? "models/llm").Trim().Replace('/', Path.DirectorySeparatorChar);
        var fullDir = Path.IsPathRooted(modelDir) ? modelDir : Path.Combine(root, modelDir);
        var file = isInstruction
            ? (_cfg.LlmLocalInstructionModelFile ?? string.Empty).Trim()
            : (_cfg.LlmLocalSplitModelFile ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(file))
        {
            throw new InvalidOperationException("Local LLM model file is not configured in settings.");
        }
        return Path.Combine(fullDir, file);
    }

    private static string BuildSegmentationPrompt(string sourceText)
    {
        var text = sourceText ?? string.Empty;
        return
            "Task: Split this audiobook chapter into ordered speaking parts for TTS generation.\n" +
            "Rules:\n" +
            "- Return JSON only, no markdown.\n" +
            "- Schema: {\"parts\":[{\"order\":1,\"speaker\":\"Narrator|Speaker_01|Speaker_02\",\"text\":\"...\"}]}\n" +
            "- Preserve full source content in order.\n" +
            "- Keep parts natural by sentence/dialogue boundaries.\n" +
            "- Prefer one speaking intent per part (narration OR one speaker line).\n" +
            "- Do not summarize, rewrite, or omit words.\n" +
            "- Preserve dialogue quote marks exactly when they exist in source text.\n" +
            "- Use speaker=Narrator for narration, Speaker_XX for dialogue.\n\n" +
            "Source text:\n" + text;
    }

    private static string BuildInstructionPrompt(IReadOnlyList<PreparedScriptPart> parts)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Task: Generate one short TTS instruction per part for expressive audiobook delivery.");
        sb.AppendLine("Rules:");
        sb.AppendLine("- Return plain text only. No JSON. No markdown.");
        sb.AppendLine("- Output format per line: <order>|<instruction>");
        sb.AppendLine("- Keep each instruction 8 to 20 words.");
        sb.AppendLine("- Mention tone, pacing, emotion, clarity, and emphasis naturally.");
        sb.AppendLine("- Write one concise sentence per part.");
        sb.AppendLine("- Do not include brackets, quotes, code, or extra explanation.");
        sb.AppendLine();
        sb.AppendLine("10 style examples:");
        sb.AppendLine("1) calm narrator, warm neutral tone, clear diction, steady pacing, smooth sentence flow");
        sb.AppendLine("2) tense dialogue, slightly faster pace, clipped delivery, stronger consonants, controlled urgency");
        sb.AppendLine("3) reflective inner thought, softer volume, slower pace, gentle pauses, intimate tone");
        sb.AppendLine("4) confident statement, medium pace, firm emphasis on key words, crisp clarity");
        sb.AppendLine("5) suspicious whisper-like tone, low energy, short pauses, restrained emotion");
        sb.AppendLine("6) angry outburst, faster pace, sharp attack, heavy stress, brief breaths");
        sb.AppendLine("7) sad confession, slower pace, fragile tone, softened consonants, lingering pauses");
        sb.AppendLine("8) excited reveal, energetic pace, bright tone, punchy emphasis, clean articulation");
        sb.AppendLine("9) authoritative narration, stable pace, full-bodied tone, precise diction, balanced rhythm");
        sb.AppendLine("10) fearful reaction, shaky urgency, uneven breath feel, tightened phrasing, clear words");
        sb.AppendLine();
        sb.AppendLine("Parts:");
        foreach (var part in parts.OrderBy(p => p.Order))
        {
            sb.AppendLine($"[{part.Order + 1}] speaker={part.SpeakerTag}");
            sb.AppendLine(part.Text);
            sb.AppendLine();
        }
        sb.AppendLine("Now return exactly one line per part in the required format.");

        return sb.ToString();
    }

    private static List<PreparedScriptPart> ParseSegmentationResult(string content)
    {
        var json = ExtractJsonPayload(content) ?? content;
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("parts", out var partsEl) || partsEl.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Segmentation output missing 'parts' array.");
        }

        var result = new List<PreparedScriptPart>();
        var i = 0;
        foreach (var item in partsEl.EnumerateArray())
        {
            var text = item.TryGetProperty("text", out var tEl) && tEl.ValueKind == JsonValueKind.String ? (tEl.GetString() ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }
            var speaker = item.TryGetProperty("speaker", out var sEl) && sEl.ValueKind == JsonValueKind.String ? (sEl.GetString() ?? "Narrator").Trim() : "Narrator";
            result.Add(new PreparedScriptPart
            {
                Id = Guid.NewGuid().ToString("N"),
                Order = i++,
                Text = text,
                SpeakerTag = string.IsNullOrWhiteSpace(speaker) ? "Narrator" : speaker,
                Instruction = string.Empty
            });
        }

        if (result.Count == 0)
        {
            throw new InvalidOperationException("Segmentation returned no usable parts.");
        }
        return result;
    }

    private static Dictionary<int, string> ParseInstructionResult(IReadOnlyList<PreparedScriptPart> parts, string content)
    {
        var map = new Dictionary<int, string>();
        if (string.IsNullOrWhiteSpace(content))
        {
            return map;
        }

        var lines = content
            .Replace("\r\n", "\n")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        foreach (var line in lines)
        {
            var m = Regex.Match(line, @"^\s*(\d+)\s*[\|\:\-\)]\s*(.+?)\s*$");
            if (!m.Success)
            {
                continue;
            }

            if (!int.TryParse(m.Groups[1].Value, out var order))
            {
                continue;
            }

            var instruction = SanitizeInstruction(m.Groups[2].Value);
            if (order <= 0 || string.IsNullOrWhiteSpace(instruction))
            {
                continue;
            }

            map[order] = instruction;
        }

        if (map.Count == 0)
        {
            var fallbackLines = lines
                .Select(l => Regex.Replace(l, @"^\s*[\-\*\d\.\)\]]+\s*", string.Empty).Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Take(parts.Count)
                .ToList();
            for (var i = 0; i < fallbackLines.Count; i++)
            {
                map[i + 1] = SanitizeInstruction(fallbackLines[i]);
            }
        }

        return map;
    }

    private static string SanitizeInstruction(string? value)
    {
        var text = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        text = text.Replace("[", string.Empty, StringComparison.Ordinal).Replace("]", string.Empty, StringComparison.Ordinal);
        if (text.Length > 220)
        {
            text = text[..220].Trim();
        }
        return text;
    }

    private static string? ExtractJsonPayload(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return null;
        }

        var s = content.Trim();
        if (s.StartsWith("```", StringComparison.Ordinal))
        {
            var fence = Regex.Match(s, @"```(?:json)?\s*(.*?)\s*```", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (fence.Success)
            {
                s = fence.Groups[1].Value.Trim();
            }
        }

        var startObj = s.IndexOf('{');
        var endObj = s.LastIndexOf('}');
        if (startObj >= 0 && endObj > startObj)
        {
            return s.Substring(startObj, endObj - startObj + 1);
        }

        return s;
    }

    private static string TrimForError(string payload)
    {
        var text = (payload ?? string.Empty).Replace('\r', ' ').Replace('\n', ' ').Trim();
        return text.Length <= 400 ? text : text[..400] + "...";
    }

    private static string? TryReadString(string json, string arrayKey, int idx, params string[] path)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty(arrayKey, out var arr) || arr.ValueKind != JsonValueKind.Array || arr.GetArrayLength() <= idx)
            {
                return null;
            }
            var cursor = arr[idx];
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

    private static class LocalLlmServerManager
    {
        private static readonly object Sync = new();
        private static readonly SemaphoreSlim StartStopGate = new(1, 1);
        private static Process? _process;
        private static Task? _stdoutPumpTask;
        private static Task? _stderrPumpTask;
        private static string? _loadedModelFile;
        private static int _port = 8091;
        private static readonly List<string> RecentLogs = new();
        private static readonly HttpClient ServerHttp = new()
        {
            Timeout = Timeout.InfiniteTimeSpan
        };

        static LocalLlmServerManager()
        {
            AppDomain.CurrentDomain.ProcessExit += (_, _) => ForceStopSync();
        }

        public static bool IsRunning
        {
            get
            {
                lock (Sync)
                {
                    return _process is { HasExited: false };
                }
            }
        }

        public static async Task EnsureRunningForModelAsync(AppConfig cfg, string modelFile, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(modelFile) || !File.Exists(modelFile))
            {
                throw new InvalidOperationException("Local LLM model file not found.");
            }

            await StartStopGate.WaitAsync(ct);
            try
            {
                lock (Sync)
                {
                    if (_process is { HasExited: false } &&
                        string.Equals(_loadedModelFile, modelFile, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }

                await StopInternalAsync();

                var runtimeExe = ResolveServerExecutable(cfg);
                var chosenPort = FindAvailablePort(8091);
                var args = BuildServerArguments(runtimeExe, modelFile, chosenPort);
                var psi = new ProcessStartInfo
                {
                    FileName = runtimeExe,
                    Arguments = args,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start local LLM server.");
                var stdout = Task.Run(async () =>
                {
                    while (!process.HasExited)
                    {
                        var line = await process.StandardOutput.ReadLineAsync();
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            AppendRecentLog("[out] " + line);
                        }
                    }
                });
                var stderr = Task.Run(async () =>
                {
                    while (!process.HasExited)
                    {
                        var line = await process.StandardError.ReadLineAsync();
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            AppendRecentLog("[err] " + line);
                        }
                    }
                });

                lock (Sync)
                {
                    _process = process;
                    _stdoutPumpTask = stdout;
                    _stderrPumpTask = stderr;
                    _loadedModelFile = modelFile;
                    _port = chosenPort;
                }

                await WaitForServerReadyAsync(ct);
            }
            finally
            {
                StartStopGate.Release();
            }
        }

        public static async Task<string> RunChatCompletionAsync(
            double temperature,
            int maxTokens,
            string prompt,
            bool expectJson,
            CancellationToken ct)
        {
            if (!IsRunning)
            {
                throw new InvalidOperationException("Local LLM server is not running.");
            }

            string lastPayload = string.Empty;
            for (var attempt = 1; attempt <= 2; attempt++)
            {
                using var req = new HttpRequestMessage(HttpMethod.Post, $"http://127.0.0.1:{_port}/v1/chat/completions");
                var body = new
                {
                    model = "local",
                    temperature = Math.Clamp(temperature, 0.0, 1.5),
                    max_tokens = Math.Clamp(maxTokens, 64, 8192),
                    stream = false,
                    messages = new object[]
                    {
                        new { role = "system", content = "Return valid JSON only." },
                        new { role = "user", content = prompt }
                    }
                };
                req.Content = JsonContent.Create(body);

                using var res = await ServerHttp.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
                var payload = await res.Content.ReadAsStringAsync(ct);
                lastPayload = payload;
                AppendRecentLog($"[chat:{attempt}] status={(int)res.StatusCode}");
                if (!res.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"Local server error: {(int)res.StatusCode} {res.ReasonPhrase} | {TrimForError(payload)}");
                }

                var content = ExtractLocalServerChatContent(payload);
                if (!string.IsNullOrWhiteSpace(content))
                {
                    WriteLlmIoDebug(prompt, payload, mode: $"chat:{attempt}");
                    return expectJson ? (ExtractJsonPayload(content) ?? content) : content;
                }

                await Task.Delay(180, ct);
            }

            // Fallback for servers that expose better plain completion than chat template.
            for (var attempt = 1; attempt <= 2; attempt++)
            {
                using var req = new HttpRequestMessage(HttpMethod.Post, $"http://127.0.0.1:{_port}/completion");
                var plainPrompt =
                    "Return valid JSON only.\n" +
                    prompt + "\n";
                var body = new
                {
                    prompt = plainPrompt,
                    temperature = Math.Clamp(temperature, 0.0, 1.5),
                    n_predict = Math.Clamp(maxTokens, 64, 8192),
                    stream = false
                };
                req.Content = JsonContent.Create(body);

                using var res = await ServerHttp.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
                var payload = await res.Content.ReadAsStringAsync(ct);
                lastPayload = payload;
                AppendRecentLog($"[completion:{attempt}] status={(int)res.StatusCode}");
                if (!res.IsSuccessStatusCode)
                {
                    continue;
                }

                var text = ExtractLocalCompletionText(payload);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    WriteLlmIoDebug(prompt, payload, mode: $"completion:{attempt}");
                    return expectJson ? (ExtractJsonPayload(text) ?? text) : text;
                }

                await Task.Delay(180, ct);
            }

            throw new InvalidOperationException(
                $"Local server returned empty content. Payload: {TrimForError(lastPayload)} | {GetRecentLogTail()}");
        }

        public static Task StopAsync()
        {
            return StopInternalAsync();
        }

        private static Task StopInternalAsync()
        {
            Process? process;
            lock (Sync)
            {
                process = _process;
                _process = null;
                _stdoutPumpTask = null;
                _stderrPumpTask = null;
                _loadedModelFile = null;
            }

            if (process is null)
            {
                return Task.CompletedTask;
            }

            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // best effort
            }
            finally
            {
                try { process.Dispose(); } catch { }
            }

            return Task.CompletedTask;
        }

        public static void ForceStopSync()
        {
            Process? process;
            lock (Sync)
            {
                process = _process;
                _process = null;
                _stdoutPumpTask = null;
                _stderrPumpTask = null;
                _loadedModelFile = null;
            }

            if (process is null)
            {
                return;
            }

            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // best effort
            }
            finally
            {
                try { process.Dispose(); } catch { }
            }
        }

        private static string ResolveServerExecutable(AppConfig cfg)
        {
            var runtimeExe = (cfg.LlmLocalRuntimePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(runtimeExe) || !File.Exists(runtimeExe))
            {
                throw new InvalidOperationException("Local LLM runtime not found. Set llama.cpp runtime path in settings.");
            }

            var fileName = Path.GetFileName(runtimeExe);
            if (fileName.Contains("server", StringComparison.OrdinalIgnoreCase))
            {
                return runtimeExe;
            }

            var dir = Path.GetDirectoryName(runtimeExe) ?? string.Empty;
            var serverExe = Path.Combine(dir, "llama-server.exe");
            if (File.Exists(serverExe))
            {
                return serverExe;
            }

            return runtimeExe;
        }

        private static string BuildServerArguments(string runtimeExe, string modelFile, int port)
        {
            var fileName = Path.GetFileName(runtimeExe).ToLowerInvariant();
            var baseArgs =
                $"-m \"{modelFile}\" --host 127.0.0.1 --port {port} --ctx-size 8192";
            if (fileName.Contains("server", StringComparison.OrdinalIgnoreCase))
            {
                return baseArgs;
            }

            return $"--server {baseArgs}";
        }

        private static async Task WaitForServerReadyAsync(CancellationToken ct)
        {
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                if (!IsRunning)
                {
                    var detail = GetRecentLogTail();
                    throw new InvalidOperationException(
                        $"Local LLM server exited during startup. {detail}");
                }

                try
                {
                    using var health = await ServerHttp.GetAsync($"http://127.0.0.1:{_port}/health", ct);
                    if (health.IsSuccessStatusCode)
                    {
                        return;
                    }
                }
                catch
                {
                    // keep polling until timeout
                }

                try
                {
                    using var models = await ServerHttp.GetAsync($"http://127.0.0.1:{_port}/v1/models", ct);
                    if (models.IsSuccessStatusCode)
                    {
                        return;
                    }
                }
                catch
                {
                    // keep polling until timeout
                }

                await Task.Delay(350, ct);
            }
        }

        private static void AppendRecentLog(string line)
        {
            lock (Sync)
            {
                RecentLogs.Add(line);
                if (RecentLogs.Count > 120)
                {
                    RecentLogs.RemoveRange(0, RecentLogs.Count - 120);
                }
            }
        }

        private static string GetRecentLogTail()
        {
            lock (Sync)
            {
                if (RecentLogs.Count == 0)
                {
                    return "No startup log output was captured.";
                }

                var tail = RecentLogs.Skip(Math.Max(0, RecentLogs.Count - 8)).ToArray();
                return "Recent log: " + string.Join(" | ", tail);
            }
        }

        private static string? ExtractLocalServerChatContent(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
            {
                return null;
            }

            try
            {
                using var doc = JsonDocument.Parse(payload);
                if (!doc.RootElement.TryGetProperty("choices", out var choices) ||
                    choices.ValueKind != JsonValueKind.Array ||
                    choices.GetArrayLength() == 0)
                {
                    return null;
                }

                var first = choices[0];
                if (first.ValueKind != JsonValueKind.Object)
                {
                    return null;
                }

                if (first.TryGetProperty("message", out var msg) && msg.ValueKind == JsonValueKind.Object)
                {
                    if (msg.TryGetProperty("content", out var content))
                    {
                        var text = ReadContentNode(content);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            return text;
                        }
                    }
                }

                if (first.TryGetProperty("text", out var textNode) && textNode.ValueKind == JsonValueKind.String)
                {
                    return textNode.GetString();
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string? ExtractLocalCompletionText(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
            {
                return null;
            }

            try
            {
                using var doc = JsonDocument.Parse(payload);
                if (doc.RootElement.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.String)
                {
                    var s = content.GetString();
                    return string.IsNullOrWhiteSpace(s) ? null : s;
                }
                if (doc.RootElement.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
                {
                    var s = text.GetString();
                    return string.IsNullOrWhiteSpace(s) ? null : s;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string? ReadContentNode(JsonElement content)
        {
            if (content.ValueKind == JsonValueKind.String)
            {
                return content.GetString();
            }

            if (content.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var sb = new StringBuilder();
            foreach (var item in content.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                {
                    sb.Append(item.GetString());
                    continue;
                }

                if (item.ValueKind == JsonValueKind.Object &&
                    item.TryGetProperty("text", out var textNode) &&
                    textNode.ValueKind == JsonValueKind.String)
                {
                    sb.Append(textNode.GetString());
                }
            }

            var merged = sb.ToString().Trim();
            return string.IsNullOrWhiteSpace(merged) ? null : merged;
        }

        private static int FindAvailablePort(int preferredPort)
        {
            if (IsPortAvailable(preferredPort))
            {
                return preferredPort;
            }

            for (var port = preferredPort + 1; port <= preferredPort + 50; port++)
            {
                if (IsPortAvailable(port))
                {
                    return port;
                }
            }

            throw new InvalidOperationException("No free local port found for LLM server startup.");
        }

        private static bool IsPortAvailable(int port)
        {
            try
            {
                var listener = new TcpListener(System.Net.IPAddress.Loopback, port);
                listener.Start();
                listener.Stop();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void WriteLlmIoDebug(string prompt, string payload, string mode)
        {
            try
            {
                var logDir = Path.Combine(RuntimePathResolver.AppRoot, "logs");
                Directory.CreateDirectory(logDir);
                var path = Path.Combine(logDir, "scriptprep_llm_io.log");
                var text =
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] mode={mode}\n" +
                    $"PROMPT:\n{TrimTo(prompt, 4000)}\n" +
                    $"RESPONSE:\n{TrimTo(payload, 4000)}\n" +
                    "----\n";
                File.AppendAllText(path, text);
            }
            catch
            {
                // best effort debug log only
            }
        }

        private static string TrimTo(string s, int max)
        {
            if (string.IsNullOrEmpty(s))
            {
                return string.Empty;
            }
            return s.Length <= max ? s : s[..max];
        }
    }
}
