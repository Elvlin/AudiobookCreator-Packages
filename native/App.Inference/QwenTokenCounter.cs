using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using App.Core.Runtime;

namespace App.Inference;

public sealed class QwenTokenCounter : IDisposable
{
    private readonly nint _tokenizer;

    public int MaxPositionEmbeddings { get; }
    public int MaxNewTokens { get; }
    public int PromptOverheadTokens { get; }
    public int RecommendedPromptTokenBudget { get; }

    private QwenTokenCounter(nint tokenizer, int maxPositionEmbeddings, int maxNewTokens, int promptOverheadTokens, int recommendedPromptTokenBudget)
    {
        _tokenizer = tokenizer;
        MaxPositionEmbeddings = maxPositionEmbeddings;
        MaxNewTokens = maxNewTokens;
        PromptOverheadTokens = promptOverheadTokens;
        RecommendedPromptTokenBudget = recommendedPromptTokenBudget;
    }

    public static bool TryCreate(
        string? modelCacheDirRaw,
        string? modelRepoId,
        string? tokenizerRepoId,
        int maxNewTokens,
        out QwenTokenCounter? counter,
        out string error)
    {
        counter = null;
        error = string.Empty;
        try
        {
            var modelCacheDir = ModelCachePath.ResolveAbsolute(modelCacheDirRaw, RuntimePathResolver.AppRoot);
            var mainRepo = string.IsNullOrWhiteSpace(modelRepoId) ? "xkos/Qwen3-TTS-12Hz-1.7B-ONNX" : modelRepoId.Trim();
            var tokRepo = string.IsNullOrWhiteSpace(tokenizerRepoId) ? "zukky/Qwen3-TTS-ONNX-DLL" : tokenizerRepoId.Trim();

            var repoRoot = ResolveRepoRoot(modelCacheDir, mainRepo);
            var (vocabPath, mergesPath, tokenizerConfigPath) = ResolveTokenizerFiles(modelCacheDir, repoRoot, tokRepo);
            EnsureRustDllPresent(modelCacheDir, repoRoot);

            var tokenizer = CreateTokenizer(vocabPath, mergesPath, tokenizerConfigPath);
            if (tokenizer == nint.Zero)
            {
                throw new InvalidOperationException(GetLastNativeError());
            }

            var maxPos = ReadMaxPositionEmbeddings(repoRoot);
            var normalizedMaxNew = Math.Clamp(maxNewTokens <= 0 ? 512 : maxNewTokens, 64, 4096);

            var emptyPromptTokens = CountAssistantPromptTokensInternal(tokenizer, string.Empty);
            var reserve = Math.Clamp(normalizedMaxNew + 384, 512, 4096);
            var rawInputBudget = Math.Max(384, maxPos - reserve);
            var recommended = Math.Clamp(Math.Min(rawInputBudget, Math.Max(1024, maxPos / 4)), 700, 3200);

            counter = new QwenTokenCounter(tokenizer, maxPos, normalizedMaxNew, emptyPromptTokens, recommended);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    public int CountAssistantPromptTokens(string text)
    {
        if (_tokenizer == nint.Zero)
        {
            throw new ObjectDisposedException(nameof(QwenTokenCounter));
        }

        return CountAssistantPromptTokensInternal(_tokenizer, text);
    }

    public void Dispose()
    {
        if (_tokenizer != nint.Zero)
        {
            FreeTokenizer(_tokenizer);
        }
    }

    private static int CountAssistantPromptTokensInternal(nint tokenizer, string text)
    {
        var prompt = BuildAssistantText(text ?? string.Empty);
        var needed = TokenizerEncodeNative(tokenizer, prompt, nint.Zero, 0);
        if (needed == 0)
        {
            throw new InvalidOperationException(GetLastNativeError());
        }
        return checked((int)needed);
    }

    private static string ResolveRepoRoot(string modelCacheDir, string repoId)
    {
        var parts = repoId.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            throw new InvalidOperationException($"Invalid Qwen repo id: {repoId}");
        }

        var repoRoot = Path.Combine(modelCacheDir, "hf-cache", $"models--{parts[0]}--{parts[1]}");
        if (!Directory.Exists(repoRoot))
        {
            throw new DirectoryNotFoundException($"Qwen repo not found: {repoRoot}");
        }
        return repoRoot;
    }

    private static (string VocabPath, string MergesPath, string TokenizerConfigPath) ResolveTokenizerFiles(string modelCacheDir, string repoRoot, string tokenizerRepoId)
    {
        if (TryFindTokenizerTriplet(repoRoot, out var vocab, out var merges, out var cfg))
        {
            return (vocab, merges, cfg);
        }

        try
        {
            var tokenizerRepoRoot = ResolveRepoRoot(modelCacheDir, tokenizerRepoId);
            if (TryFindTokenizerTriplet(tokenizerRepoRoot, out vocab, out merges, out cfg))
            {
                return (vocab, merges, cfg);
            }
        }
        catch
        {
            // Fallback scan below.
        }

        var hfRoot = Path.Combine(modelCacheDir, "hf-cache");
        if (Directory.Exists(hfRoot))
        {
            var candidates = new List<(string VocabPath, string MergesPath, string TokenizerConfigPath)>();
            foreach (var foundVocab in Directory.GetFiles(hfRoot, "vocab.json", SearchOption.AllDirectories))
            {
                var dir = Path.GetDirectoryName(foundVocab);
                if (string.IsNullOrWhiteSpace(dir))
                {
                    continue;
                }
                var foundMerges = Path.Combine(dir, "merges.txt");
                var foundCfg = Path.Combine(dir, "tokenizer_config.json");
                if (File.Exists(foundMerges) && File.Exists(foundCfg))
                {
                    candidates.Add((foundVocab, foundMerges, foundCfg));
                }
            }

            if (candidates.Count > 0)
            {
                var preferred = candidates.FirstOrDefault(c =>
                    c.VocabPath.IndexOf("1.7B", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.VocabPath.IndexOf("12Hz-1.7B-Base", StringComparison.OrdinalIgnoreCase) >= 0);
                return !string.IsNullOrWhiteSpace(preferred.VocabPath) ? preferred : candidates[0];
            }
        }

        throw new FileNotFoundException("Qwen tokenizer files (vocab.json / merges.txt / tokenizer_config.json) not found.");
    }

    private static bool TryFindTokenizerTriplet(string root, out string vocab, out string merges, out string cfg)
    {
        vocab = string.Empty;
        merges = string.Empty;
        cfg = string.Empty;

        var directTokenizer = Path.Combine(root, "tokenizer");
        var directVocab = Path.Combine(directTokenizer, "vocab.json");
        var directMerges = Path.Combine(directTokenizer, "merges.txt");
        var directCfg = Path.Combine(directTokenizer, "tokenizer_config.json");
        if (File.Exists(directVocab) && File.Exists(directMerges) && File.Exists(directCfg))
        {
            vocab = directVocab;
            merges = directMerges;
            cfg = directCfg;
            return true;
        }

        foreach (var foundVocab in Directory.GetFiles(root, "vocab.json", SearchOption.AllDirectories))
        {
            var dir = Path.GetDirectoryName(foundVocab);
            if (string.IsNullOrWhiteSpace(dir))
            {
                continue;
            }
            var foundMerges = Path.Combine(dir, "merges.txt");
            var foundCfg = Path.Combine(dir, "tokenizer_config.json");
            if (File.Exists(foundMerges) && File.Exists(foundCfg))
            {
                vocab = foundVocab;
                merges = foundMerges;
                cfg = foundCfg;
                return true;
            }
        }

        return false;
    }

    private static void EnsureRustDllPresent(string modelCacheDir, string repoRoot)
    {
        var dllSource = ResolveRustDllPath(modelCacheDir, repoRoot);
        var appDll = Path.Combine(RuntimePathResolver.AppRoot, "qwen3_tts_rust.dll");
        if (!File.Exists(appDll) || new FileInfo(appDll).Length != new FileInfo(dllSource).Length)
        {
            File.Copy(dllSource, appDll, true);
        }
    }

    private static string ResolveRustDllPath(string modelCacheDir, string repoRoot)
    {
        var appDll = Path.Combine(RuntimePathResolver.AppRoot, "qwen3_tts_rust.dll");
        if (File.Exists(appDll))
        {
            return appDll;
        }

        var bundled = QwenEmbeddedRuntime.EnsureBundledRustDllAtAppRoot();
        if (File.Exists(bundled))
        {
            return bundled;
        }

        var direct = Path.Combine(repoRoot, "qwen3_tts_rust.dll");
        if (File.Exists(direct))
        {
            return direct;
        }

        var anyInRepo = Directory.GetFiles(repoRoot, "qwen3_tts_rust.dll", SearchOption.AllDirectories).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(anyInRepo))
        {
            return anyInRepo;
        }

        var hfRoot = Path.Combine(modelCacheDir, "hf-cache");
        if (Directory.Exists(hfRoot))
        {
            var any = Directory.GetFiles(hfRoot, "qwen3_tts_rust.dll", SearchOption.AllDirectories).FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(any))
            {
                return any;
            }
        }

        throw new FileNotFoundException("qwen3_tts_rust.dll not found in model cache.");
    }

    private static int ReadMaxPositionEmbeddings(string repoRoot)
    {
        try
        {
            var splitCfg = Path.Combine(repoRoot, "voice_clone_config.json");
            if (File.Exists(splitCfg))
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(splitCfg));
                if (doc.RootElement.TryGetProperty("talker_config", out var talker) &&
                    talker.TryGetProperty("max_position_embeddings", out var maxPos))
                {
                    return Math.Clamp(maxPos.GetInt32(), 1024, 131072);
                }
            }

            foreach (var configPath in Directory.GetFiles(repoRoot, "config.json", SearchOption.AllDirectories))
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
                if (doc.RootElement.TryGetProperty("talker_config", out var talker) &&
                    talker.TryGetProperty("max_position_embeddings", out var maxPos))
                {
                    return Math.Clamp(maxPos.GetInt32(), 1024, 131072);
                }
                if (doc.RootElement.TryGetProperty("max_position_embeddings", out var topLevelMax))
                {
                    return Math.Clamp(topLevelMax.GetInt32(), 1024, 131072);
                }
            }
        }
        catch
        {
            // Fallback below.
        }

        return 8192;
    }

    private static string BuildAssistantText(string text)
    {
        var cap = Math.Max(4096, (text?.Length ?? 0) * 4 + 128);
        var buffer = new byte[cap];
        var handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            _ = BuildAssistantTextNative(text ?? string.Empty, handle.AddrOfPinnedObject(), (nuint)buffer.Length);
            var len = Array.IndexOf(buffer, (byte)0);
            if (len < 0)
            {
                len = buffer.Length;
            }
            return Encoding.UTF8.GetString(buffer, 0, len);
        }
        finally
        {
            handle.Free();
        }
    }

    private static string GetLastNativeError()
    {
        var buffer = new byte[4096];
        var gc = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            _ = LastErrorMessage(gc.AddrOfPinnedObject(), (nuint)buffer.Length);
            var len = Array.IndexOf(buffer, (byte)0);
            if (len < 0)
            {
                len = buffer.Length;
            }
            return Encoding.UTF8.GetString(buffer, 0, len);
        }
        finally
        {
            gc.Free();
        }
    }

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_last_error_message")]
    private static extern nuint LastErrorMessage(nint outBuf, nuint outLen);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_build_assistant_text")]
    private static extern nuint BuildAssistantTextNative([MarshalAs(UnmanagedType.LPUTF8Str)] string text, nint outBuf, nuint outLen);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_create")]
    private static extern nint CreateTokenizer([MarshalAs(UnmanagedType.LPUTF8Str)] string vocabPath, [MarshalAs(UnmanagedType.LPUTF8Str)] string mergesPath, [MarshalAs(UnmanagedType.LPUTF8Str)] string tokenizerConfigPath);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_free")]
    private static extern void FreeTokenizer(nint handle);

    [DllImport("qwen3_tts_rust.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "qwen3_tts_tokenizer_encode")]
    private static extern nuint TokenizerEncodeNative(nint handle, [MarshalAs(UnmanagedType.LPUTF8Str)] string text, nint outIds, nuint outLen);
}
