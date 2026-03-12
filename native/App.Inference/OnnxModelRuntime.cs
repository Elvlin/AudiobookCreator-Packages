using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using App.Core.Runtime;
using System.Diagnostics;

namespace App.Inference;

internal static class OnnxModelRuntime
{
    internal readonly record struct ProsodyControl(
        float SpeedFactor,
        float PitchFactor,
        float EnergyFactor,
        float SpectralTilt
    );
    internal sealed class ProsodyEnvelope
    {
        public required float[] Pitch { get; init; }
        public required float[] Energy { get; init; }
        public required float[] Tilt { get; init; }
    }
    private static readonly object Sync = new();
    private static string _lastPackKey = string.Empty;

    private static InferenceSession? _conditionalDecoder;
    private static InferenceSession? _embedTokens;
    private static InferenceSession? _languageModel;
    private static InferenceSession? _speechEncoder;

    private static readonly string[] RequiredModelFiles =
    {
        "onnx/conditional_decoder.onnx",
        "onnx/embed_tokens.onnx",
        "onnx/language_model.onnx",
        "onnx/speech_encoder.onnx",
        "tokenizer.json"
    };

    public static void EnsureLoaded(string modelCacheDir, string modelRepoId, string? preferDevice = null)
    {
        if (!IsOnnxRepo(modelRepoId))
        {
            return;
        }

        var packRoot = ResolvePackRoot(modelCacheDir, modelRepoId);
        var device = NormalizePreferDevice(preferDevice);
        var key = $"{packRoot.ToLowerInvariant()}|{device}";

        lock (Sync)
        {
            if (_lastPackKey == key && _conditionalDecoder is not null && _embedTokens is not null && _languageModel is not null && _speechEncoder is not null)
            {
                return;
            }

            Unload();
            ValidatePackFiles(packRoot);

            var sessionOptions = CreateSessionOptions(device);

            _conditionalDecoder = new InferenceSession(Path.Combine(packRoot, "onnx", "conditional_decoder.onnx"), sessionOptions);
            _embedTokens = new InferenceSession(Path.Combine(packRoot, "onnx", "embed_tokens.onnx"), sessionOptions);
            _languageModel = new InferenceSession(Path.Combine(packRoot, "onnx", "language_model.onnx"), sessionOptions);
            _speechEncoder = new InferenceSession(Path.Combine(packRoot, "onnx", "speech_encoder.onnx"), sessionOptions);
            _lastPackKey = key;
        }
    }

    public static ProsodyControl ComputeProsodyControl(string modelCacheDir, string modelRepoId, string text, string? preferDevice = null)
    {
        if (!IsOnnxRepo(modelRepoId))
        {
            return new ProsodyControl(1.0f, 1.0f, 1.0f, 0.0f);
        }

        EnsureLoaded(modelCacheDir, modelRepoId, preferDevice);
        lock (Sync)
        {
            if (_embedTokens is null || _languageModel is null || _speechEncoder is null || _conditionalDecoder is null)
            {
                return new ProsodyControl(1.0f, 1.0f, 1.0f, 0.0f);
            }

            double scoreA;
            double scoreB;
            double scoreC;
            double scoreD;
            try
            {
                scoreA = TryRunSessionScore(_embedTokens, text);
                scoreB = TryRunSessionScore(_languageModel, text);
                scoreC = TryRunSessionScore(_speechEncoder, text);
                scoreD = TryRunSessionScore(_conditionalDecoder, text);
            }
            catch
            {
                return new ProsodyControl(1.0f, 1.0f, 1.0f, 0.0f);
            }

            var score = (scoreA + scoreB + scoreC + scoreD) / 4.0;
            if (double.IsNaN(score) || double.IsInfinity(score))
            {
                return new ProsodyControl(1.0f, 1.0f, 1.0f, 0.0f);
            }

            var speed = 0.90f + (float)Math.Min(0.22, Math.Max(0.0, scoreA % 0.22));
            var pitch = 0.92f + (float)Math.Min(0.20, Math.Max(0.0, scoreB % 0.20));
            var energy = 0.88f + (float)Math.Min(0.28, Math.Max(0.0, scoreC % 0.28));
            var tilt = -0.20f + (float)Math.Min(0.40, Math.Max(0.0, scoreD % 0.40));
            return new ProsodyControl(speed, pitch, energy, tilt);
        }
    }

    public static ProsodyEnvelope ComputeProsodyEnvelope(string modelCacheDir, string modelRepoId, string text, int frameCount, string? preferDevice = null)
    {
        frameCount = Math.Clamp(frameCount, 16, 4096);
        if (!IsOnnxRepo(modelRepoId))
        {
            return FlatEnvelope(frameCount, 1.0f, 1.0f, 0.0f);
        }

        EnsureLoaded(modelCacheDir, modelRepoId, preferDevice);
        lock (Sync)
        {
            if (_speechEncoder is null || _conditionalDecoder is null)
            {
                return FlatEnvelope(frameCount, 1.0f, 1.0f, 0.0f);
            }

            float[] a;
            float[] b;
            try
            {
                a = TryRunSessionVector(_speechEncoder, text, frameCount);
                b = TryRunSessionVector(_conditionalDecoder, text, frameCount);
            }
            catch
            {
                return FlatEnvelope(frameCount, 1.0f, 1.0f, 0.0f);
            }

            var pitch = new float[frameCount];
            var energy = new float[frameCount];
            var tilt = new float[frameCount];
            for (var i = 0; i < frameCount; i++)
            {
                var av = a[i % a.Length];
                var bv = b[i % b.Length];
                pitch[i] = (float)Math.Clamp(0.88 + Math.Abs(av) * 0.28, 0.75, 1.30);
                energy[i] = (float)Math.Clamp(0.82 + Math.Abs(bv) * 0.36, 0.70, 1.35);
                var t = (av - bv) * 0.18;
                tilt[i] = (float)Math.Clamp(t, -0.30, 0.30);
            }

            return new ProsodyEnvelope
            {
                Pitch = pitch,
                Energy = energy,
                Tilt = tilt
            };
        }
    }

    private static double TryRunSessionScore(InferenceSession session, string text)
    {
        var inputs = BuildInputsForSession(session, text);
        if (inputs.Count == 0)
        {
            return 0.08;
        }

        using var results = session.Run(inputs);
        double sum = 0.0;
        int count = 0;

        foreach (var r in results)
        {
            switch (r.Value)
            {
                case DenseTensor<float> tf:
                    foreach (var v in tf)
                    {
                        sum += Math.Abs(v);
                        count++;
                        if (count > 4096) break;
                    }
                    break;
                case DenseTensor<double> td:
                    foreach (var v in td)
                    {
                        sum += Math.Abs(v);
                        count++;
                        if (count > 4096) break;
                    }
                    break;
                case DenseTensor<long> tl:
                    foreach (var v in tl)
                    {
                        sum += Math.Abs(v);
                        count++;
                        if (count > 4096) break;
                    }
                    break;
            }

            if (count > 4096)
            {
                break;
            }
        }

        return count == 0 ? 0.08 : sum / count;
    }

    private static float[] TryRunSessionVector(InferenceSession session, string text, int desired)
    {
        var inputs = BuildInputsForSession(session, text);
        using var results = session.Run(inputs);
        foreach (var r in results)
        {
            if (r.Value is DenseTensor<float> tf)
            {
                return ExpandAndNormalize(tf.ToArray(), desired);
            }
            if (r.Value is DenseTensor<double> td)
            {
                var src = td.ToArray();
                var f = new float[src.Length];
                for (var i = 0; i < src.Length; i++) f[i] = (float)src[i];
                return ExpandAndNormalize(f, desired);
            }
            if (r.Value is DenseTensor<long> tl)
            {
                var src = tl.ToArray();
                var f = new float[src.Length];
                for (var i = 0; i < src.Length; i++) f[i] = src[i];
                return ExpandAndNormalize(f, desired);
            }
        }
        return ExpandAndNormalize(new float[] { 0.1f, 0.2f, 0.15f, 0.05f }, desired);
    }

    private static float[] ExpandAndNormalize(float[] src, int desired)
    {
        if (src.Length == 0)
        {
            src = new float[] { 0.1f };
        }
        var max = 1e-6f;
        for (var i = 0; i < src.Length; i++)
        {
            max = Math.Max(max, Math.Abs(src[i]));
        }
        var outv = new float[desired];
        for (var i = 0; i < desired; i++)
        {
            var v = src[i % src.Length] / max;
            outv[i] = v;
        }
        return outv;
    }

    private static ProsodyEnvelope FlatEnvelope(int frames, float pitch, float energy, float tilt)
    {
        var p = Enumerable.Repeat(pitch, frames).ToArray();
        var e = Enumerable.Repeat(energy, frames).ToArray();
        var t = Enumerable.Repeat(tilt, frames).ToArray();
        return new ProsodyEnvelope { Pitch = p, Energy = e, Tilt = t };
    }

    private static List<NamedOnnxValue> BuildInputsForSession(InferenceSession session, string text)
    {
        var list = new List<NamedOnnxValue>();
        var tokenIds = BuildTokenIds(text);

        foreach (var kv in session.InputMetadata)
        {
            var name = kv.Key;
            var meta = kv.Value;

            if (!meta.IsTensor)
            {
                continue;
            }

            var dims = NormalizeDims(meta.Dimensions, tokenIds.Length);
            var elemType = meta.ElementType;

            if (elemType == typeof(long))
            {
                var tensor = new DenseTensor<long>(dims);
                FillLongTensor(tensor, name, tokenIds);
                list.Add(NamedOnnxValue.CreateFromTensor(name, tensor));
            }
            else if (elemType == typeof(int))
            {
                var tensor = new DenseTensor<int>(dims);
                FillIntTensor(tensor, name, tokenIds);
                list.Add(NamedOnnxValue.CreateFromTensor(name, tensor));
            }
            else if (elemType == typeof(float))
            {
                var tensor = new DenseTensor<float>(dims);
                FillFloatTensor(tensor, name, tokenIds);
                list.Add(NamedOnnxValue.CreateFromTensor(name, tensor));
            }
            else if (elemType == typeof(double))
            {
                var tensor = new DenseTensor<double>(dims);
                FillDoubleTensor(tensor, name, tokenIds);
                list.Add(NamedOnnxValue.CreateFromTensor(name, tensor));
            }
            else if (elemType == typeof(bool))
            {
                var tensor = new DenseTensor<bool>(dims);
                FillBoolTensor(tensor, name);
                list.Add(NamedOnnxValue.CreateFromTensor(name, tensor));
            }
        }

        return list;
    }

    private static int[] NormalizeDims(IReadOnlyList<int> rawDims, int tokenLen)
    {
        if (rawDims.Count == 0)
        {
            return new[] { 1 };
        }

        var dims = new int[rawDims.Count];
        for (var i = 0; i < rawDims.Count; i++)
        {
            var d = rawDims[i];
            if (d <= 0)
            {
                // Prefer batch=1, sequence=tokenLen, fallback=1.
                if (i == 0) d = 1;
                else if (i == 1) d = Math.Max(1, Math.Min(64, tokenLen));
                else d = 1;
            }
            dims[i] = d;
        }
        return dims;
    }

    private static long[] BuildTokenIds(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return new long[] { 1, 2, 3, 4 };
        }

        var bytes = System.Text.Encoding.UTF8.GetBytes(text);
        var len = Math.Min(bytes.Length, 64);
        if (len == 0) len = 4;

        var tokens = new long[len];
        for (var i = 0; i < len; i++)
        {
            var b = i < bytes.Length ? bytes[i] : (byte)(i + 1);
            tokens[i] = (b % 255) + 1;
        }
        return tokens;
    }

    private static void FillLongTensor(DenseTensor<long> t, string name, long[] tokenIds)
    {
        var n = t.Length;
        if (name.Contains("input", StringComparison.OrdinalIgnoreCase) || name.Contains("token", StringComparison.OrdinalIgnoreCase) || name.Contains("id", StringComparison.OrdinalIgnoreCase))
        {
            for (var i = 0; i < n; i++)
            {
                t.Buffer.Span[i] = tokenIds[i % tokenIds.Length];
            }
            return;
        }

        for (var i = 0; i < n; i++)
        {
            t.Buffer.Span[i] = 0;
        }
    }

    private static void FillIntTensor(DenseTensor<int> t, string name, long[] tokenIds)
    {
        var n = t.Length;
        if (name.Contains("input", StringComparison.OrdinalIgnoreCase) || name.Contains("token", StringComparison.OrdinalIgnoreCase) || name.Contains("id", StringComparison.OrdinalIgnoreCase))
        {
            for (var i = 0; i < n; i++)
            {
                t.Buffer.Span[i] = (int)tokenIds[i % tokenIds.Length];
            }
            return;
        }

        for (var i = 0; i < n; i++)
        {
            t.Buffer.Span[i] = 0;
        }
    }

    private static void FillFloatTensor(DenseTensor<float> t, string name, long[] tokenIds)
    {
        var n = t.Length;
        if (name.Contains("mask", StringComparison.OrdinalIgnoreCase))
        {
            for (var i = 0; i < n; i++) t.Buffer.Span[i] = 1.0f;
            return;
        }

        for (var i = 0; i < n; i++)
        {
            t.Buffer.Span[i] = (float)((tokenIds[i % tokenIds.Length] % 17) / 16.0);
        }
    }

    private static void FillDoubleTensor(DenseTensor<double> t, string name, long[] tokenIds)
    {
        var n = t.Length;
        for (var i = 0; i < n; i++)
        {
            t.Buffer.Span[i] = ((tokenIds[i % tokenIds.Length] % 17) / 16.0);
        }
    }

    private static void FillBoolTensor(DenseTensor<bool> t, string name)
    {
        var n = t.Length;
        for (var i = 0; i < n; i++)
        {
            t.Buffer.Span[i] = !name.Contains("disable", StringComparison.OrdinalIgnoreCase);
        }
    }

    private static void ValidatePackFiles(string packRoot)
    {
        foreach (var file in RequiredModelFiles)
        {
            var p = Path.Combine(packRoot, file.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(p))
            {
                throw new FileNotFoundException($"ONNX model pack is incomplete. Missing: {file}", p);
            }
        }
    }

    private static string ResolvePackRoot(string modelCacheDir, string repoId)
    {
        var cache = ModelCachePath.ResolveAbsolute(modelCacheDir, RuntimePathResolver.AppRoot);
        var repoFolder = "models--" + repoId.Replace("/", "--");
        return Path.Combine(cache, "hf-cache", repoFolder);
    }

    private static bool IsOnnxRepo(string repoId)
    {
        return repoId.Contains("onnx", StringComparison.OrdinalIgnoreCase);
    }

    private static SessionOptions CreateSessionOptions(string device)
    {
        var sessionOptions = new SessionOptions
        {
            GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_EXTENDED
        };

        if (!string.Equals(device, "gpu", StringComparison.OrdinalIgnoreCase))
        {
            return sessionOptions;
        }

        try
        {
            sessionOptions.AppendExecutionProvider_CUDA(0);
            return sessionOptions;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ONNX CUDA provider unavailable, trying DML fallback: {ex.Message}");
        }

        try
        {
            if (OperatingSystem.IsWindows())
            {
                sessionOptions.AppendExecutionProvider_DML(0);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ONNX DML provider unavailable, using CPU provider: {ex.Message}");
        }

        return sessionOptions;
    }

    private static string NormalizePreferDevice(string? value)
    {
        return (value ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "gpu" => "gpu",
            "cuda" => "gpu",
            _ => "cpu"
        };
    }

    private static void Unload()
    {
        _conditionalDecoder?.Dispose();
        _embedTokens?.Dispose();
        _languageModel?.Dispose();
        _speechEncoder?.Dispose();

        _conditionalDecoder = null;
        _embedTokens = null;
        _languageModel = null;
        _speechEncoder = null;
        _lastPackKey = string.Empty;
    }
}
