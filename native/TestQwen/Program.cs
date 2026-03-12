using System;
using System.IO;
using System.Text;
using App.Core.Runtime;
using App.Inference;

// Run from customer_release so RuntimePathResolver.AppRoot points to customer_release.
// Usage: cd customer_release && TestQwen.exe
// Or: dotnet run --project native\TestQwen -c Release (then cd customer_release && TestQwen.exe)

var appRoot = RuntimePathResolver.AppRoot;
var voicesDir = Path.Combine(appRoot, "voices");
var modelsDir = Path.Combine(appRoot, "models");
Directory.CreateDirectory(voicesDir);

var voicePath = Environment.GetEnvironmentVariable("QWEN_TEST_VOICE");
if (string.IsNullOrWhiteSpace(voicePath))
    voicePath = Path.Combine(voicesDir, "Callum.wav");
if (!Path.IsPathRooted(voicePath))
    voicePath = Path.Combine(appRoot, voicePath);

if (!File.Exists(voicePath))
{
    var fallbackVoice = Path.Combine(voicesDir, "test_qwen_voice.wav");
    voicePath = fallbackVoice;
    WriteMinimalWav(voicePath, sampleRate: 24000, durationSeconds: 1f);
    Console.WriteLine($"Created test voice: {voicePath}");
}

var refTextPath = Path.Combine(Path.GetDirectoryName(voicePath) ?? voicesDir, Path.GetFileNameWithoutExtension(voicePath) + ".ref.txt");
var refText = File.Exists(refTextPath) ? File.ReadAllText(refTextPath).Trim() : string.Empty;

var outputPath = Path.Combine(appRoot, "test_qwen_output.wav");
var testText = Environment.GetEnvironmentVariable("QWEN_TEST_TEXT");
if (string.IsNullOrWhiteSpace(testText))
    testText = "Hello. This is a short test.";
var request = new TtsRequest
{
    Text = testText,
    VoicePath = voicePath,
    RefText = refText,
    OutputPath = outputPath,
    Speed = 1.0f
};

// Prefer env so we can run from TestQwenPublish and point at customer_release\models
var modelCacheDir = Environment.GetEnvironmentVariable("QWEN_MODEL_CACHE");
if (string.IsNullOrWhiteSpace(modelCacheDir))
    modelCacheDir = "models";

var options = new LocalInferenceOptions
{
    ModelCacheDir = modelCacheDir,
    ModelRepoId = "xkos/Qwen3-TTS-12Hz-1.7B-ONNX",
    PreferDevice = Environment.GetEnvironmentVariable("QWEN_DEVICE") ?? "cpu"
};

Console.WriteLine($"AppRoot: {appRoot}");
Console.WriteLine($"Model cache: {Path.Combine(appRoot, "models")}");
Console.WriteLine($"Device: {options.PreferDevice}");
Console.WriteLine($"Voice: {voicePath}");
Console.WriteLine($"RefText: {(string.IsNullOrWhiteSpace(refText) ? "NO" : $"YES ({refText.Length} chars)")}");
Console.WriteLine($"Text: {request.Text}");
Console.WriteLine("Starting Qwen Split backend and one synthesis...");

try
{
    var backendKind = (Environment.GetEnvironmentVariable("QWEN_TEST_BACKEND") ?? "split").Trim().ToLowerInvariant();
    ITtsBackend backend = backendKind switch
    {
        "python" => new QwenPythonBackend(options),
        "dll" => new Qwen3OnnxDllBackend(options),
        _ => new Qwen3OnnxSplitBackend(options)
    };
    using var _ = backend as IDisposable;
    Console.WriteLine($"Backend: {backend.Name}");
    await backend.SynthesizeAsync(request);
    var exists = File.Exists(outputPath);
    var size = exists ? new FileInfo(outputPath).Length : 0;
    if (exists && TryInspectWav(outputPath, out var info))
    {
        Console.WriteLine($"WAV: sr={info.SampleRate} ch={info.Channels} bits={info.BitsPerSample} dur={info.DurationSeconds:F2}s rms={info.Rms:F4}");
    }
    Console.WriteLine(exists
        ? $"OK. Output: {outputPath} ({size} bytes)"
        : $"FAIL. Output file not created: {outputPath}");
    Environment.Exit(exists ? 0 : 1);
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
    Environment.Exit(1);
}

static void WriteMinimalWav(string path, int sampleRate, float durationSeconds)
{
    var numSamples = (int)(sampleRate * durationSeconds);
    var dataSize = numSamples * 2; // 16-bit
    using var fs = File.Create(path);
    using var bw = new BinaryWriter(fs);
    var byteRate = sampleRate * 2;
    bw.Write(Encoding.ASCII.GetBytes("RIFF"));
    bw.Write(36 + dataSize);
    bw.Write(Encoding.ASCII.GetBytes("WAVEfmt "));
    bw.Write(16);
    bw.Write((short)1);
    bw.Write((short)1);
    bw.Write(sampleRate);
    bw.Write(byteRate);
    bw.Write((short)2);
    bw.Write((short)16);
    bw.Write(Encoding.ASCII.GetBytes("data"));
    bw.Write(dataSize);
    for (var i = 0; i < numSamples; i++)
        bw.Write((short)0);
}

static bool TryInspectWav(string path, out (ushort Channels, uint SampleRate, ushort BitsPerSample, double DurationSeconds, double Rms) info)
{
    info = default;
    if (!File.Exists(path))
        return false;

    using var fs = File.OpenRead(path);
    using var br = new BinaryReader(fs);
    if (new string(br.ReadChars(4)) != "RIFF")
        return false;
    _ = br.ReadInt32();
    if (new string(br.ReadChars(4)) != "WAVE")
        return false;

    ushort channels = 0, bits = 0, fmtTag = 0;
    uint sampleRate = 0;
    byte[]? data = null;
    while (fs.Position + 8 <= fs.Length)
    {
        var id = new string(br.ReadChars(4));
        var size = br.ReadInt32();
        if (size < 0 || fs.Position + size > fs.Length) return false;
        if (id == "fmt ")
        {
            fmtTag = br.ReadUInt16();
            channels = br.ReadUInt16();
            sampleRate = br.ReadUInt32();
            _ = br.ReadUInt32();
            _ = br.ReadUInt16();
            bits = br.ReadUInt16();
            var remaining = size - 16;
            if (remaining > 0) br.ReadBytes(remaining);
        }
        else if (id == "data")
        {
            data = br.ReadBytes(size);
        }
        else
        {
            br.ReadBytes(size);
        }
        if ((size & 1) == 1 && fs.Position < fs.Length) fs.Position++;
    }
    if (data is null || data.Length == 0 || sampleRate == 0 || channels == 0 || bits == 0)
        return false;
    var dur = bits switch
    {
        16 => data.Length / (double)(sampleRate * channels * 2),
        32 => data.Length / (double)(sampleRate * channels * 4),
        _ => 0
    };
    double rms = 0;
    if (fmtTag == 1 && bits == 16)
    {
        long n = data.Length / 2;
        if (n > 0)
        {
            double sumSq = 0;
            for (int i = 0; i + 1 < data.Length; i += 2)
            {
                short s = (short)(data[i] | (data[i + 1] << 8));
                var x = s / 32768.0;
                sumSq += x * x;
            }
            rms = Math.Sqrt(sumSq / n);
        }
    }
    info = (channels, sampleRate, bits, dur, rms);
    return true;
}
