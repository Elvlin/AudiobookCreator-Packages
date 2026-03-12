using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using App.Core.Runtime;

namespace App.Inference;

public sealed class LocalDllBackend : ITtsBackend
{
    public string Name => "local-dll";

    private readonly LocalInferenceOptions _options;

    public LocalDllBackend(LocalInferenceOptions? options = null)
    {
        _options = options ?? new LocalInferenceOptions();
    }

    public Task SynthesizeAsync(TtsRequest request, CancellationToken ct = default)
    {
        return Task.Run(() =>
        {
            if (string.IsNullOrWhiteSpace(request.OutputPath))
            {
                throw new ArgumentException("OutputPath is required.");
            }
            if (string.IsNullOrWhiteSpace(request.Text))
            {
                throw new ArgumentException("Input text is empty.");
            }
            if (string.IsNullOrWhiteSpace(request.VoicePath))
            {
                throw new ArgumentException("Voice path is required for local backend.");
            }
            if (!File.Exists(request.VoicePath))
            {
                throw new FileNotFoundException($"Voice file not found: {request.VoicePath}");
            }

            var outputDir = Path.GetDirectoryName(request.OutputPath);
            if (!string.IsNullOrWhiteSpace(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var modelCacheDir = ResolveModelCacheDir(_options.ModelCacheDir);
            var modelRepoId = string.IsNullOrWhiteSpace(_options.ModelRepoId) ? "onnx-community/chatterbox-ONNX" : _options.ModelRepoId.Trim();
            if (modelRepoId.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Qwen3-TTS native inference is not implemented in this build. " +
                    "Current native_tts_engine output is disabled because it produces synthetic/distorted audio.");
            }
            var runtimeSpeed = request.Speed;
            var pitchFactor = 1.0f;
            var energyFactor = 1.0f;
            var spectralTilt = 0.0f;
            float[]? envPitch = null;
            float[]? envEnergy = null;
            float[]? envTilt = null;
            if (_options.ValidateOnnxRuntimeSessions)
            {
                try
                {
                    OnnxModelRuntime.EnsureLoaded(modelCacheDir, modelRepoId, _options.PreferDevice);
                    var control = OnnxModelRuntime.ComputeProsodyControl(modelCacheDir, modelRepoId, request.Text, _options.PreferDevice);
                    runtimeSpeed = (float)Math.Clamp(request.Speed * control.SpeedFactor, 0.6, 1.6);
                    pitchFactor = (float)Math.Clamp(control.PitchFactor, 0.75, 1.30);
                    energyFactor = (float)Math.Clamp(control.EnergyFactor, 0.70, 1.35);
                    spectralTilt = (float)Math.Clamp(control.SpectralTilt, -0.30, 0.30);
                    var frames = Math.Clamp(request.Text.Length * 2, 32, 2048);
                    var env = OnnxModelRuntime.ComputeProsodyEnvelope(modelCacheDir, modelRepoId, request.Text, frames, _options.PreferDevice);
                    envPitch = env.Pitch;
                    envEnergy = env.Energy;
                    envTilt = env.Tilt;
                }
                catch (Exception ex)
                {
                    // Keep synthesis available even if ONNX prosody side data is broken.
                    Debug.WriteLine($"ONNX prosody runtime unavailable, falling back to baseline synthesis: {ex.Message}");
                    runtimeSpeed = request.Speed;
                    pitchFactor = 1.0f;
                    energyFactor = 1.0f;
                    spectralTilt = 0.0f;
                    envPitch = null;
                    envEnergy = null;
                    envTilt = null;
                }
            }

            var err = new StringBuilder(2048);
            int code;
            try
            {
                if (envPitch is not null && envEnergy is not null && envTilt is not null)
                {
                    code = SynthesizeToWavWithEnvelopeUtf8(
                        request.Text,
                        request.VoicePath,
                        request.OutputPath,
                        modelCacheDir,
                        modelRepoId,
                        runtimeSpeed,
                        pitchFactor,
                        energyFactor,
                        spectralTilt,
                        envPitch,
                        envEnergy,
                        envTilt,
                        envPitch.Length,
                        err,
                        err.Capacity);
                }
                else
                {
                    code = SynthesizeToWavWithFeaturesUtf8(
                        request.Text,
                        request.VoicePath,
                        request.OutputPath,
                        modelCacheDir,
                        modelRepoId,
                        runtimeSpeed,
                        pitchFactor,
                        energyFactor,
                        spectralTilt,
                        err,
                        err.Capacity);
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    code = SynthesizeToWavWithFeaturesUtf8(
                        request.Text,
                        request.VoicePath,
                        request.OutputPath,
                        modelCacheDir,
                        modelRepoId,
                        runtimeSpeed,
                        pitchFactor,
                        energyFactor,
                        spectralTilt,
                        err,
                        err.Capacity);
                }
                catch (EntryPointNotFoundException)
                {
                    code = SynthesizeToWavUtf8(
                        request.Text,
                        request.VoicePath,
                        request.OutputPath,
                        modelCacheDir,
                        modelRepoId,
                        runtimeSpeed,
                        err,
                        err.Capacity);
                }
            }
            catch (DllNotFoundException)
            {
                code = SynthesizeToWavUtf8(
                    request.Text,
                    request.VoicePath,
                    request.OutputPath,
                    modelCacheDir,
                    modelRepoId,
                    runtimeSpeed,
                    err,
                    err.Capacity);
            }

            if (code != 0)
            {
                var msg = err.Length > 0 ? err.ToString() : $"native_tts_engine failed with code {code}";
                throw new InvalidOperationException(msg);
            }
        }, ct);
    }

    private static string ResolveModelCacheDir(string modelCacheDir)
    {
        return ModelCachePath.ResolveAbsolute(modelCacheDir, RuntimePathResolver.AppRoot);
    }

    [DllImport("native_tts_engine.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "synthesize_to_wav_utf8")]
    private static extern int SynthesizeToWavUtf8(
        [MarshalAs(UnmanagedType.LPUTF8Str)] string text,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string voicePath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string outputPath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelCacheDir,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelRepoId,
        float speed,
        StringBuilder errorBuffer,
        int errorBufferLength);

    [DllImport("native_tts_engine.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "synthesize_to_wav_with_features_utf8")]
    private static extern int SynthesizeToWavWithFeaturesUtf8(
        [MarshalAs(UnmanagedType.LPUTF8Str)] string text,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string voicePath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string outputPath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelCacheDir,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelRepoId,
        float speed,
        float pitchFactor,
        float energyFactor,
        float spectralTilt,
        StringBuilder errorBuffer,
        int errorBufferLength);

    [DllImport("native_tts_engine.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "synthesize_to_wav_with_envelope_utf8")]
    private static extern int SynthesizeToWavWithEnvelopeUtf8(
        [MarshalAs(UnmanagedType.LPUTF8Str)] string text,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string voicePath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string outputPath,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelCacheDir,
        [MarshalAs(UnmanagedType.LPUTF8Str)] string modelRepoId,
        float speed,
        float pitchFactor,
        float energyFactor,
        float spectralTilt,
        [In] float[] pitchEnv,
        [In] float[] energyEnv,
        [In] float[] tiltEnv,
        int envLength,
        StringBuilder errorBuffer,
        int errorBufferLength);
}
