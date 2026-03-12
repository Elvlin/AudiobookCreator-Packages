using System.Reflection;
using App.Core.Runtime;

namespace App.Inference;

internal static class QwenEmbeddedRuntime
{
    private const string ResourceName = "App.Inference.Resources.qwen3_tts_rust.dll";

    public static string EnsureBundledRustDllAtAppRoot()
    {
        var appDll = Path.Combine(RuntimePathResolver.AppRoot, "qwen3_tts_rust.dll");
        if (File.Exists(appDll))
        {
            return appDll;
        }

        var asm = typeof(QwenEmbeddedRuntime).Assembly;
        using var stream = asm.GetManifestResourceStream(ResourceName);
        if (stream is null)
        {
            return appDll;
        }

        var dir = Path.GetDirectoryName(appDll);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using var fs = File.Create(appDll);
        stream.CopyTo(fs);
        return appDll;
    }
}
