using System;
using System.IO;
using App.Core.Runtime;

namespace AudiobookCreator.UI;

internal static class AppFeatureFlags
{
    private static readonly Lazy<bool> CachedAudioEnhancementEnabled = new(() =>
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var disabledMarker = Path.Combine(appRoot, "no_audiox.flag");
        return !File.Exists(disabledMarker);
    });

    public static bool AudioEnhancementEnabled => CachedAudioEnhancementEnabled.Value;
}
