namespace App.Core.Runtime;

public static class RuntimePathResolver
{
    private static readonly Lazy<string> CachedAppRoot = new(ResolveAppRoot);

    public static string AppRoot => CachedAppRoot.Value;

    private static string ResolveAppRoot()
    {
        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath))
        {
            var processDir = Path.GetDirectoryName(processPath);
            if (!string.IsNullOrWhiteSpace(processDir))
            {
                return processDir;
            }
        }

        return AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
