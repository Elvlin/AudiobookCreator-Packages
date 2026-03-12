namespace App.Core.Runtime;

public static class ModelCachePath
{
    public static string NormalizeInput(string? modelCacheDir)
    {
        var value = string.IsNullOrWhiteSpace(modelCacheDir) ? "models" : modelCacheDir.Trim();
        value = value.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (string.IsNullOrWhiteSpace(value))
        {
            return "models";
        }

        var name = Path.GetFileName(value);
        if (!string.Equals(name, "hf-cache", StringComparison.OrdinalIgnoreCase))
        {
            return value;
        }

        var parent = Path.GetDirectoryName(value);
        return string.IsNullOrWhiteSpace(parent) ? "models" : parent;
    }

    public static string ResolveAbsolute(string? modelCacheDir, string appRoot)
    {
        var normalized = NormalizeInput(modelCacheDir);
        if (Path.IsPathRooted(normalized))
        {
            return normalized;
        }

        return Path.Combine(appRoot, normalized);
    }
}
