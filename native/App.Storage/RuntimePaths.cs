using App.Core.Runtime;

namespace App.Storage;

public static class RuntimePaths
{
    public static string AppRoot => RuntimePathResolver.AppRoot;

    public static string ConfigPath => Path.Combine(AppRoot, "defaults", "app_config.json");

    public static string InstallStatePath => Path.Combine(AppRoot, "defaults", "install_state.json");

    public static string SetupRequiredFlagPath => Path.Combine(AppRoot, "defaults", "setup_required.flag");

    public static string ProjectsDir => Path.Combine(AppRoot, "projects");
}
