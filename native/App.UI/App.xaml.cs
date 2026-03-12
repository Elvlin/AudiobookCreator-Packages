using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using App.Core.Runtime;

namespace AudiobookCreator.UI;

public partial class App : Application
{
    private static string CrashLogPath => Path.Combine(RuntimePathResolver.AppRoot, "crash_runtime.log");
    private static bool _showingUnhandledError;

    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        ShutdownMode = ShutdownMode.OnMainWindowClose;

        base.OnStartup(e);

        // Dev mode override: skip startup integrity/SHA enforcement while the app is under active changes.
        // Set ABC_DISABLE_INTEGRITY=0 to re-enable checks without code edits.
        var disableIntegrityCheckForNow =
            !string.Equals(Environment.GetEnvironmentVariable("ABC_DISABLE_INTEGRITY"), "0", StringComparison.Ordinal);
        if (!disableIntegrityCheckForNow)
        {
#if DEBUG
            const bool strictIntegrity = false;
#else
            const bool strictIntegrity = true;
#endif
            var check = ReleaseIntegrityVerifier.VerifyForStartup(RuntimePathResolver.AppRoot, strictIntegrity);
            if (!check.Success)
            {
                MessageBox.Show(
                    check.Message + Environment.NewLine + Environment.NewLine +
                    "Integrity validation failed. App will now close.",
                    "Integrity Check Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Shutdown(-1);
                return;
            }
        }

        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppendCrashLog("DispatcherUnhandledException", e.Exception);
        e.Handled = true;
        if (_showingUnhandledError)
        {
            return;
        }

        try
        {
            _showingUnhandledError = true;
            MessageBox.Show(
                "A runtime error occurred and was caught to prevent a full app crash.\n" +
                "Current action may have been interrupted.\n\n" +
                $"Details: {e.Exception.Message}",
                "Runtime Error (Recovered)",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _showingUnhandledError = false;
        }
    }

    private void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            AppendCrashLog("AppDomainUnhandledException", ex);
            return;
        }

        AppendCrashLogRaw($"[{DateTime.UtcNow:O}] AppDomainUnhandledException: non-exception object. IsTerminating={e.IsTerminating}");
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        AppendCrashLog("UnobservedTaskException", e.Exception);
    }

    private static void AppendCrashLog(string source, Exception ex)
    {
        var text =
            $"[{DateTime.UtcNow:O}] {source}{Environment.NewLine}" +
            $"{ex.GetType().FullName}: {ex.Message}{Environment.NewLine}" +
            $"{ex.StackTrace}{Environment.NewLine}{Environment.NewLine}";
        AppendCrashLogRaw(text);
    }

    private static void AppendCrashLogRaw(string text)
    {
        try
        {
            File.AppendAllText(CrashLogPath, text + Environment.NewLine);
        }
        catch
        {
            // Never crash the app while trying to log crashes.
        }
    }
}
