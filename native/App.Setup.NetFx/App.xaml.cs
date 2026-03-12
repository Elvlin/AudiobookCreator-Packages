using System.Threading;

namespace AudiobookCreator.SetupNetFx;

public partial class App : System.Windows.Application
{
    private Mutex? _singleInstanceMutex;

    protected override void OnStartup(System.Windows.StartupEventArgs e)
    {
        _singleInstanceMutex = new Mutex(true, @"Local\AudiobookCreatorSetup", out var createdNew);
        if (!createdNew)
        {
            System.Windows.MessageBox.Show(
                "Audiobook Creator Setup is already running.",
                "Audiobook Creator Setup",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    protected override void OnExit(System.Windows.ExitEventArgs e)
    {
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        base.OnExit(e);
    }
}
