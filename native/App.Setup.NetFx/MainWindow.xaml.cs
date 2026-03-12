using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using AudiobookCreator.SetupNetFx.Models;
using AudiobookCreator.SetupNetFx.Services;
using Controls = System.Windows.Controls;
using Forms = System.Windows.Forms;

namespace AudiobookCreator.SetupNetFx;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly SetupConfig _setupConfig = new();
    private readonly HttpClient _httpClient = new();
    private readonly EnvironmentService _environmentService = new();
    private readonly PackagePlanner _planner = new();
    private readonly InstallStateService _installStateService = new();
    private readonly AppConfigWriter _appConfigWriter = new();
    private readonly ManifestService _manifestService;
    private readonly DownloadInstallerService _downloadInstallerService;
    private readonly string _logPath;

    private PackageManifest? _manifest;
    private DetectedMachineProfile? _machine;
    private ResolvedInstallPlan? _plan;
    private readonly SetupSelection _selection = new();
    private int _stepIndex;
    private bool _installStarted;

    public MainWindow()
    {
        InitializeComponent();
        _manifestService = new ManifestService(_httpClient, _setupConfig);
        _downloadInstallerService = new DownloadInstallerService(_httpClient, _setupConfig);
        _logPath = InitializeLogPath();

        LicenseTextBox.Text = EulaText.Get();
        AutoModeRadio.IsChecked = true;
        ApiProviderCombo.ItemsSource = new[] { "openai", "alibaba" };
        ApiProviderCombo.SelectedIndex = 0;
        ChatterboxBackendCombo.ItemsSource = new[] { "none", "onnx", "python" };
        KittenBackendCombo.ItemsSource = new[] { "none", "onnx", "python" };
        ChatterboxBackendCombo.SelectedIndex = 0;
        KittenBackendCombo.SelectedIndex = 0;
        _selection.InstallDirectory = _setupConfig.DefaultInstallDir;
        InstallDirectoryTextBox.Text = _selection.InstallDirectory;

        Loaded += OnLoaded;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        UpdateStepView();
        await LoadManifestAndScanAsync();
    }

    private async Task LoadManifestAndScanAsync()
    {
        try
        {
            ScanStatusText.Text = "Loading package feed and scanning your PC...";
            ApplySelectionFromUi();
            _manifest = await _manifestService.LoadManifestAsync(CancellationToken.None);
            var existingState = await _installStateService.LoadAsync(_selection.InstallDirectory, CancellationToken.None);
            var releaseFeedReachable = await _environmentService.CanReachReleaseFeedAsync(_setupConfig, CancellationToken.None);
            _machine = _environmentService.Detect(_selection.InstallDirectory, existingState, releaseFeedReachable);
            ScanDetailsText.Text = BuildScanDetails();
            ScanStatusText.Text = "PC scan complete.";
            RecalculatePlan();
        }
        catch (Exception ex)
        {
            ScanStatusText.Text = "Setup could not load the package feed.";
            ScanDetailsText.Text = ex.Message;
            AppendLog($"Manifest/scan error: {ex.Message}");
        }
    }

    private string BuildScanDetails()
    {
        if (_machine is null)
        {
            return "Scan data unavailable.";
        }

        var sb = new StringBuilder();
        sb.AppendLine($"OS: {_machine.OperatingSystem}");
        sb.AppendLine($"64-bit Windows: {_machine.Is64Bit}");
        sb.AppendLine($".NET Desktop Runtime: {_machine.DotNetStatus}");
        sb.AppendLine($"CPU threads: {_machine.CpuLogical}");
        sb.AppendLine($"RAM: {_machine.RamGb:F1} GB");
        sb.AppendLine($"GPU: {_machine.GpuName}");
        sb.AppendLine($"GPU vendor: {_machine.GpuVendor}");
        sb.AppendLine($"GPU VRAM: {_machine.GpuVramGb:F1} GB");
        sb.AppendLine($"Internet: {_machine.InternetStatus}");
        sb.AppendLine($"Disk: {_machine.DiskStatus}");
        sb.AppendLine($"Install state: {_machine.ExistingInstallStatus}");
        return sb.ToString();
    }

    private void RecalculatePlan()
    {
        if (_manifest is null || _machine is null)
        {
            return;
        }

        ApplySelectionFromUi();
        _plan = _planner.Resolve(_manifest, _selection, _machine);
        PlanSummaryText.Text = _plan.Summary;
        PlanWarningsText.Text = _plan.Warnings.Count == 0
            ? "No setup warnings."
            : string.Join(Environment.NewLine, _plan.Warnings.Select(x => $"* {x}"));
        SidebarPlanText.Text = BuildSidebarPlan();
        UpdateCustomWarnings();
        OnPropertyChanged(nameof(_plan));
    }

    private void ApplySelectionFromUi()
    {
        _selection.InstallMode = AutoModeRadio.IsChecked == true ? "Auto" : "Custom";
        _selection.ApiOnly = ApiOnlyCheckBox.IsChecked == true;
        _selection.ChatterboxBackend = GetComboValue(ChatterboxBackendCombo, AutoModeRadio.IsChecked == true ? "auto" : "none");
        _selection.KittenBackend = GetComboValue(KittenBackendCombo, AutoModeRadio.IsChecked == true ? "auto" : "none");
        _selection.IncludeQwen = QwenCheckBox.IsChecked == true;
        _selection.IncludeLocalLlm = LlmCheckBox.IsChecked == true;
        _selection.EnterApiKeyNow = EnterApiKeyCheckBox.IsChecked == true;
        _selection.ApiProvider = (ApiProviderCombo.SelectedItem?.ToString() ?? "openai").Trim();
        _selection.ApiKey = ApiKeyPasswordBox.Password ?? string.Empty;
        _selection.InstallDirectory = InstallDirectoryTextBox.Text.Trim();
    }

    private static string GetComboValue(Controls.ComboBox comboBox, string fallback)
    {
        return comboBox.SelectedItem?.ToString() ?? fallback;
    }

    private string BuildSidebarPlan()
    {
        if (_plan is null)
        {
            return "No package plan resolved yet.";
        }

        var downloadGb = _plan.DownloadBytes / 1024d / 1024d / 1024d;
        var installGb = _plan.FinalInstallBytes / 1024d / 1024d / 1024d;
        return $"{_plan.PackageIds.Count} packages\n{downloadGb:F2} GB download\n{installGb:F2} GB installed\n{string.Join(", ", _plan.PackageIds.Select(GetFriendlyPackageName))}";
    }

    private void UpdateCustomWarnings()
    {
        ChatterboxWarningText.Text = string.Empty;
        KittenWarningText.Text = string.Empty;
        QwenWarningText.Text = string.Empty;
        LlmWarningText.Text = string.Empty;

        if (_machine is null)
        {
            return;
        }

        if (GetComboValue(ChatterboxBackendCombo, "none") == "onnx" &&
            !string.Equals(_machine.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase))
        {
            ChatterboxWarningText.Text = "Likely not suitable on this PC. NVIDIA is the recommended ONNX path.";
        }
        else if (GetComboValue(ChatterboxBackendCombo, "none") == "python" && _machine.RamGb < 12)
        {
            ChatterboxWarningText.Text = "Likely not suitable on this PC. Recommended RAM is 12 GB or higher.";
        }

        if (GetComboValue(KittenBackendCombo, "none") == "onnx" &&
            !string.Equals(_machine.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase))
        {
            KittenWarningText.Text = "Likely not suitable on this PC. NVIDIA is the recommended ONNX path.";
        }
        else if (GetComboValue(KittenBackendCombo, "none") == "python" && _machine.RamGb < 8)
        {
            KittenWarningText.Text = "Likely not suitable on this PC. Recommended RAM is 8 GB or higher.";
        }

        if (QwenCheckBox.IsChecked == true && _machine.RamGb < 24)
        {
            QwenWarningText.Text = "Likely not suitable on this PC. Qwen is recommended on systems with 24 GB RAM or higher.";
        }

        if (LlmCheckBox.IsChecked == true && _machine.RamGb < 16)
        {
            LlmWarningText.Text = "Likely not suitable on this PC. Local LLM prep is recommended on systems with 16 GB RAM or higher.";
        }
    }

    private void UpdateStepView()
    {
        WelcomePanel.Visibility = _stepIndex == 0 ? Visibility.Visible : Visibility.Collapsed;
        LicensePanel.Visibility = _stepIndex == 1 ? Visibility.Visible : Visibility.Collapsed;
        ScanPanel.Visibility = _stepIndex == 2 ? Visibility.Visible : Visibility.Collapsed;
        ModePanel.Visibility = _stepIndex == 3 ? Visibility.Visible : Visibility.Collapsed;
        ApiPanel.Visibility = _stepIndex == 4 ? Visibility.Visible : Visibility.Collapsed;
        FolderPanel.Visibility = _stepIndex == 5 ? Visibility.Visible : Visibility.Collapsed;
        InstallPanel.Visibility = _stepIndex == 6 ? Visibility.Visible : Visibility.Collapsed;
        FinishPanel.Visibility = _stepIndex == 7 ? Visibility.Visible : Visibility.Collapsed;

        CustomOptionsBorder.IsEnabled = CustomModeRadio.IsChecked == true && ApiOnlyCheckBox.IsChecked != true;
        ApiEntryPanel.IsEnabled = EnterApiKeyCheckBox.IsChecked == true;
        BackButton.IsEnabled = _stepIndex > 0 && _stepIndex < 7;
        NextButton.IsEnabled = CanAdvanceFromCurrentStep();
        NextButton.Content = _stepIndex switch
        {
            6 => _installStarted ? "Installing..." : "Install",
            7 => "Close",
            _ => "Next"
        };

        InstallFolderStatusText.Text = BuildInstallFolderStatus();
        UpdateStepIndicator();
    }

    private bool CanAdvanceFromCurrentStep()
    {
        return _stepIndex switch
        {
            1 => AcceptLicenseCheckBox.IsChecked == true,
            6 => !_installStarted,
            _ => true
        };
    }

    private string BuildInstallFolderStatus()
    {
        try
        {
            var path = InstallDirectoryTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                return "Choose an install folder.";
            }

            var fullPath = Path.GetFullPath(path);
            var root = Path.GetPathRoot(fullPath);
            if (string.IsNullOrWhiteSpace(root))
            {
                return "Install path is not valid.";
            }

            var drive = new DriveInfo(root);
            var freeGb = drive.AvailableFreeSpace / 1024d / 1024d / 1024d;
            var existingStatePath = Path.Combine(fullPath, "defaults", "install_state.json");
            var existingState = File.Exists(existingStatePath)
                ? $"{Environment.NewLine}Existing Audiobook Creator install detected."
                : string.Empty;
            var requirement = _plan is null
                ? string.Empty
                : $"{Environment.NewLine}Required working space: {_plan.WorkingBytesRequired / 1024d / 1024d / 1024d:F2} GB";
            return $"Install path: {fullPath}{Environment.NewLine}Target drive free space: {freeGb:F1} GB{requirement}{existingState}";
        }
        catch
        {
            return "Install path is not valid.";
        }
    }

    private void UpdateStepIndicator()
    {
        StepIndicatorText.Text = $"Step {_stepIndex + 1} of 8";
        var steps = new[]
        {
            StepWelcomeText, StepLicenseText, StepScanText, StepModeText,
            StepApiText, StepFolderText, StepInstallText, StepFinishText
        };

        for (var i = 0; i < steps.Length; i++)
        {
            steps[i].Foreground = i == _stepIndex
                ? System.Windows.Media.Brushes.White
                : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(168, 195, 180));
            steps[i].FontWeight = i == _stepIndex ? FontWeights.SemiBold : FontWeights.Normal;
        }
    }

    private void AppendLog(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        var line = $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}";
        InstallLogTextBox.AppendText(line);
        InstallLogTextBox.ScrollToEnd();
        File.AppendAllText(_logPath, line);
    }

    private async void NextButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_stepIndex == 7)
        {
            if (LaunchAfterSetupCheckBox.IsChecked == true)
            {
                LaunchInstalledApp();
            }

            Close();
            return;
        }

        if (_stepIndex == 2)
        {
            await LoadManifestAndScanAsync();
        }

        if (_stepIndex == 6)
        {
            await RunInstallAsync();
            return;
        }

        _stepIndex++;
        UpdateStepView();
    }

    private void BackButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_stepIndex <= 0 || _stepIndex >= 7)
        {
            return;
        }

        _stepIndex--;
        UpdateStepView();
    }

    private async Task RunInstallAsync()
    {
        if (_manifest is null || _plan is null)
        {
            AppendLog("Cannot install because the manifest or install plan is missing.");
            return;
        }

        if (_machine is not null && _machine.AvailableInstallDriveBytes < _plan.WorkingBytesRequired)
        {
            AppendLog("Install blocked: not enough free disk space for the selected packages.");
            InstallSummaryText.Text = "Not enough free disk space for staging, extraction, and final install files.";
            return;
        }

        _installStarted = true;
        UpdateStepView();
        InstallProgressBar.Value = 0;
        InstallLogTextBox.Clear();
        InstallSummaryText.Text =
            $"Installing {_plan.PackageIds.Count} packages to {_selection.InstallDirectory}{Environment.NewLine}" +
            $"Download: {_plan.DownloadBytes / 1024d / 1024d / 1024d:F2} GB | Installed: {_plan.FinalInstallBytes / 1024d / 1024d / 1024d:F2} GB";
        AppendLog("Preparing installation...");

        try
        {
            var downloadDir = Path.Combine(Path.GetTempPath(), "AudiobookCreatorSetup", "runtime");
            var dotnetReady = await _environmentService.EnsureDotNetDesktopRuntimeAsync(_setupConfig, downloadDir, CancellationToken.None);
            if (!dotnetReady)
            {
                throw new InvalidOperationException("Failed to install .NET Desktop Runtime 7.x.");
            }

            Directory.CreateDirectory(_selection.InstallDirectory);
            await _installStateService.SaveAsync(new InstalledState
            {
                InstallPath = _selection.InstallDirectory,
                CompletionState = "installing",
                Packages = _plan.PackageIds.Select(id =>
                {
                    var pkg = _manifest.Packages.First(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
                    return new InstalledPackage { Id = pkg.Id, Version = pkg.Version };
                }).ToList()
            }, _selection.InstallDirectory, CancellationToken.None);

            var progress = new Progress<InstallProgress>(p =>
            {
                InstallProgressBar.Value = p.Percent;
                InstallSummaryText.Text = p.Message;
                AppendLog(p.Message);
            });

            await _downloadInstallerService.InstallAsync(_manifest, _plan, _selection.InstallDirectory, progress, CancellationToken.None);
            await _appConfigWriter.ApplyAsync(_selection.InstallDirectory, _selection, CancellationToken.None);
            await _installStateService.SaveAsync(new InstalledState
            {
                InstallPath = _selection.InstallDirectory,
                CompletionState = "installed",
                Packages = _plan.PackageIds.Select(id =>
                {
                    var pkg = _manifest.Packages.First(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
                    return new InstalledPackage { Id = pkg.Id, Version = pkg.Version };
                }).ToList()
            }, _selection.InstallDirectory, CancellationToken.None);

            FinishSummaryText.Text =
                $"Audiobook Creator was installed to:{Environment.NewLine}{_selection.InstallDirectory}{Environment.NewLine}{Environment.NewLine}Installed components:{Environment.NewLine}{string.Join(Environment.NewLine, _plan.PackageIds.Select(GetFriendlyPackageName))}";
            _installStarted = false;
            _stepIndex = 7;
            UpdateStepView();
        }
        catch (Exception ex)
        {
            AppendLog($"Install failed: {ex.Message}");
            InstallSummaryText.Text = "Installation failed. Review the log, then retry.";
            if (_plan is not null)
            {
                await _installStateService.SaveAsync(new InstalledState
                {
                    InstallPath = _selection.InstallDirectory,
                    CompletionState = "failed",
                    Packages = _plan.PackageIds.Select(id =>
                    {
                        var pkg = _manifest!.Packages.First(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
                        return new InstalledPackage { Id = pkg.Id, Version = pkg.Version };
                    }).ToList()
                }, _selection.InstallDirectory, CancellationToken.None);
            }
            _installStarted = false;
            UpdateStepView();
        }
    }

    private void LaunchInstalledApp()
    {
        var exePath = Path.Combine(_selection.InstallDirectory, "AudiobookCreator.exe");
        if (!File.Exists(exePath))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = exePath,
            WorkingDirectory = _selection.InstallDirectory,
            UseShellExecute = true
        });
    }

    private async void ScanAgainButton_OnClick(object sender, RoutedEventArgs e)
    {
        await LoadManifestAndScanAsync();
    }

    private void AcceptLicenseCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        UpdateStepView();
    }

    private void ModeInputs_OnChanged(object sender, RoutedEventArgs e)
    {
        RecalculatePlan();
        UpdateStepView();
    }

    private void ApiInputs_OnChanged(object sender, RoutedEventArgs e)
    {
        ApplySelectionFromUi();
        UpdateStepView();
    }

    private void ApiKeyPasswordBox_OnChanged(object sender, RoutedEventArgs e)
    {
        ApplySelectionFromUi();
        UpdateStepView();
    }

    private void InstallDirectoryTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        ApplySelectionFromUi();
        if (_machine is not null)
        {
            _machine = _environmentService.Detect(_selection.InstallDirectory, null, _machine.InternetAvailable);
            ScanDetailsText.Text = BuildScanDetails();
        }

        RecalculatePlan();
        UpdateStepView();
    }

    private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Choose where Audiobook Creator should be installed.",
            SelectedPath = InstallDirectoryTextBox.Text
        };

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            InstallDirectoryTextBox.Text = dialog.SelectedPath;
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        _httpClient.Dispose();
        base.OnClosing(e);
    }

    private string InitializeLogPath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            $"{_setupConfig.AppId}Setup",
            "logs");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, $"setup-{DateTime.Now:yyyyMMdd-HHmmss}.log");
    }

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static string GetFriendlyPackageName(string packageId)
    {
        return packageId switch
        {
            "app_core_no_audiox" => "Audiobook Creator core app",
            "onnx_runtime_win_x64" => "ONNX runtime",
            "python_qwen" => "Qwen Python runtime",
            "python_chatterbox" => "Chatterbox Python runtime",
            "python_kitten" => "Kitten Python runtime",
            "tools_basic" => "Basic tools",
            "tools_llm" => "Local LLM tools",
            "tools_qwen_worker" => "Qwen worker tools",
            "model_chatterbox_onnx" => "Chatterbox ONNX model",
            "model_chatterbox_python" => "Chatterbox Python model",
            "model_kitten" => "Kitten model",
            "model_qwen_onnx" => "Qwen ONNX model",
            "model_qwen_dll" => "Qwen support files",
            "model_qwen_python_cache" => "Qwen Python cache",
            "model_llm" => "Local LLM models",
            _ => packageId.Replace('_', ' ')
        };
    }
}
