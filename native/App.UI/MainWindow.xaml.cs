using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Data;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using App.Core.Models;
using App.Core.Runtime;
using App.Diagnostics;
using App.Inference;
using App.Storage;
using AudiobookCreator.UI.SetupFlow.Services;
using Microsoft.Win32;
using NAudio.Wave;

namespace AudiobookCreator.UI;

public partial class MainWindow : Window
{
    private readonly JsonConfigStore _configStore = new();
    private readonly JsonProjectStore _projectStore = new();
    private readonly HuggingFaceModelDownloader _modelDownloader = new();
    private readonly InstallStateService _installStateService = new();

    private AppConfig _config = new();
    private ProjectDocument _project = new();
    private string? _currentProjectPath;

    private readonly ObservableCollection<InputFileEntry> _inputFiles = new();
    private readonly ObservableCollection<QueueRow> _queueRows = new();
    private readonly ObservableCollection<AudioEnhanceInputFileEntry> _enhanceInputFiles = new();
    private readonly ObservableCollection<QueueRow> _enhanceQueueRows = new();
    private readonly ObservableCollection<HistoryRow> _historyRows = new();
    private readonly List<string?> _projectSelectorPaths = new();
    private static readonly object GenerateLogLock = new();
    private bool _isGenerating;
    private bool _isRefreshingProjectSelector;
    private bool _isInitializing = true;
    private readonly List<PreparedScriptOption> _preparedScripts = new();
    private WaveInEvent? _waveIn;
    private WaveFileWriter? _waveWriter;
    private DispatcherTimer? _recordingTimer;
    private WaveInEvent? _micTestIn;
    private DispatcherTimer? _micTestTimer;
    private DispatcherTimer? _voiceDesignTimer;
    private DispatcherTimer? _voiceDesignDraftAutosaveTimer;
    private DateTime _recordingStartedAtUtc;
    private DateTime _voiceDesignStartedAtUtc;
    private int _voiceDesignCurrentStep;
    private int _voiceDesignTotalSteps;
    private float _recordingPeak;
    private float _micTestPeak;
    private string? _lastRecordedPath;
    private CancellationTokenSource? _generationCts;
    private CancellationTokenSource? _voiceDesignCts;
    private bool _pauseGenerationRequested;
    private bool _isUpdatingInputSelection;
    private bool _isUpdatingStylePreset;
    private bool _isApplyingVoiceDesignDraft;
    private bool _isWarmingQwenCustomVoiceSamples;
    private bool _isEnhancingAudio;
    private bool _isUpdatingEnhanceInputSelection;
    private bool _isNormalizingVoiceSelection;
    private CancellationTokenSource? _enhanceAudioCts;
    public MainWindow()
    {
        InitializeComponent();
        InputFilesList.ItemsSource = _inputFiles;
        QueueGrid.ItemsSource = _queueRows;
        EnhanceInputFilesList.ItemsSource = _enhanceInputFiles;
        EnhanceQueueGrid.ItemsSource = _enhanceQueueRows;
        HistoryGrid.ItemsSource = _historyRows;
        WireQueueSummaryEvents();
        WireEnhanceQueueSummaryEvents();
        ApplyFeatureFlagUi();
        ApplyUiIcons();
        Closing += MainWindow_OnClosing;
        Closed += MainWindow_OnClosed;
        LoadHeaderIcon();
        _voiceDesignDraftAutosaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
        _voiceDesignDraftAutosaveTimer.Tick += VoiceDesignDraftAutosaveTimer_OnTick;
        _ = InitializeAsync();
    }

    private void ApplyFeatureFlagUi()
    {
        if (AppFeatureFlags.AudioEnhancementEnabled)
        {
            return;
        }

        if (AudioEnhancementTabItem is not null)
        {
            AudioEnhancementTabItem.Visibility = Visibility.Collapsed;
        }

        if (EnhanceAudioCheckBox is not null)
        {
            EnhanceAudioCheckBox.Visibility = Visibility.Collapsed;
            EnhanceAudioCheckBox.IsChecked = false;
        }
    }

    private void WireQueueSummaryEvents()
    {
        _queueRows.CollectionChanged += QueueRows_OnCollectionChanged;
        foreach (var row in _queueRows)
        {
            row.PropertyChanged += QueueRow_OnPropertyChanged;
        }

        UpdateJobQueueSummary();
    }

    private void QueueRows_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems)
            {
                if (item is QueueRow row)
                {
                    row.PropertyChanged -= QueueRow_OnPropertyChanged;
                }
            }
        }

        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems)
            {
                if (item is QueueRow row)
                {
                    row.PropertyChanged += QueueRow_OnPropertyChanged;
                }
            }
        }

        UpdateJobQueueSummary();
    }

    private void WireEnhanceQueueSummaryEvents()
    {
        _enhanceQueueRows.CollectionChanged += EnhanceQueueRows_OnCollectionChanged;
        foreach (var row in _enhanceQueueRows)
        {
            row.PropertyChanged += EnhanceQueueRow_OnPropertyChanged;
        }

        UpdateEnhanceQueueSummary();
    }

    private void EnhanceQueueRows_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems)
            {
                if (item is QueueRow row)
                {
                    row.PropertyChanged -= EnhanceQueueRow_OnPropertyChanged;
                }
            }
        }

        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems)
            {
                if (item is QueueRow row)
                {
                    row.PropertyChanged += EnhanceQueueRow_OnPropertyChanged;
                }
            }
        }

        UpdateEnhanceQueueSummary();
    }

    private void EnhanceQueueRow_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(QueueRow.Status) or nameof(QueueRow.IsRunning))
        {
            UpdateEnhanceQueueSummary();
        }
    }

    private void QueueRow_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(QueueRow.Status) or nameof(QueueRow.IsRunning))
        {
            UpdateJobQueueSummary();
        }
    }

    private void UpdateJobQueueSummary()
    {
        if (JobQueueSummaryText is null)
        {
            return;
        }

        var total = _queueRows.Count;
        var done = _queueRows.Count(r => string.Equals(r.Status, "Done", StringComparison.OrdinalIgnoreCase));
        var left = Math.Max(0, total - done);
        JobQueueSummaryText.Text = $"({done}/{total} done, {left} left)";
    }

    private void UpdateEnhanceQueueSummary()
    {
        if (EnhanceQueueSummaryText is null)
        {
            return;
        }

        var total = _enhanceQueueRows.Count;
        var done = _enhanceQueueRows.Count(r => string.Equals(r.Status, "Done", StringComparison.OrdinalIgnoreCase));
        var left = Math.Max(0, total - done);
        EnhanceQueueSummaryText.Text = $"({done}/{total} done, {left} left)";
    }

    private async Task InitializeAsync()
    {
        try
        {
            _config = await _configStore.LoadAsync();
            var canContinue = await EnsureFirstRunSetupCompletedAsync();
            if (!canContinue)
            {
                return;
            }
            _config = await _configStore.LoadAsync();
            EnsureConfigDefaults();
            EnsureRuntimeFolders();
            await TryLoadStartupProjectAsync();
            if (_config.AutoDownloadModel)
            {
                _ = TryAutoDownloadModelAsync();
            }
            LoadVoiceList();
            RefreshStylePresetOptions();
            RefreshCloneVoicesList();
            InitializePreparedScripts();
            LoadProjectToUi(_project);
            RefreshProjectSelector();
            RefreshModelSummary();
            RefreshModelControlSummary();
            RefreshVoiceDesignReadiness();
            SyncMainDeviceComboFromConfig();
            RenderSystemInfo();
            RefreshInputFilePreparedState();
            ApplyAudioEnhanceUiFromProject();
        }
        finally
        {
            _isInitializing = false;
            _isGenerating = false;
            if (GenerateButton is not null)
            {
                GenerateButton.IsEnabled = true;
            }
        }
    }

    private async Task<bool> EnsureFirstRunSetupCompletedAsync()
    {
        if (!ShouldRunFirstRunSetup())
        {
            return true;
        }

        var existingState = await _installStateService.LoadAsync(RuntimePaths.AppRoot, CancellationToken.None);
        if (existingState is not null &&
            string.Equals(existingState.CompletionState, "installed", StringComparison.OrdinalIgnoreCase) &&
            !File.Exists(RuntimePaths.SetupRequiredFlagPath))
        {
            return true;
        }

        bool installCompleted = false;
        await Dispatcher.InvokeAsync(() =>
        {
            var wizard = new FirstRunSetupWindow(RuntimePaths.AppRoot)
            {
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            installCompleted = wizard.ShowDialog() == true && wizard.InstallCompleted;
        });

        if (!installCompleted)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                MessageBox.Show(
                    this,
                    "Audiobook Creator needs first-run component setup before it can continue.",
                    "Audiobook Creator",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                Close();
            });
        }
        else if (File.Exists(RuntimePaths.SetupRequiredFlagPath))
        {
            File.Delete(RuntimePaths.SetupRequiredFlagPath);
        }

        return installCompleted;
    }

    private static bool ShouldRunFirstRunSetup()
    {
        if (File.Exists(RuntimePaths.SetupRequiredFlagPath))
        {
            return true;
        }

        if (!File.Exists(RuntimePaths.InstallStatePath))
        {
            return false;
        }

        try
        {
            var json = File.ReadAllText(RuntimePaths.InstallStatePath);
            return json.IndexOf("\"completionState\":\"shell_only\"", StringComparison.OrdinalIgnoreCase) >= 0
                || json.IndexOf("\"completionState\": \"shell_only\"", StringComparison.OrdinalIgnoreCase) >= 0
                || json.IndexOf("\"completionState\":\"failed\"", StringComparison.OrdinalIgnoreCase) >= 0
                || json.IndexOf("\"completionState\": \"failed\"", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch
        {
            return true;
        }
    }

    private async Task TryLoadStartupProjectAsync()
    {
        var candidatePaths = new List<string>();
        if (!string.IsNullOrWhiteSpace(_currentProjectPath))
        {
            candidatePaths.Add(_currentProjectPath);
        }
        if (_config.RecentProjects is not null)
        {
            candidatePaths.AddRange(_config.RecentProjects.Where(path => !string.IsNullOrWhiteSpace(path)));
        }
        if (Directory.Exists(RuntimePaths.ProjectsDir))
        {
            candidatePaths.AddRange(
                Directory.GetFiles(RuntimePaths.ProjectsDir, $"*{JsonProjectStore.Extension}", SearchOption.TopDirectoryOnly)
                    .OrderByDescending(File.GetLastWriteTimeUtc));
        }

        foreach (var candidate in candidatePaths.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(candidate) || !File.Exists(candidate))
            {
                continue;
            }

            try
            {
                _project = await _projectStore.LoadAsync(candidate);
                _currentProjectPath = candidate;
                _config.LastOpenDir = Path.GetDirectoryName(candidate) ?? RuntimePaths.ProjectsDir;
                return;
            }
            catch
            {
                // Try next candidate.
            }
        }
    }

    private void EnsureConfigDefaults()
    {
        _config.ModelProfiles ??= new Dictionary<string, SynthesisSettings>();
        _config.ModelCacheDir = ModelCachePath.NormalizeInput(_config.ModelCacheDir);
        _config.PreferDevice = NormalizePreferDevice(_config.PreferDevice);
        _config.LocalModelBackend = NormalizeLocalModelBackendChoice(_config.LocalModelBackend, _config.LocalModelPreset, _config.ModelRepoId);
        var backendMode = (_config.BackendMode ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(backendMode) || backendMode == "auto")
        {
            _config.BackendMode = "local";
        }
        if (string.Equals(_config.ModelRepoId?.Trim(), "Qwen/Qwen3-TTS-0.6B", StringComparison.OrdinalIgnoreCase))
        {
            _config.ModelRepoId = "zukky/Qwen3-TTS-ONNX-DLL";
        }
        if (string.Equals(_config.ModelRepoId?.Trim(), "Qwen/Qwen3-TTS-12Hz-0.6B-Base", StringComparison.OrdinalIgnoreCase))
        {
            _config.ModelRepoId = "zukky/Qwen3-TTS-ONNX-DLL";
        }
        if (string.Equals(_config.ModelRepoId?.Trim(), "Qwen/Qwen3-TTS-12Hz-1.7B-Base", StringComparison.OrdinalIgnoreCase))
        {
            _config.ModelRepoId = "xkos/Qwen3-TTS-12Hz-1.7B-ONNX";
        }

        if (string.IsNullOrWhiteSpace(_config.DefaultOutputDir))
        {
            _config.DefaultOutputDir = "output";
        }
        if (string.IsNullOrWhiteSpace(_config.LocalModelPreset))
        {
            _config.LocalModelPreset = "chatterbox_onnx";
        }
        if (string.IsNullOrWhiteSpace(_config.ModelRepoId))
        {
            _config.ModelRepoId = "onnx-community/chatterbox-ONNX";
        }
        else if (string.Equals((_config.LocalModelPreset ?? string.Empty).Trim(), "chatterbox_onnx", StringComparison.OrdinalIgnoreCase))
        {
            _config.ModelRepoId = string.Equals(_config.LocalModelBackend, "python", StringComparison.OrdinalIgnoreCase)
                ? "ResembleAI/chatterbox"
                : "onnx-community/chatterbox-ONNX";
        }
        else if (string.Equals((_config.LocalModelPreset ?? string.Empty).Trim(), "kitten_tts", StringComparison.OrdinalIgnoreCase))
        {
            _config.ModelRepoId = "KittenML/kitten-tts-mini-0.8";
        }
        if (string.IsNullOrWhiteSpace(_config.ApiPreset))
        {
            _config.ApiPreset = "openai_default";
        }
        if (string.IsNullOrWhiteSpace(_config.ApiProvider))
        {
            _config.ApiProvider = "openai";
        }
        if (string.IsNullOrWhiteSpace(_config.ApiKeyAlibaba) &&
            string.IsNullOrWhiteSpace(_config.ApiKeyOpenAi) &&
            !string.IsNullOrWhiteSpace(_config.ApiKey))
        {
            var provider = (_config.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
            if (provider == "alibaba")
            {
                _config.ApiKeyAlibaba = _config.ApiKey;
            }
            else if (provider == "openai")
            {
                _config.ApiKeyOpenAi = _config.ApiKey;
            }
            else
            {
                _config.ApiKeyOpenAi = _config.ApiKey;
            }
        }
        if (string.IsNullOrWhiteSpace(_config.ApiModelId))
        {
            _config.ApiModelId = string.Equals(_config.ApiProvider, "openai", StringComparison.OrdinalIgnoreCase)
                ? "gpt-4o-mini-tts"
                : string.Equals(_config.ApiProvider, "alibaba", StringComparison.OrdinalIgnoreCase)
                    ? "qwen3-tts-flash"
                    : "gpt-4o-mini-tts";
        }
        if (string.IsNullOrWhiteSpace(_config.ApiBaseUrl))
        {
            _config.ApiBaseUrl = string.Equals(_config.ApiProvider, "alibaba", StringComparison.OrdinalIgnoreCase)
                ? "https://dashscope-intl.aliyuncs.com/api/v1"
                : string.Empty;
        }
        if (string.IsNullOrWhiteSpace(_config.ApiLanguageType))
        {
            _config.ApiLanguageType = "Auto";
        }
        if (string.IsNullOrWhiteSpace(_config.ApiVoiceDesignTargetModel))
        {
            _config.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
        }

        _config.LlmPrepProvider = string.IsNullOrWhiteSpace(_config.LlmPrepProvider) ? "local" : _config.LlmPrepProvider.Trim().ToLowerInvariant();
        _config.LlmPrepSplitModel = string.IsNullOrWhiteSpace(_config.LlmPrepSplitModel) ? "qwen2.5-7b-instruct" : _config.LlmPrepSplitModel.Trim();
        _config.LlmPrepInstructionModel = string.IsNullOrWhiteSpace(_config.LlmPrepInstructionModel) ? "qwen2.5-14b-instruct" : _config.LlmPrepInstructionModel.Trim();
        _config.LlmLocalModelDir = string.IsNullOrWhiteSpace(_config.LlmLocalModelDir) ? "models/llm" : _config.LlmLocalModelDir.Trim();
        _config.LlmLocalSplitModelFile = string.IsNullOrWhiteSpace(_config.LlmLocalSplitModelFile) ? "qwen2.5-7b-instruct.gguf" : _config.LlmLocalSplitModelFile.Trim();
        _config.LlmLocalInstructionModelFile = string.IsNullOrWhiteSpace(_config.LlmLocalInstructionModelFile) ? "qwen2.5-14b-instruct.gguf" : _config.LlmLocalInstructionModelFile.Trim();
        _config.LlmPrepTemperatureSplit = Math.Clamp(_config.LlmPrepTemperatureSplit <= 0 ? 0.2 : _config.LlmPrepTemperatureSplit, 0.0, 1.5);
        _config.LlmPrepTemperatureInstruction = Math.Clamp(_config.LlmPrepTemperatureInstruction <= 0 ? 0.6 : _config.LlmPrepTemperatureInstruction, 0.0, 1.5);
        _config.LlmPrepMaxTokensSplit = Math.Clamp(_config.LlmPrepMaxTokensSplit <= 0 ? 1024 : _config.LlmPrepMaxTokensSplit, 128, 8192);
        _config.LlmPrepMaxTokensInstruction = Math.Clamp(_config.LlmPrepMaxTokensInstruction <= 0 ? 512 : _config.LlmPrepMaxTokensInstruction, 64, 4096);
        _config.AudioEnhanceMode = string.IsNullOrWhiteSpace(_config.AudioEnhanceMode) ? "sfx_ambience" : _config.AudioEnhanceMode.Trim().ToLowerInvariant();
        _config.AudioEnhanceProvider = string.IsNullOrWhiteSpace(_config.AudioEnhanceProvider) ? "local_audiox" : _config.AudioEnhanceProvider.Trim().ToLowerInvariant();
        _config.AudioEnhanceVariant = string.IsNullOrWhiteSpace(_config.AudioEnhanceVariant) ? "audiox_base" : _config.AudioEnhanceVariant.Trim().ToLowerInvariant();
        _config.AudioEnhanceModelDir = string.IsNullOrWhiteSpace(_config.AudioEnhanceModelDir) ? "models/audiox" : _config.AudioEnhanceModelDir.Trim();
        var audioEnhanceRepo = (_config.AudioEnhanceModelRepoId ?? string.Empty).Trim();
        _config.AudioEnhanceModelRepoId = string.IsNullOrWhiteSpace(audioEnhanceRepo)
            ? "HKUSTAudio/AudioX"
            : (string.Equals(audioEnhanceRepo, "ZeyueT/AudioX", StringComparison.OrdinalIgnoreCase)
                || string.Equals(audioEnhanceRepo, "audiox", StringComparison.OrdinalIgnoreCase)
                ? "HKUSTAudio/AudioX"
                : audioEnhanceRepo);
        _config.AudioEnhanceAmbienceDb = Math.Clamp(_config.AudioEnhanceAmbienceDb, -40.0, 0.0);
        _config.AudioEnhanceOneShotDb = Math.Clamp(_config.AudioEnhanceOneShotDb, -30.0, 0.0);
        _config.AudioEnhanceDuckingDb = Math.Clamp(_config.AudioEnhanceDuckingDb, -24.0, 0.0);
        _config.AudioEnhanceCueMaxPerMinute = Math.Clamp(_config.AudioEnhanceCueMaxPerMinute <= 0 ? 10 : _config.AudioEnhanceCueMaxPerMinute, 1, 60);
    }

    private void InitializePreparedScripts()
    {
        if (_preparedScripts.Count > 0)
        {
            return;
        }

        _preparedScripts.Add(new PreparedScriptOption(
            "Warm Narration",
            "This is a voice sample for audiobook narration. I speak with a calm pace, clear diction, and steady volume from start to finish. "
            + "Each sentence is read naturally, without rushing, so the story sounds smooth, warm, and easy to follow."));
        _preparedScripts.Add(new PreparedScriptOption(
            "Dialogue Style",
            "The old library door opened with a soft creak. \"Are you sure this is the place?\" Mira asked quietly. "
            + "\"Yes,\" Daniel said, \"the map points to the third shelf near the window.\" "
            + "They paused for a moment, took a breath, and stepped forward together."));
        _preparedScripts.Add(new PreparedScriptOption(
            "Neutral Reading",
            "This recording is used to build a text to speech voice. Read in a neutral tone with stable loudness and clean pronunciation. "
            + "Avoid whispering, sharp peaks, or dramatic style changes while you complete the full paragraph."));
        _preparedScripts.Add(new PreparedScriptOption(
            "Coverage Check",
            "Today is April twenty third, and the temperature is seventy two degrees. I can read names like Alex, Chloe, George, and Sofia, "
            + "plus numbers such as seven, nineteen, forty six, and one hundred. Please keep every word clear and evenly paced."));

        PreparedScriptCombo.ItemsSource = _preparedScripts;
        PreparedScriptCombo.SelectedIndex = 0;
        RecordingLevelBar.Value = 0;
        RecordingStatusText.Text = "Recorder idle.";
        RecordingTimerText.Text = "00:00";
        RefreshMicrophoneDetectionInfo();
    }

    private void EnsureRuntimeFolders()
    {
        Directory.CreateDirectory(Path.Combine(RuntimePaths.AppRoot, "voices"));
        Directory.CreateDirectory(Path.Combine(RuntimePaths.AppRoot, "models"));
        Directory.CreateDirectory(ResolveOutputDir(_config.AudioEnhanceModelDir));
        Directory.CreateDirectory(ResolveOutputDir(_config.LlmLocalModelDir));
        Directory.CreateDirectory(ResolveOutputDir(_config.DefaultOutputDir));
        Directory.CreateDirectory(RuntimePaths.ProjectsDir);

        if (string.IsNullOrWhiteSpace(_project.OutputDir))
        {
            _project.OutputDir = ResolveProjectOutputDir(_config.DefaultOutputDir, _project.Name);
        }
    }

    private void LoadHeaderIcon()
    {
        var logoPng = FindIconPath("app_logo_512.png");
        var logoBitmap = LoadIconBitmapFromPathOrResource(logoPng, "app_logo_512.png");
        if (logoBitmap is not null)
        {
            try
            {
                HeaderIcon.Source = TrimTransparentPadding(logoBitmap);
            }
            catch
            {
                // Fallback to ico loading below.
            }
        }

        var candidates = new[]
        {
            Path.Combine(RuntimePaths.AppRoot, "audiobook_creator_icon.ico"),
            Path.Combine(RuntimePaths.AppRoot, "audiobookcreator_icon.ico"),
            Path.Combine(RuntimePaths.AppRoot, "audiobookcreator.ico")
        };
        var iconPath = candidates.FirstOrDefault(File.Exists)
            ?? Directory.GetFiles(RuntimePaths.AppRoot, "*.ico", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(path => Path.GetFileName(path).Contains("audiobook", StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(iconPath) || !File.Exists(iconPath))
        {
            return;
        }

        try
        {
            var decoder = new IconBitmapDecoder(new Uri(iconPath), BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            var frame = decoder.Frames
                .OrderByDescending(f => f.PixelWidth * f.PixelHeight)
                .FirstOrDefault();
            if (frame is null)
            {
                return;
            }
            if (HeaderIcon.Source is null)
            {
                HeaderIcon.Source = frame;
            }
            Icon = frame;
        }
        catch
        {
            // Keep running even if icon decode fails.
        }
    }

    private void ApplyUiIcons()
    {
        SetButtonIcon(NewProjectButton, "New Project", "new_project_512.png");
        SetButtonIcon(OpenProjectButton, "Open Project", "open_project_512.png");
        SetButtonIcon(SaveProjectButton, "Save", "save_512.png");
        SetButtonIcon(SettingsButton, "Settings", "settings_512.png");
        SetButtonIcon(InlineNewProjectButton, "New", "new_project_512.png");
        SetButtonIcon(AddFilesButton, "Add Files", "new_project_512.png");
    }

    private void SetButtonIcon(Button button, string text, string iconFileName)
    {
        var iconPath = FindIconPath(iconFileName);
        var bmp = LoadIconBitmapFromPathOrResource(iconPath, iconFileName);
        if (bmp is null)
        {
            button.Content = text;
            return;
        }

        try
        {
            var icon = TrimTransparentPadding(bmp);
            var iconSize = button == InlineNewProjectButton ? 16.0 : 15.0;

            var stack = new DockPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                LastChildFill = false
            };
            var iconImage = new Image
            {
                Source = icon,
                Width = iconSize,
                Height = iconSize,
                Margin = new Thickness(0, 0, 6, 0)
            };
            DockPanel.SetDock(iconImage, Dock.Left);
            stack.Children.Add(iconImage);
            stack.Children.Add(new TextBlock
            {
                Text = text,
                VerticalAlignment = VerticalAlignment.Center
            });

            button.Content = stack;
        }
        catch
        {
            button.Content = text;
        }
    }

    private static BitmapImage? LoadIconBitmapFromPathOrResource(string? iconPath, string iconFileName)
    {
        if (!string.IsNullOrWhiteSpace(iconPath) && File.Exists(iconPath))
        {
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(iconPath);
                bmp.EndInit();
                return bmp;
            }
            catch
            {
                // Fall through to embedded resource.
            }
        }

        try
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.UriSource = new Uri($"pack://application:,,,/Icon/{iconFileName}", UriKind.Absolute);
            bmp.EndInit();
            return bmp;
        }
        catch
        {
            return null;
        }
    }

    private static string? FindIconPath(string fileName)
    {
        var candidates = new[]
        {
            Path.Combine(RuntimePaths.AppRoot, "Icon", fileName),
            Path.Combine(RuntimePaths.AppRoot, "icon", fileName)
        };

        return candidates.FirstOrDefault(File.Exists);
    }

    private static string? ExtractPathFromListItem(object? item)
    {
        return item switch
        {
            FileChoice choice => choice.FullPath,
            VoiceItem voice => voice.FullPath,
            string raw when !string.IsNullOrWhiteSpace(raw) => raw,
            _ => null
        };
    }

    private string? GetSelectedVoicePath()
    {
        return ExtractPathFromListItem(VoiceCombo.SelectedItem);
    }

    private static string GetRefTextPathForVoice(string voicePath)
    {
        if (string.IsNullOrWhiteSpace(voicePath) || !File.Exists(voicePath))
            return string.Empty;
        var dir = Path.GetDirectoryName(voicePath);
        var stem = Path.GetFileNameWithoutExtension(voicePath);
        return string.IsNullOrEmpty(dir) ? string.Empty : Path.Combine(dir, stem + ".ref.txt");
    }

    private static string LoadRefTextForVoice(string voicePath)
    {
        var path = GetRefTextPathForVoice(voicePath);
        if (string.IsNullOrEmpty(path) || !File.Exists(path))
            return string.Empty;
        try
        {
            return File.ReadAllText(path).Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private bool SelectVoiceByPath(string? voicePath)
    {
        if (string.IsNullOrWhiteSpace(voicePath))
        {
            return false;
        }

        foreach (var item in VoiceCombo.Items)
        {
            var candidate = ExtractPathFromListItem(item);
            if (string.Equals(candidate, voicePath, StringComparison.OrdinalIgnoreCase))
            {
                VoiceCombo.SelectedItem = item;
                return true;
            }
        }

        return false;
    }

    private bool SelectCloneVoiceByPath(string? voicePath)
    {
        if (string.IsNullOrWhiteSpace(voicePath))
        {
            return false;
        }

        foreach (var item in CloneVoicesList.Items)
        {
            var candidate = ExtractPathFromListItem(item);
            if (string.Equals(candidate, voicePath, StringComparison.OrdinalIgnoreCase))
            {
                CloneVoicesList.SelectedItem = item;
                CloneVoicesList.ScrollIntoView(item);
                return true;
            }
        }

        return false;
    }

    private static bool IsKittenRepo(string? repoId)
        => (repoId ?? string.Empty).IndexOf("kitten-tts", StringComparison.OrdinalIgnoreCase) >= 0;

    private bool IsActiveKittenModel()
        => string.Equals((_config.BackendMode ?? "auto").Trim(), "api", StringComparison.OrdinalIgnoreCase) == false &&
           IsKittenRepo(_config.ModelRepoId);

    private bool IsApiOpenAiMode()
    {
        var backend = (_config.BackendMode ?? string.Empty).Trim().ToLowerInvariant();
        if (backend != "api")
        {
            return false;
        }

        var provider = (_config.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(provider))
        {
            var preset = (_config.ApiPreset ?? string.Empty).Trim().ToLowerInvariant();
            provider = preset.StartsWith("alibaba", StringComparison.OrdinalIgnoreCase) ? "alibaba" : "openai";
        }

        return provider == "openai";
    }

    private static string BuildKittenVoiceToken(string internalId)
        => $"kitten://{internalId}";

    private static bool TryParseKittenVoiceToken(string? raw, out string internalId)
    {
        internalId = string.Empty;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var s = raw.Trim();
        if (s.StartsWith("kitten://", StringComparison.OrdinalIgnoreCase))
        {
            internalId = s["kitten://".Length..];
            return !string.IsNullOrWhiteSpace(internalId);
        }
        if (s.StartsWith("kitten:", StringComparison.OrdinalIgnoreCase))
        {
            internalId = s["kitten:".Length..];
            return !string.IsNullOrWhiteSpace(internalId);
        }

        return false;
    }

    private string ResolveApiKeyForProvider(string? provider)
    {
        var key = (provider ?? string.Empty).Trim().ToLowerInvariant();
        if (key == "alibaba")
        {
            return _config.ApiKeyAlibaba ?? string.Empty;
        }
        if (key == "openai")
        {
            return _config.ApiKeyOpenAi ?? string.Empty;
        }
        return _config.ApiKeyOpenAi ?? _config.ApiKey ?? string.Empty;
    }

    private static readonly string[] QwenCustomVoiceSpeakers =
    {
        "Vivian", "Serena", "Uncle_Fu", "Dylan", "Eric", "Ryan", "Aiden", "Ono_Anna", "Sohee"
    };
    private static readonly string[] OpenAiApiVoices =
    {
        "alloy", "ash", "ballad", "coral", "echo", "sage", "shimmer", "verse"
    };
    private const string MixedVoiceToken = "voice-mixed://prepared";
    private const string MixedVoiceDisplayName = "Mixed (Prepare Chapter)";
    private const string OpenAiVoiceTokenPrefix = "api-openai-voice://";
    private const string QwenCustomVoiceSampleVersion = "v2_emotion_paragraph";
    private const string QwenCustomVoicePreviewText =
        "I found the old letter at sunrise, and my voice is calm as I begin to read. " +
        "Then the warning inside makes me tense, and I speak with urgency and fear. " +
        "When I realize my family is safe, relief softens my words into warm gratitude. " +
        "At last, hope returns, and I finish with quiet confidence for the journey ahead.";

    private static string BuildQwenCustomVoiceToken(string speaker)
        => $"qwen-customvoice://{speaker}";

    private static string BuildOpenAiVoiceToken(string voice)
        => $"{OpenAiVoiceTokenPrefix}{voice.Trim()}";

    private static bool TryParseOpenAiVoiceToken(string? raw, out string voice)
    {
        voice = string.Empty;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var s = raw.Trim();
        if (!s.StartsWith(OpenAiVoiceTokenPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        voice = s[OpenAiVoiceTokenPrefix.Length..].Trim();
        return !string.IsNullOrWhiteSpace(voice);
    }

    private static string BuildMixedVoiceToken()
        => MixedVoiceToken;

    private static bool IsMixedVoiceToken(string? raw)
    {
        return string.Equals((raw ?? string.Empty).Trim(), MixedVoiceToken, StringComparison.OrdinalIgnoreCase);
    }

    private bool IsPrepareMixedVoiceSelected()
    {
        return IsMixedVoiceToken(GetSelectedVoicePath());
    }

    private bool ShouldShowMixedVoiceOption()
    {
        return !string.Equals((_config.BackendMode ?? "auto").Trim(), "api", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsPrepareScriptAvailableForCurrentSelection()
    {
        return ShouldShowMixedVoiceOption() && IsPrepareMixedVoiceSelected();
    }

    private static bool TryParseQwenCustomVoiceToken(string? raw, out string speaker)
    {
        speaker = string.Empty;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var s = raw.Trim();
        if (!s.StartsWith("qwen-customvoice://", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        speaker = s["qwen-customvoice://".Length..].Trim();
        return !string.IsNullOrWhiteSpace(speaker);
    }

    private void LoadQwenCustomVoiceList()
    {
        VoiceCombo.Items.Clear();
        foreach (var speaker in QwenCustomVoiceSpeakers)
        {
            VoiceCombo.Items.Add(new VoiceItem(speaker, BuildQwenCustomVoiceToken(speaker), false, false));
        }
        if (ShouldShowMixedVoiceOption())
        {
            VoiceCombo.Items.Add(new VoiceItem(MixedVoiceDisplayName, BuildMixedVoiceToken(), false, false));
        }

        if (!string.IsNullOrWhiteSpace(_project.VoicePath))
        {
            if (IsMixedVoiceToken(_project.VoicePath) && SelectVoiceByPath(BuildMixedVoiceToken()))
            {
                _project.VoicePath = GetSelectedVoicePath() ?? BuildMixedVoiceToken();
                UpdateVoiceSelectionMode();
                return;
            }

            if (TryParseQwenCustomVoiceToken(_project.VoicePath, out var savedSpeaker) &&
                SelectVoiceByPath(BuildQwenCustomVoiceToken(savedSpeaker)))
            {
                _project.VoicePath = GetSelectedVoicePath() ?? BuildQwenCustomVoiceToken(savedSpeaker);
                UpdateVoiceSelectionMode();
                return;
            }

            var stem = Path.GetFileNameWithoutExtension(_project.VoicePath);
            var matched = QwenCustomVoiceSpeakers.FirstOrDefault(s => string.Equals(s, stem, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(matched) && SelectVoiceByPath(BuildQwenCustomVoiceToken(matched)))
            {
                _project.VoicePath = GetSelectedVoicePath() ?? BuildQwenCustomVoiceToken(matched);
                UpdateVoiceSelectionMode();
                return;
            }
        }

        if (VoiceCombo.Items.Count > 0)
        {
            VoiceCombo.SelectedIndex = 0;
            _project.VoicePath = GetSelectedVoicePath() ?? string.Empty;
        }

        UpdateVoiceSelectionMode();
        _ = WarmQwenCustomVoiceSamplesAsync();
    }

    private static string GetQwenCustomVoiceSamplePath(string speaker)
    {
        var dir = Path.Combine(RuntimePaths.AppRoot, "voices", "customvoice_samples");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, $"{SanitizeFileName(speaker)}.wav");
    }

    private static string GetQwenCustomVoiceSampleVersionPath()
    {
        var dir = Path.Combine(RuntimePaths.AppRoot, "voices", "customvoice_samples");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, ".preview_version");
    }

    private static string? TryResolveExistingQwenCustomVoiceSamplePath(string speaker)
    {
        var safe = SanitizeFileName(speaker);
        var candidates = new[]
        {
            GetQwenCustomVoiceSamplePath(speaker),
            Path.Combine(RuntimePaths.AppRoot, "voices", $"{safe}.wav"),
            Path.Combine(RuntimePaths.AppRoot, "voices", "customvoice_samples", $"{safe}.wav")
        };

        foreach (var path in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(path) && new FileInfo(path).Length > 0)
            {
                return path;
            }
        }

        return null;
    }

    private async Task WarmQwenCustomVoiceSamplesAsync()
    {
        if (!IsLocalQwenCustomVoiceMode() || _isWarmingQwenCustomVoiceSamples)
        {
            return;
        }

        var versionPath = GetQwenCustomVoiceSampleVersionPath();
        var currentVersion = File.Exists(versionPath) ? (File.ReadAllText(versionPath).Trim()) : string.Empty;
        var forceRegenerate = !string.Equals(currentVersion, QwenCustomVoiceSampleVersion, StringComparison.Ordinal);

        var targets = forceRegenerate
            ? QwenCustomVoiceSpeakers.ToList()
            : QwenCustomVoiceSpeakers.Where(s => !File.Exists(GetQwenCustomVoiceSamplePath(s))).ToList();
        if (targets.Count == 0)
        {
            return;
        }

        _isWarmingQwenCustomVoiceSamples = true;
        try
        {
            ITtsBackend? backend = null;
            try
            {
                backend = CreateBackend();
            }
            catch
            {
                return;
            }

            foreach (var speaker in targets)
            {
                if (!IsLocalQwenCustomVoiceMode())
                {
                    break;
                }

                try
                {
                    await EnsureQwenCustomVoiceSampleAsync(speaker, backend, forceRegenerate);
                }
                catch
                {
                    // Best effort warmup only.
                }
            }

            if (targets.All(s => File.Exists(GetQwenCustomVoiceSamplePath(s))))
            {
                File.WriteAllText(versionPath, QwenCustomVoiceSampleVersion);
            }
        }
        finally
        {
            _isWarmingQwenCustomVoiceSamples = false;
        }
    }

    private async Task<string> EnsureQwenCustomVoiceSampleAsync(string speaker, ITtsBackend? backend = null, bool forceRebuild = false)
    {
        var samplePath = GetQwenCustomVoiceSamplePath(speaker);
        if (!forceRebuild && File.Exists(samplePath) && new FileInfo(samplePath).Length > 0)
        {
            return samplePath;
        }

        var effectiveBackend = backend ?? CreateBackend();
        var settings = GetEffectiveLocalSettings();
        settings.StylePresetKey = "standard";
        ApplySelectedStyleRuntimeOverrides(settings);
        await effectiveBackend.SynthesizeAsync(
            CreateLocalTtsRequest(
                QwenCustomVoicePreviewText,
                BuildQwenCustomVoiceToken(speaker),
                samplePath,
                1.0f,
                settings),
            CancellationToken.None);
        return samplePath;
    }

    private static IReadOnlyList<(string Alias, string InternalId)> GetKittenVoiceChoices()
        => KittenTtsOnnxBackend.GetDefaultVoices();

    private void LoadKittenVoiceList()
    {
        VoiceCombo.Items.Clear();
        foreach (var (alias, internalId) in GetKittenVoiceChoices())
        {
            VoiceCombo.Items.Add(new VoiceItem($"{alias} ({internalId})", BuildKittenVoiceToken(internalId), false, false));
        }
        if (ShouldShowMixedVoiceOption())
        {
            VoiceCombo.Items.Add(new VoiceItem(MixedVoiceDisplayName, BuildMixedVoiceToken(), false, false));
        }

        if (!string.IsNullOrWhiteSpace(_project.VoicePath))
        {
            var target = _project.VoicePath;
            if (!TryParseKittenVoiceToken(target, out _) &&
                GetKittenVoiceChoices().FirstOrDefault(v =>
                    string.Equals(v.Alias, target, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(v.InternalId, target, StringComparison.OrdinalIgnoreCase)) is var match &&
                !string.IsNullOrWhiteSpace(match.InternalId))
            {
                target = BuildKittenVoiceToken(match.InternalId);
            }

            if (SelectVoiceByPath(target))
            {
                _project.VoicePath = GetSelectedVoicePath() ?? target;
                UpdateVoiceSelectionMode();
                return;
            }
        }

        if (VoiceCombo.Items.Count > 0)
        {
            VoiceCombo.SelectedIndex = 0;
            _project.VoicePath = GetSelectedVoicePath() ?? string.Empty;
        }
        UpdateVoiceSelectionMode();
    }

    private void UpdateVoiceSelectionMode()
    {
        var kitten = IsActiveKittenModel();
        var mixed = IsPrepareMixedVoiceSelected();
        var openAiApi = IsApiOpenAiMode();

        if (PlaySampleButton is not null)
        {
            PlaySampleButton.IsEnabled = !kitten && !mixed;
            PlaySampleButton.Opacity = (kitten || mixed) ? 0.75 : 1.0;
            PlaySampleButton.ToolTip = kitten
                ? "Kitten TTS uses embedded model voices. Preview by generating a short sample."
                : mixed
                ? "Mixed mode has no single narrator sample. Use Prepare Script and assign voices per part."
                : openAiApi
                ? "OpenAI API voices are cloud voices (no local WAV sample file to play)."
                : null;
        }

        if (VoiceModelNoteText is not null)
        {
            VoiceModelNoteText.Visibility = (kitten || mixed) ? Visibility.Visible : Visibility.Collapsed;
            VoiceModelNoteText.Text = kitten
                ? "Kitten TTS uses built-in model voices only (no voice cloning). Select a model voice from the list above."
                : mixed
                ? "Mixed mode: use Prepare Script to assign voice per part."
                : string.Empty;
        }
    }

    private void RefreshStylePresetOptions()
    {
        if (StylePresetCombo is null || StylePresetLabel is null)
        {
            return;
        }

        var selectedBefore = GetSelectedStylePresetKey();
        var targetKey = ResolvePreferredStylePresetKey();
        if (string.IsNullOrWhiteSpace(targetKey))
        {
            targetKey = selectedBefore;
        }

        var options = GetStyleOptionsForCurrentModel();
        _isUpdatingStylePreset = true;
        try
        {
            StylePresetCombo.Items.Clear();
            foreach (var option in options)
            {
                StylePresetCombo.Items.Add(new ComboBoxItem { Content = option.DisplayName, Tag = option.Key });
            }

            if (!TrySelectStylePresetByKey(targetKey) &&
                !TrySelectStylePresetByKey(selectedBefore) &&
                StylePresetCombo.Items.Count > 0)
            {
                StylePresetCombo.SelectedIndex = 0;
            }

            var styleSupported = options.Count > 0;
            StylePresetLabel.Visibility = styleSupported ? Visibility.Visible : Visibility.Collapsed;
            StylePresetCombo.Visibility = styleSupported ? Visibility.Visible : Visibility.Collapsed;
            StylePresetCombo.IsEnabled = styleSupported;
            StylePresetCombo.ToolTip = styleSupported
                ? null
                : "This model uses voice selection only and does not expose a separate style preset.";
        }
        finally
        {
            _isUpdatingStylePreset = false;
        }
    }

    private bool TrySelectStylePresetByKey(string? key)
    {
        if (string.IsNullOrWhiteSpace(key) || StylePresetCombo is null)
        {
            return false;
        }

        foreach (var item in StylePresetCombo.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), key, StringComparison.OrdinalIgnoreCase))
            {
                StylePresetCombo.SelectedItem = item;
                return true;
            }
        }

        return false;
    }

    private void ApplySavedStylePresetSelection()
    {
        RefreshStylePresetOptions();
    }

    private string GetSelectedStylePresetKey()
    {
        if (StylePresetCombo?.SelectedItem is ComboBoxItem { Tag: not null } comboItem)
        {
            var tag = comboItem.Tag?.ToString();
            if (!string.IsNullOrWhiteSpace(tag))
            {
                return tag.Trim().ToLowerInvariant();
            }
        }

        var fallback = StylePresetCombo?.Text;
        return string.IsNullOrWhiteSpace(fallback) ? "standard" : fallback.Trim().ToLowerInvariant();
    }

    private string ResolvePreferredStylePresetKey()
    {
        if (TryGetBestActiveModelProfile(out var perModel, out _) &&
            !string.IsNullOrWhiteSpace(perModel.StylePresetKey))
        {
            return perModel.StylePresetKey;
        }

        if (_project.Settings is not null && !string.IsNullOrWhiteSpace(_project.Settings.StylePresetKey))
        {
            return _project.Settings.StylePresetKey;
        }

        return "standard";
    }

    private List<StylePresetOption> GetStyleOptionsForCurrentModel()
    {
        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        if (repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            return new List<StylePresetOption>
            {
                new("standard", "Narration (Neutral)"),
                new("expressive", "Expressive Dialogue"),
                new("calm", "Calm / Soft"),
                new("dramatic", "Dramatic"),
                new("energetic", "Energetic"),
                new("whispered", "Whispered")
            };
        }

        if (repo.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase))
        {
            return new List<StylePresetOption>();
        }

        return new List<StylePresetOption>
        {
            new("standard", "Standard"),
            new("expressive", "Expressive"),
            new("calm", "Calm")
        };
    }

    private void LoadVoiceList()
    {
        if (IsApiOpenAiMode())
        {
            LoadOpenAiApiVoiceList();
            return;
        }

        if (IsActiveKittenModel())
        {
            LoadKittenVoiceList();
            return;
        }

        if (IsLocalQwenCustomVoiceMode())
        {
            LoadQwenCustomVoiceList();
            return;
        }

        NormalizeLocalVoiceLibrary();

        VoiceCombo.Items.Clear();
        var voiceDir = Path.Combine(RuntimePaths.AppRoot, "voices");
        var candidates = Directory.Exists(voiceDir)
            ? Directory.GetFiles(voiceDir, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
                .OrderBy(Path.GetFileName)
                .ToList()
            : new List<string>();

        var isQwen = IsQwenRepo(_config.ModelRepoId);
        foreach (var path in candidates)
        {
            var refPath = GetRefTextPathForVoice(path);
            var hasRef = !string.IsNullOrEmpty(refPath) && File.Exists(refPath);
            VoiceCombo.Items.Add(new VoiceItem(Path.GetFileName(path), path, hasRef, hasRef && isQwen));
        }
        if (ShouldShowMixedVoiceOption())
        {
            VoiceCombo.Items.Add(new VoiceItem(MixedVoiceDisplayName, BuildMixedVoiceToken(), false, false));
        }

        if (!string.IsNullOrWhiteSpace(_project.VoicePath) &&
            (File.Exists(_project.VoicePath) || IsMixedVoiceToken(_project.VoicePath)))
        {
            if (SelectVoiceByPath(_project.VoicePath))
            {
                UpdateVoiceSelectionMode();
                return;
            }
        }

        if (VoiceCombo.Items.Count > 0)
        {
            VoiceCombo.SelectedIndex = 0;
            _project.VoicePath = GetSelectedVoicePath() ?? string.Empty;
        }

        UpdateVoiceSelectionMode();
    }

    private void LoadOpenAiApiVoiceList()
    {
        VoiceCombo.Items.Clear();

        var options = new List<string>(OpenAiApiVoices);
        var hasProjectVoice = TryParseOpenAiVoiceToken(_project.VoicePath, out var projectVoice) &&
                              options.Contains(projectVoice, StringComparer.OrdinalIgnoreCase);
        var hasConfigVoice = !string.IsNullOrWhiteSpace(_config.ApiVoice) &&
                             options.Contains(_config.ApiVoice.Trim(), StringComparer.OrdinalIgnoreCase);

        foreach (var voice in options.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            VoiceCombo.Items.Add(new VoiceItem(voice, BuildOpenAiVoiceToken(voice), false, false));
        }

        var preferred = hasProjectVoice
            ? projectVoice
            : hasConfigVoice
                ? _config.ApiVoice.Trim()
                : OpenAiApiVoices[0];
        var preferredToken = BuildOpenAiVoiceToken(preferred);
        if (!SelectVoiceByPath(preferredToken) && VoiceCombo.Items.Count > 0)
        {
            VoiceCombo.SelectedIndex = 0;
        }

        if (TryParseOpenAiVoiceToken(GetSelectedVoicePath(), out var selectedVoice))
        {
            _config.ApiVoice = selectedVoice;
            _project.VoicePath = BuildOpenAiVoiceToken(selectedVoice);
        }
        else
        {
            _config.ApiVoice = OpenAiApiVoices[0];
            _project.VoicePath = BuildOpenAiVoiceToken(OpenAiApiVoices[0]);
        }

        UpdateVoiceSelectionMode();
    }

    private void RefreshCloneVoicesList()
    {
        CloneVoicesList.Items.Clear();
        var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
        if (!Directory.Exists(voicesDir))
        {
            return;
        }

        var clones = Directory.GetFiles(voicesDir, "*.wav", SearchOption.TopDirectoryOnly)
            .OrderBy(Path.GetFileName)
            .ToList();

        foreach (var clone in clones)
        {
            CloneVoicesList.Items.Add(new FileChoice(Path.GetFileName(clone), clone));
        }
    }

    private void RefreshModelSummary()
    {
        var backendMode = (_config.BackendMode ?? "auto").Trim().ToLowerInvariant();
        if (backendMode == "api")
        {
            var provider = string.IsNullOrWhiteSpace(_config.ApiProvider) ? "api" : _config.ApiProvider.Trim();
            var apiModel = string.IsNullOrWhiteSpace(_config.ApiModelId) ? "default" : _config.ApiModelId.Trim();
            ModelSummaryText.Text = $"AI Model: API ({provider} / {apiModel})";
            return;
        }

        var repo = string.IsNullOrWhiteSpace(_config.ModelRepoId) ? "onnx-community/chatterbox-ONNX" : _config.ModelRepoId.Trim();
        var display = repo;
        if (repo.Contains("chatterbox", StringComparison.OrdinalIgnoreCase))
        {
            display = "chatterbox-ONNX";
        }
        else if (repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            display = "Qwen3-TTS";
        }
        else if (repo.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase))
        {
            display = "Kitten TTS Mini 0.8";
        }
        else
        {
            var normalized = repo.Replace('\\', '/');
            var last = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
            if (!string.IsNullOrWhiteSpace(last))
            {
                display = last;
            }
        }
        ModelSummaryText.Text = $"AI Model: {display}";
    }

    private void SyncMainDeviceComboFromConfig()
    {
        if (MainDeviceCombo is null)
        {
            return;
        }

        var target = NormalizePreferDevice(_config.PreferDevice);
        foreach (var item in MainDeviceCombo.Items)
        {
            if (item is ComboBoxItem comboItem &&
                string.Equals(comboItem.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase))
            {
                MainDeviceCombo.SelectedItem = comboItem;
                return;
            }
        }

        MainDeviceCombo.Text = target;
    }

    private void RenderSystemInfo()
    {
        if (DetectedSystemText is null)
        {
            return;
        }

        var profile = SystemProbe.Detect();
        var cudaSummary = profile.CudaAvailable
            ? "ready"
            : "not ready";
        DetectedSystemText.Text =
            $"Detected: CPU {profile.CpuLogical} threads, RAM {profile.RamGb} GB, GPU-{profile.GpuName}, CUDA-{(profile.CudaAvailable ? "Yes" : "No")} ({cudaSummary}), FP16-{(profile.Fp16Available ? "Yes" : "No")}";
    }

    private void RefreshProjectSelector()
    {
        _isRefreshingProjectSelector = true;
        try
        {
            ProjectSelectorCombo.Items.Clear();
            _projectSelectorPaths.Clear();

            var addedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            void AddItem(string display, string? path)
            {
                if (string.IsNullOrWhiteSpace(display))
                {
                    return;
                }

                string? normalized = null;
                if (!string.IsNullOrWhiteSpace(path))
                {
                    normalized = Path.GetFullPath(path);
                    if (!File.Exists(normalized))
                    {
                        return;
                    }
                    if (!addedPaths.Add(normalized))
                    {
                        return;
                    }
                }

                ProjectSelectorCombo.Items.Add(display);
                _projectSelectorPaths.Add(normalized);
            }

            var currentDisplay = string.IsNullOrWhiteSpace(_project.Name) ? "New Audiobook Project" : _project.Name;
            if (!string.IsNullOrWhiteSpace(_currentProjectPath) && File.Exists(_currentProjectPath))
            {
                AddItem(currentDisplay, _currentProjectPath);
            }
            else
            {
                AddItem(currentDisplay, null);
            }

            foreach (var recent in _config.RecentProjects.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(recent))
                {
                    continue;
                }
                var display = Path.GetFileNameWithoutExtension(recent);
                AddItem(display, recent);
            }

            if (Directory.Exists(RuntimePaths.ProjectsDir))
            {
                foreach (var path in Directory.GetFiles(RuntimePaths.ProjectsDir, $"*{JsonProjectStore.Extension}", SearchOption.TopDirectoryOnly)
                             .OrderBy(Path.GetFileNameWithoutExtension))
                {
                    var display = Path.GetFileNameWithoutExtension(path);
                    AddItem(display, path);
                }
            }

            if (ProjectSelectorCombo.Items.Count == 0)
            {
                ProjectSelectorCombo.Items.Add("New Audiobook Project");
                _projectSelectorPaths.Add(null);
            }

            var selectedIndex = 0;
            if (!string.IsNullOrWhiteSpace(_currentProjectPath))
            {
                for (var i = 0; i < _projectSelectorPaths.Count; i++)
                {
                    if (string.Equals(_projectSelectorPaths[i], _currentProjectPath, StringComparison.OrdinalIgnoreCase))
                    {
                        selectedIndex = i;
                        break;
                    }
                }
            }

            ProjectSelectorCombo.SelectedIndex = Math.Clamp(selectedIndex, 0, ProjectSelectorCombo.Items.Count - 1);
            ProjectSelectorCombo.Text = currentDisplay;
        }
        finally
        {
            _isRefreshingProjectSelector = false;
        }
    }

    private void LoadProjectToUi(ProjectDocument project)
    {
        project.PreparedScriptsBySourcePath ??= new Dictionary<string, PreparedScriptDocument>(StringComparer.OrdinalIgnoreCase);
        project.CueOverrides ??= new List<EnhanceCueOverride>();
        project.AudioEnhanceInputFiles ??= new List<ProjectAudioEnhanceInputFileState>();
        project.AudioEnhancePreparedSegmentsBySourcePath ??= new Dictionary<string, List<EnhancePreparedSegment>>(StringComparer.OrdinalIgnoreCase);
        var baseOutput = ResolveOutputDir(_config.DefaultOutputDir);
        var resolvedOutput = ResolveOutputDir(project.OutputDir);
        if (string.IsNullOrWhiteSpace(project.OutputDir) ||
            string.Equals(resolvedOutput, baseOutput, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(project.OutputDir.Trim(), "output", StringComparison.OrdinalIgnoreCase))
        {
            project.OutputDir = ResolveProjectOutputDir(_config.DefaultOutputDir, project.Name);
        }
        else
        {
            project.OutputDir = resolvedOutput;
        }
        OutputFolderTextBox.Text = project.OutputDir;
        if (EnhanceOutputFolderTextBox is not null)
        {
            EnhanceOutputFolderTextBox.Text = string.IsNullOrWhiteSpace(project.AudioEnhanceOutputDir)
                ? ResolveEnhanceOutputDir(project.OutputDir, forceFromBaseOutput: true)
                : ResolveOutputDir(project.AudioEnhanceOutputDir);
        }
        if (EnhanceTabEnableCheckBox is not null)
        {
            EnhanceTabEnableCheckBox.IsChecked = project.AudioEnhanceEnabled;
        }
        if (EnhanceSceneNotesTextBox is not null)
        {
            EnhanceSceneNotesTextBox.Text = project.AudioEnhanceSceneNotes ?? string.Empty;
        }

        _isUpdatingInputSelection = true;
        _inputFiles.Clear();
        _queueRows.Clear();
        _historyRows.Clear();
        var sourceFiles = project.SourceTextFiles;
        if ((sourceFiles is null || sourceFiles.Count == 0) && !string.IsNullOrWhiteSpace(project.SourceTextPath))
        {
            sourceFiles = new List<string> { project.SourceTextPath };
        }
        foreach (var file in sourceFiles ?? Enumerable.Empty<string>())
        {
            if (string.IsNullOrWhiteSpace(file))
            {
                continue;
            }

            _inputFiles.Add(new InputFileEntry(file, false));
        }
        RefreshInputFilePreparedState();

        if (UseAllInputFilesCheckBox is not null)
        {
            UseAllInputFilesCheckBox.IsChecked = false;
        }
        _isUpdatingInputSelection = false;

        _isUpdatingEnhanceInputSelection = true;
        _enhanceInputFiles.Clear();
        _enhanceQueueRows.Clear();
        var storedEnhanceInputs = project.AudioEnhanceInputFiles ?? new List<ProjectAudioEnhanceInputFileState>();
        var savedEnhanceStates = storedEnhanceInputs
            .Where(x => !string.IsNullOrWhiteSpace(x.FullPath))
            .GroupBy(x => x.FullPath, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Last().IsSelected, StringComparer.OrdinalIgnoreCase);

        var enhanceInputs = storedEnhanceInputs
            .Where(x => !string.IsNullOrWhiteSpace(x.FullPath))
            .ToList();
        if (enhanceInputs.Count == 0 && project.AudioEnhancePreparedSegmentsBySourcePath.Count > 0)
        {
            enhanceInputs = project.AudioEnhancePreparedSegmentsBySourcePath.Keys
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => new ProjectAudioEnhanceInputFileState { FullPath = path, IsSelected = false })
                .ToList();
        }

        foreach (var outputAudio in EnumerateProjectEnhanceOutputAudioFiles(project))
        {
            if (enhanceInputs.Any(x => string.Equals(x.FullPath, outputAudio, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            enhanceInputs.Add(new ProjectAudioEnhanceInputFileState
            {
                FullPath = outputAudio,
                IsSelected = savedEnhanceStates.TryGetValue(outputAudio, out var selected) && selected
            });
        }

        try
        {
            foreach (var file in enhanceInputs)
            {
                _enhanceInputFiles.Add(new AudioEnhanceInputFileEntry(file.FullPath, file.IsSelected));
            }

            if (EnhanceUseAllInputFilesCheckBox is not null)
            {
                var useAll = project.AudioEnhanceUseAllInputFiles;
                EnhanceUseAllInputFilesCheckBox.IsChecked = useAll;
                if (useAll)
                {
                    foreach (var entry in _enhanceInputFiles)
                    {
                        entry.IsSelected = true;
                    }
                }
            }
        }
        finally
        {
            _isUpdatingEnhanceInputSelection = false;
        }
        UpdateEnhanceQueueSummary();

        if (project.QueueRows is { Count: > 0 })
        {
            foreach (var saved in project.QueueRows)
            {
                if (string.IsNullOrWhiteSpace(saved.SourcePath))
                {
                    continue;
                }

                var restoredOutput = string.IsNullOrWhiteSpace(saved.OutputPath) ? string.Empty : saved.OutputPath;
                if (string.IsNullOrWhiteSpace(restoredOutput) || !File.Exists(restoredOutput))
                {
                    if (TryResolveExistingProjectOutput(project, saved.SourcePath, out var inferredOutput))
                    {
                        restoredOutput = inferredOutput;
                    }
                }
                var restoredStatus = string.IsNullOrWhiteSpace(saved.Status) ? "Queued" : saved.Status;
                var looksStaleQueued = string.Equals(restoredStatus, "Queued", StringComparison.OrdinalIgnoreCase) &&
                                       !string.IsNullOrWhiteSpace(restoredOutput) &&
                                       File.Exists(restoredOutput);

                _queueRows.Add(new QueueRow
                {
                    FileName = string.IsNullOrWhiteSpace(saved.FileName) ? Path.GetFileName(saved.SourcePath) : saved.FileName,
                    SourcePath = saved.SourcePath,
                    Status = looksStaleQueued ? "Done" : restoredStatus,
                    ProgressLabel = looksStaleQueued ? "100%" : (string.IsNullOrWhiteSpace(saved.ProgressLabel) ? "0%" : saved.ProgressLabel),
                    ProgressValue = looksStaleQueued ? 100 : Math.Clamp(saved.ProgressValue, 0, 100),
                    ChunkInfo = looksStaleQueued ? "Done" : (string.IsNullOrWhiteSpace(saved.ChunkInfo) ? "Queued" : saved.ChunkInfo),
                    Eta = looksStaleQueued ? "Done (restored)" : (string.IsNullOrWhiteSpace(saved.Eta) ? "Queued" : saved.Eta),
                    OutputPath = restoredOutput,
                    IsRunning = false,
                    IsPaused = false
                });
            }
        }

        if (project.HistoryRows is { Count: > 0 })
        {
            foreach (var saved in project.HistoryRows)
            {
                if (string.IsNullOrWhiteSpace(saved.FileName) && string.IsNullOrWhiteSpace(saved.SourcePath))
                {
                    continue;
                }

                _historyRows.Add(new HistoryRow
                {
                    FileName = string.IsNullOrWhiteSpace(saved.FileName)
                        ? Path.GetFileName(saved.SourcePath ?? string.Empty)
                        : saved.FileName,
                    SourcePath = saved.SourcePath ?? string.Empty,
                    OutputPath = saved.OutputPath ?? string.Empty,
                    ModelLabel = saved.ModelLabel ?? string.Empty,
                    DeviceLabel = saved.DeviceLabel ?? string.Empty,
                    VoiceLabel = saved.VoiceLabel ?? string.Empty,
                    CompletedAt = saved.CompletedAt ?? string.Empty,
                    DurationLabel = saved.DurationLabel ?? string.Empty
                });
            }
        }

        BackfillHistoryFromRestoredDoneQueueRows();

        // Ensure newly added input files appear in queue even if the saved project predates queue persistence.
        foreach (var input in _inputFiles)
        {
            if (_queueRows.Any(r => string.Equals(r.SourcePath, input.FullPath, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            _queueRows.Add(new QueueRow
            {
                FileName = input.FileName,
                SourcePath = input.FullPath,
                Status = TryResolveExistingProjectOutput(project, input.FullPath, out var inferredOutput) ? "Done" : "Queued",
                ProgressLabel = inferredOutput is not null ? "100%" : "0%",
                ProgressValue = inferredOutput is not null ? 100 : 0,
                ChunkInfo = inferredOutput is not null ? "Done" : "Queued",
                Eta = inferredOutput is not null ? "Done (restored)" : "Queued",
                OutputPath = inferredOutput ?? string.Empty
            });
        }

        if (!string.IsNullOrWhiteSpace(project.VoicePath) &&
            (File.Exists(project.VoicePath) ||
             IsMixedVoiceToken(project.VoicePath) ||
             (IsApiOpenAiMode() && TryParseOpenAiVoiceToken(project.VoicePath, out _)) ||
             (IsActiveKittenModel() && TryParseKittenVoiceToken(project.VoicePath, out _)) ||
             (IsLocalQwenCustomVoiceMode() && TryParseQwenCustomVoiceToken(project.VoicePath, out _))))
        {
            SelectVoiceByPath(project.VoicePath);
        }
        else if (VoiceCombo.Items.Count > 0 && VoiceCombo.SelectedIndex < 0)
        {
            VoiceCombo.SelectedIndex = 0;
            _project.VoicePath = GetSelectedVoicePath() ?? string.Empty;
        }

        var speed = project.Settings?.Atempo ?? 1.0;
        SpeedSlider.Value = Math.Clamp(speed, 0.5, 1.5);
        SpeedLabel.Text = $"{SpeedSlider.Value:0.0}x";
        ApplySavedStylePresetSelection();
        RefreshModelControlSummary();
        RefreshInputFilePreparedState();
        ApplyVoiceDesignDraftToUi(project);
        ApplyAudioEnhanceUiFromProject();
        ProjectSelectorCombo.Text = string.IsNullOrWhiteSpace(project.Name) ? "New Audiobook Project" : project.Name;
    }

    private ProjectDocument CollectProjectFromUi()
    {
        var selectedVoice = GetSelectedVoicePath() ?? string.Empty;
        if (IsApiOpenAiMode() && TryParseOpenAiVoiceToken(selectedVoice, out var apiVoice))
        {
            _config.ApiVoice = apiVoice;
        }
        _project.Name = ProjectSelectorCombo.Text?.Trim() is { Length: > 0 } name ? name : "New Audiobook Project";
        _project.VoicePath = selectedVoice;
        var requestedOutput = OutputFolderTextBox.Text.Trim();
        var baseOutput = ResolveOutputDir(_config.DefaultOutputDir);
        var resolvedRequested = ResolveOutputDir(requestedOutput);
        _project.OutputDir = string.IsNullOrWhiteSpace(requestedOutput) ||
                             string.Equals(resolvedRequested, baseOutput, StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(requestedOutput, "output", StringComparison.OrdinalIgnoreCase)
            ? ResolveProjectOutputDir(_config.DefaultOutputDir, _project.Name)
            : resolvedRequested;
        _project.SourceTextFiles = _inputFiles
            .Select(x => x.FullPath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        _project.SourceTextPath = _project.SourceTextFiles.FirstOrDefault() ?? string.Empty;
        _project.Settings.Atempo = Math.Round(SpeedSlider.Value, 2);
        _project.Settings.StylePresetKey = GetSelectedStylePresetKey();
        CaptureVoiceDesignDraftFromUi(_project);
        _project.AudioEnhanceEnabled = EnhanceAudioCheckBox?.IsChecked == true;
        _project.AudioEnhanceUseLlmCueRefine = _config.AudioEnhanceUseLlmCueRefine;
        _project.AudioEnhanceProfile = "balanced";
        _project.AudioEnhanceUseAllInputFiles = EnhanceUseAllInputFilesCheckBox?.IsChecked == true;
        _project.AudioEnhanceOutputDir = (EnhanceOutputFolderTextBox?.Text ?? string.Empty).Trim();
        _project.AudioEnhanceSceneNotes = (EnhanceSceneNotesTextBox?.Text ?? string.Empty).Trim();
        _project.AudioEnhanceInputFiles = _enhanceInputFiles
            .Select(x => new ProjectAudioEnhanceInputFileState
            {
                FullPath = x.FullPath,
                IsSelected = x.IsSelected
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.FullPath))
            .DistinctBy(x => x.FullPath, StringComparer.OrdinalIgnoreCase)
            .ToList();
        _project.CueOverrides ??= new List<EnhanceCueOverride>();
        _project.AudioEnhancePreparedSegmentsBySourcePath ??= new Dictionary<string, List<EnhancePreparedSegment>>(StringComparer.OrdinalIgnoreCase);
        _project.PreparedScriptsBySourcePath ??= new Dictionary<string, PreparedScriptDocument>(StringComparer.OrdinalIgnoreCase);
        _project.QueueRows = _queueRows
            .Select(r => new ProjectQueueRowState
            {
                FileName = r.FileName,
                SourcePath = r.SourcePath,
                Status = r.IsRunning ? "Queued" : (r.Status ?? "Queued"),
                ProgressLabel = r.IsRunning ? "0%" : (r.ProgressLabel ?? "0%"),
                ProgressValue = r.IsRunning ? 0 : Math.Clamp(r.ProgressValue, 0, 100),
                ChunkInfo = r.IsRunning ? "Queued" : (r.ChunkInfo ?? "Queued"),
                Eta = r.IsRunning ? "Queued" : (r.Eta ?? "Queued"),
                OutputPath = r.OutputPath ?? string.Empty
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.SourcePath))
            .ToList();
        _project.HistoryRows = _historyRows
            .Select(h => new ProjectHistoryRowState
            {
                FileName = h.FileName,
                SourcePath = h.SourcePath,
                OutputPath = h.OutputPath,
                ModelLabel = h.ModelLabel,
                DeviceLabel = h.DeviceLabel,
                VoiceLabel = h.VoiceLabel,
                CompletedAt = h.CompletedAt,
                DurationLabel = h.DurationLabel
            })
            .ToList();
        return _project;
    }

    private void CaptureVoiceDesignDraftFromUi(ProjectDocument project)
    {
        project.VoiceDesignDraftName = (QwenVoiceDesignNameTextBox?.Text ?? string.Empty).Trim();
        project.VoiceDesignDraftText = (QwenVoiceDesignTextBox?.Text ?? string.Empty).Trim();
        project.VoiceDesignDraftPrompt = (QwenVoicePromptTextBox?.Text ?? string.Empty).Trim();
        project.VoiceDesignDraftLanguage = ResolveVoiceDesignLanguageId();
    }

    private void VoiceDesignDraftField_OnChanged(object sender, TextChangedEventArgs e)
    {
        QueueVoiceDesignDraftAutoSave();
    }

    private void VoiceDesignDraftLanguage_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        QueueVoiceDesignDraftAutoSave();
    }

    private void QueueVoiceDesignDraftAutoSave()
    {
        if (_isInitializing || _isApplyingVoiceDesignDraft || _isRefreshingProjectSelector)
        {
            return;
        }

        CaptureVoiceDesignDraftFromUi(_project);
        if (_voiceDesignDraftAutosaveTimer is null)
        {
            _ = AutoSaveCurrentProjectAsync();
            return;
        }

        _voiceDesignDraftAutosaveTimer.Stop();
        _voiceDesignDraftAutosaveTimer.Start();
    }

    private async void VoiceDesignDraftAutosaveTimer_OnTick(object? sender, EventArgs e)
    {
        if (_voiceDesignDraftAutosaveTimer is not null)
        {
            _voiceDesignDraftAutosaveTimer.Stop();
        }
        await AutoSaveCurrentProjectAsync();
    }

    private void ApplyVoiceDesignDraftToUi(ProjectDocument project)
    {
        _isApplyingVoiceDesignDraft = true;
        try
        {
        if (QwenVoiceDesignNameTextBox is not null)
        {
            QwenVoiceDesignNameTextBox.Text = project.VoiceDesignDraftName ?? string.Empty;
        }
        if (QwenVoiceDesignTextBox is not null)
        {
            QwenVoiceDesignTextBox.Text = project.VoiceDesignDraftText ?? string.Empty;
        }
        if (QwenVoicePromptTextBox is not null)
        {
            QwenVoicePromptTextBox.Text = project.VoiceDesignDraftPrompt ?? string.Empty;
        }

        var language = string.IsNullOrWhiteSpace(project.VoiceDesignDraftLanguage)
            ? "auto"
            : project.VoiceDesignDraftLanguage.Trim().ToLowerInvariant();
        if (QwenVoiceDesignLanguageCombo is null)
        {
            return;
        }

        var target = QwenVoiceDesignLanguageCombo.Items
            .OfType<ComboBoxItem>()
            .FirstOrDefault(i => string.Equals(i.Tag?.ToString(), language, StringComparison.OrdinalIgnoreCase));
        if (target is not null)
        {
            QwenVoiceDesignLanguageCombo.SelectedItem = target;
        }
        else if (QwenVoiceDesignLanguageCombo.Items.Count > 0)
        {
            QwenVoiceDesignLanguageCombo.SelectedIndex = 0;
        }
        }
        finally
        {
            _isApplyingVoiceDesignDraft = false;
        }
    }

    private static bool TryResolveExistingProjectOutput(ProjectDocument project, string sourcePath, out string outputPath)
    {
        outputPath = string.Empty;
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return false;
        }

        var outputDir = string.IsNullOrWhiteSpace(project.OutputDir)
            ? ResolveOutputDir("output")
            : ResolveOutputDir(project.OutputDir);
        var stem = Path.GetFileNameWithoutExtension(sourcePath);
        if (string.IsNullOrWhiteSpace(stem))
        {
            return false;
        }

        var preferredExt = NormalizeLocalOutputExtension(project.Settings?.OutputFormat);
        var candidates = new List<string>(3)
        {
            Path.Combine(outputDir, $"{stem}.{preferredExt}"),
            Path.Combine(outputDir, $"{stem}.wav"),
            Path.Combine(outputDir, $"{stem}.mp3")
        };

        foreach (var path in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(path))
            {
                outputPath = path;
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<string> EnumerateProjectEnhanceOutputAudioFiles(ProjectDocument project)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var history in project.HistoryRows ?? Enumerable.Empty<ProjectHistoryRowState>())
        {
            var output = (history.OutputPath ?? string.Empty).Trim();
            if (!IsProjectEnhanceCandidateAudioFile(output))
            {
                continue;
            }

            if (seen.Add(output))
            {
                yield return output;
            }
        }

        foreach (var queue in project.QueueRows ?? Enumerable.Empty<ProjectQueueRowState>())
        {
            var output = (queue.OutputPath ?? string.Empty).Trim();
            if (!IsProjectEnhanceCandidateAudioFile(output))
            {
                continue;
            }

            if (seen.Add(output))
            {
                yield return output;
            }
        }

        var outputDir = string.IsNullOrWhiteSpace(project.OutputDir)
            ? ResolveOutputDir("output")
            : ResolveOutputDir(project.OutputDir);
        if (!Directory.Exists(outputDir))
        {
            yield break;
        }

        foreach (var file in Directory.EnumerateFiles(outputDir, "*.*", SearchOption.TopDirectoryOnly))
        {
            if (!IsProjectEnhanceCandidateAudioFile(file))
            {
                continue;
            }

            if (seen.Add(file))
            {
                yield return file;
            }
        }
    }

    private static bool IsProjectEnhanceCandidateAudioFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return false;
        }

        var ext = Path.GetExtension(path);
        if (!string.Equals(ext, ".wav", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(ext, ".mp3", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(ext, ".flac", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(ext, ".m4a", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(ext, ".ogg", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var fileName = Path.GetFileName(path);
        if (fileName.EndsWith(".narration_only.wav", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith("_enhanced.wav", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private List<string> GetSelectedInputFilesFromUi()
    {
        var useAll = UseAllInputFilesCheckBox?.IsChecked == true;
        var query = useAll
            ? _inputFiles.Select(x => x.FullPath)
            : _inputFiles.Where(x => x.IsSelected).Select(x => x.FullPath);

        return query
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizePreparedScriptKey(string sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(sourcePath).Trim().ToLowerInvariant();
        }
        catch
        {
            return sourcePath.Trim().ToLowerInvariant();
        }
    }

    private static string BuildSourceSignature(string sourcePath)
    {
        try
        {
            var fi = new FileInfo(sourcePath);
            return $"{fi.Length}:{fi.LastWriteTimeUtc.Ticks}";
        }
        catch
        {
            return string.Empty;
        }
    }

    private bool IsLocalQwenCustomVoiceMode()
    {
        var backend = (_config.BackendMode ?? string.Empty).Trim().ToLowerInvariant();
        if (backend == "api")
        {
            return false;
        }

        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        return repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase) &&
               repo.Contains("customvoice", StringComparison.OrdinalIgnoreCase);
    }

    private PreparedScriptDocument? TryGetPreparedScript(string sourcePath)
    {
        var key = NormalizePreparedScriptKey(sourcePath);
        if (string.IsNullOrWhiteSpace(key))
        {
            return null;
        }

        _project.PreparedScriptsBySourcePath ??= new Dictionary<string, PreparedScriptDocument>(StringComparer.OrdinalIgnoreCase);
        return _project.PreparedScriptsBySourcePath.TryGetValue(key, out var prepared) ? prepared : null;
    }

    private bool HasUsablePreparedScript(string sourcePath)
    {
        var prepared = TryGetPreparedScript(sourcePath);
        return prepared?.Parts?.Any(p => !string.IsNullOrWhiteSpace((p.Text ?? string.Empty).Trim())) == true;
    }

    private void SetPreparedScript(string sourcePath, PreparedScriptDocument document)
    {
        var key = NormalizePreparedScriptKey(sourcePath);
        if (string.IsNullOrWhiteSpace(key))
        {
            return;
        }

        _project.PreparedScriptsBySourcePath ??= new Dictionary<string, PreparedScriptDocument>(StringComparer.OrdinalIgnoreCase);
        _project.PreparedScriptsBySourcePath[key] = document;
    }

    private void RemovePreparedScript(string sourcePath)
    {
        var key = NormalizePreparedScriptKey(sourcePath);
        if (string.IsNullOrWhiteSpace(key) || _project.PreparedScriptsBySourcePath is null)
        {
            return;
        }

        _project.PreparedScriptsBySourcePath.Remove(key);
    }

    private void RefreshInputFilePreparedState()
    {
        var showPrep = IsPrepareScriptAvailableForCurrentSelection();
        foreach (var entry in _inputFiles)
        {
            var prep = TryGetPreparedScript(entry.FullPath);
            var count = prep?.Parts?.Count(p => !string.IsNullOrWhiteSpace((p.Text ?? string.Empty).Trim())) ?? 0;
            var stale = prep is not null && !string.Equals(prep.SourceSignature ?? string.Empty, BuildSourceSignature(entry.FullPath), StringComparison.Ordinal);
            entry.IsPrepared = prep is not null && count > 0;
            entry.PreparedPartCount = count;
            entry.IsPreparedStale = stale;
            entry.CanPrepareScript = showPrep;
        }
    }

    private async void NewProjectButton_OnClick(object sender, RoutedEventArgs e)
    {
        var projectName = PromptForProjectName(_project.Name);
        if (string.IsNullOrWhiteSpace(projectName))
        {
            return;
        }

        _project = new ProjectDocument
        {
            Name = projectName,
            OutputDir = ResolveProjectOutputDir(_config.DefaultOutputDir, projectName)
        };
        _inputFiles.Clear();
        _queueRows.Clear();
        _historyRows.Clear();

        var safeFileName = SanitizeFileName(projectName);
        if (string.IsNullOrWhiteSpace(safeFileName))
        {
            safeFileName = "New_Audiobook_Project";
        }
        _currentProjectPath = MakeUniqueProjectPath(safeFileName);

        await _projectStore.SaveAsync(_project, _currentProjectPath);
        await RegisterRecentProjectAsync(_currentProjectPath);

        LoadProjectToUi(_project);
        RefreshProjectSelector();
        ProjectSelectorCombo.Text = _project.Name;
        RefreshModelSummary();
        await _configStore.SaveAsync(_config);
        MessageBox.Show(this, $"Project created and saved:\n{_currentProjectPath}", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async void OpenProjectButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Audiobook Project (*.abproj)|*.abproj",
            InitialDirectory = Directory.Exists(_config.LastOpenDir) ? _config.LastOpenDir : RuntimePaths.ProjectsDir
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        _project = await _projectStore.LoadAsync(dialog.FileName);
        _currentProjectPath = dialog.FileName;
        _config.LastOpenDir = Path.GetDirectoryName(dialog.FileName) ?? RuntimePaths.ProjectsDir;
        await RegisterRecentProjectAsync(dialog.FileName);
        LoadProjectToUi(_project);
        RefreshProjectSelector();
        ProjectSelectorCombo.Text = _project.Name;
    }

    private async void SaveProjectButton_OnClick(object sender, RoutedEventArgs e)
    {
        CollectProjectFromUi();

        if (string.IsNullOrWhiteSpace(_currentProjectPath))
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Audiobook Project (*.abproj)|*.abproj",
                InitialDirectory = Directory.Exists(_config.LastOpenDir) ? _config.LastOpenDir : RuntimePaths.ProjectsDir,
                FileName = $"{_project.Name}.abproj"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            _currentProjectPath = dialog.FileName;
            _config.LastOpenDir = Path.GetDirectoryName(dialog.FileName) ?? RuntimePaths.ProjectsDir;
        }

        await _projectStore.SaveAsync(_project, _currentProjectPath!);
        await RegisterRecentProjectAsync(_currentProjectPath!);
        MessageBox.Show(this, "Project saved.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async Task RegisterRecentProjectAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        _config.RecentProjects.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        _config.RecentProjects.Insert(0, path);
        _config.RecentProjects = _config.RecentProjects.Take(20).ToList();
        await _configStore.SaveAsync(_config);
    }

    private async void AddFilesButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt",
            Multiselect = true,
            InitialDirectory = Directory.Exists(_config.LastOpenDir) ? _config.LastOpenDir : RuntimePaths.AppRoot
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var selectAll = UseAllInputFilesCheckBox?.IsChecked == true;
        foreach (var file in dialog.FileNames)
        {
            if (!_inputFiles.Any(x => string.Equals(x.FullPath, file, StringComparison.OrdinalIgnoreCase)))
            {
                _inputFiles.Add(new InputFileEntry(file, selectAll));
            }
        }

        RefreshInputFilePreparedState();

        _config.LastOpenDir = Path.GetDirectoryName(dialog.FileName) ?? _config.LastOpenDir;
        await AutoSaveCurrentProjectAsync();
    }

    private async void RemoveInputFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string path } || string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var existing = _inputFiles.FirstOrDefault(x => string.Equals(x.FullPath, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            _inputFiles.Remove(existing);
            RemovePreparedScript(path);
        }

        RefreshInputFilePreparedState();
        await AutoSaveCurrentProjectAsync();
    }

    private async void ClearAllInputFilesButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_inputFiles.Count == 0)
        {
            return;
        }

        var result = MessageBox.Show(
            this,
            "Clear all input files from the list?",
            "Input Files",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _inputFiles.Clear();
        _project.PreparedScriptsBySourcePath?.Clear();
        RefreshInputFilePreparedState();
        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceAddFilesButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Audio Files (*.wav;*.mp3;*.flac;*.m4a;*.ogg)|*.wav;*.mp3;*.flac;*.m4a;*.ogg|All Files (*.*)|*.*",
            Multiselect = true,
            InitialDirectory = Directory.Exists(_config.LastOpenDir) ? _config.LastOpenDir : RuntimePaths.AppRoot
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var selectAll = EnhanceUseAllInputFilesCheckBox?.IsChecked == true;
        foreach (var file in dialog.FileNames)
        {
            if (!_enhanceInputFiles.Any(x => string.Equals(x.FullPath, file, StringComparison.OrdinalIgnoreCase)))
            {
                _enhanceInputFiles.Add(new AudioEnhanceInputFileEntry(file, selectAll));
            }
        }

        _config.LastOpenDir = Path.GetDirectoryName(dialog.FileName) ?? _config.LastOpenDir;
        UpdateEnhanceQueueSummary();
        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceClearInputFilesButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_enhanceInputFiles.Count == 0)
        {
            return;
        }

        var result = MessageBox.Show(
            this,
            "Clear all audio input files from enhancement list?",
            "Audio Enhancement",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _enhanceInputFiles.Clear();
        _enhanceQueueRows.Clear();
        UpdateEnhanceQueueSummary();
        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceRemoveInputFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string path } || string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var existing = _enhanceInputFiles.FirstOrDefault(x => string.Equals(x.FullPath, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            _enhanceInputFiles.Remove(existing);
        }

        var queueExisting = _enhanceQueueRows.FirstOrDefault(x => string.Equals(x.SourcePath, path, StringComparison.OrdinalIgnoreCase));
        if (queueExisting is not null && !queueExisting.IsRunning)
        {
            _enhanceQueueRows.Remove(queueExisting);
        }

        UpdateEnhanceQueueSummary();
        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceUseAllInputFilesCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingEnhanceInputSelection || EnhanceUseAllInputFilesCheckBox is null)
        {
            return;
        }

        var useAll = EnhanceUseAllInputFilesCheckBox.IsChecked == true;
        _isUpdatingEnhanceInputSelection = true;
        try
        {
            foreach (var entry in _enhanceInputFiles)
            {
                entry.IsSelected = useAll;
            }
        }
        finally
        {
            _isUpdatingEnhanceInputSelection = false;
        }

        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceInputFileSelectionCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingEnhanceInputSelection)
        {
            return;
        }

        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceBrowseOutputButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        EnhanceOutputFolderTextBox.Text = dialog.SelectedPath;
        await AutoSaveCurrentProjectAsync();
    }

    private async void EnhanceOpenPrepForFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string sourcePath } || string.IsNullOrWhiteSpace(sourcePath))
        {
            return;
        }

        if (!File.Exists(sourcePath))
        {
            MessageBox.Show(this, "Audio file not found:\n" + sourcePath, "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _project.AudioEnhancePreparedSegmentsBySourcePath ??= new Dictionary<string, List<EnhancePreparedSegment>>(StringComparer.OrdinalIgnoreCase);
        var current = _project.AudioEnhancePreparedSegmentsBySourcePath.TryGetValue(sourcePath, out var segs)
            ? segs
            : new List<EnhancePreparedSegment>();

        var window = new AudioEnhancePrepWindow(sourcePath, current)
        {
            Owner = this
        };
        if (window.ShowDialog() != true)
        {
            return;
        }

        var saved = (window.ResultSegments ?? new List<EnhancePreparedSegment>())
            .Where(s => s.EndSeconds > s.StartSeconds)
            .OrderBy(s => s.Order)
            .ThenBy(s => s.StartSeconds)
            .ToList();

        if (saved.Count == 0)
        {
            _project.AudioEnhancePreparedSegmentsBySourcePath.Remove(sourcePath);
        }
        else
        {
            _project.AudioEnhancePreparedSegmentsBySourcePath[sourcePath] = saved;
        }

        await AutoSaveCurrentProjectAsync();
    }
    private static bool TryFindSubtitleSidecar(string sourceAudioPath, out string subtitlePath)
    {
        subtitlePath = string.Empty;
        if (string.IsNullOrWhiteSpace(sourceAudioPath))
        {
            return false;
        }

        var basePath = Path.Combine(Path.GetDirectoryName(sourceAudioPath) ?? string.Empty, Path.GetFileNameWithoutExtension(sourceAudioPath));
        var srt = basePath + ".srt";
        if (File.Exists(srt))
        {
            subtitlePath = srt;
            return true;
        }

        var ass = basePath + ".ass";
        if (File.Exists(ass))
        {
            subtitlePath = ass;
            return true;
        }

        return false;
    }

    private static List<EnhancePreparedSegment> ParseSubtitleFileToPreparedSegments(string subtitlePath)
    {
        if (string.IsNullOrWhiteSpace(subtitlePath) || !File.Exists(subtitlePath))
        {
            return new List<EnhancePreparedSegment>();
        }

        var ext = Path.GetExtension(subtitlePath).Trim().ToLowerInvariant();
        var cues = ext switch
        {
            ".ass" => ParseAssCues(File.ReadAllLines(subtitlePath)),
            _ => ParseSrtCues(File.ReadAllLines(subtitlePath))
        };

        return cues
            .Where(c => c.EndSeconds > c.StartSeconds)
            .Select((c, i) => new EnhancePreparedSegment
            {
                Id = Guid.NewGuid().ToString("N"),
                Order = i + 1,
                StartSeconds = c.StartSeconds,
                EndSeconds = c.EndSeconds,
                Text = c.Text,
                AmbiencePrompt = string.Empty,
                OneShotPrompt = string.Empty,
                OneShotSeconds = (c.StartSeconds + c.EndSeconds) * 0.5,
                Intensity = 0.6,
                Enabled = true
            })
            .ToList();
    }

    private static List<SubtitleCue> ParseSrtCues(IReadOnlyList<string> lines)
    {
        var cues = new List<SubtitleCue>();
        var block = new List<string>();
        void FlushBlock()
        {
            if (block.Count == 0)
            {
                return;
            }
            var timeLine = block.FirstOrDefault(l => l.Contains("-->", StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(timeLine))
            {
                var parts = timeLine.Split(new[] { "-->" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 && TryParseSrtTimestamp(parts[0].Trim(), out var start) && TryParseSrtTimestamp(parts[1].Trim(), out var end))
                {
                    var textLines = block
                        .Where(l => !l.Contains("-->", StringComparison.Ordinal) && !Regex.IsMatch(l.Trim(), @"^\d+$"))
                        .ToArray();
                    var text = string.Join(" ", textLines).Trim();
                    cues.Add(new SubtitleCue(start, end, text));
                }
            }
            block.Clear();
        }

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                FlushBlock();
                continue;
            }
            block.Add(line.Trim());
        }
        FlushBlock();
        return cues;
    }

    private static List<SubtitleCue> ParseAssCues(IReadOnlyList<string> lines)
    {
        var cues = new List<SubtitleCue>();
        foreach (var raw in lines)
        {
            var line = raw?.Trim() ?? string.Empty;
            if (!line.StartsWith("Dialogue:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var payload = line.Substring(line.IndexOf(':') + 1).Trim();
            var fields = payload.Split(',');
            if (fields.Length < 10)
            {
                continue;
            }

            if (!TryParseAssTimestamp(fields[1].Trim(), out var start) ||
                !TryParseAssTimestamp(fields[2].Trim(), out var end))
            {
                continue;
            }

            var text = string.Join(",", fields.Skip(9));
            text = Regex.Replace(text, @"\{.*?\}", string.Empty).Replace(@"\N", " ").Trim();
            cues.Add(new SubtitleCue(start, end, text));
        }
        return cues;
    }

    private static bool TryParseSrtTimestamp(string value, out double seconds)
    {
        seconds = 0;
        var parts = value.Split(':', ',', '.');
        if (parts.Length < 4)
        {
            return false;
        }
        if (!int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var h) ||
            !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var m) ||
            !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var s) ||
            !int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var ms))
        {
            return false;
        }
        seconds = (h * 3600) + (m * 60) + s + (ms / 1000.0);
        return seconds >= 0;
    }

    private static bool TryParseAssTimestamp(string value, out double seconds)
    {
        seconds = 0;
        var parts = value.Split(':', '.');
        if (parts.Length < 4)
        {
            return false;
        }
        if (!int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var h) ||
            !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var m) ||
            !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var s) ||
            !int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var cs))
        {
            return false;
        }
        seconds = (h * 3600) + (m * 60) + s + (cs / 100.0);
        return seconds >= 0;
    }

    private bool TryBuildPreparedEnhanceCues(string sourcePath, out List<EnhanceCue> cues, out List<TimedTextUnit> units)
    {
        cues = new List<EnhanceCue>();
        units = new List<TimedTextUnit>();

        _project.AudioEnhancePreparedSegmentsBySourcePath ??= new Dictionary<string, List<EnhancePreparedSegment>>(StringComparer.OrdinalIgnoreCase);
        if (!_project.AudioEnhancePreparedSegmentsBySourcePath.TryGetValue(sourcePath, out var segments) || segments is null || segments.Count == 0)
        {
            return false;
        }

        var duration = TryGetAudioDurationSeconds(sourcePath, out var sec) ? sec : 0;
        foreach (var seg in segments.OrderBy(s => s.Order).ThenBy(s => s.StartSeconds))
        {
            if (!seg.Enabled)
            {
                continue;
            }

            var start = Math.Max(0, seg.StartSeconds);
            var end = Math.Max(start + 0.05, seg.EndSeconds);
            if (duration > 0)
            {
                start = Math.Clamp(start, 0, duration);
                end = Math.Clamp(end, start + 0.05, duration);
            }

            units.Add(new TimedTextUnit(seg.Text ?? string.Empty, start, end));

            if (!string.IsNullOrWhiteSpace(seg.AmbiencePrompt))
            {
                cues.Add(new EnhanceCue
                {
                    Id = $"prep_amb_{seg.Id}",
                    StartSeconds = start,
                    EndSeconds = end,
                    CueType = "ambience",
                    Prompt = seg.AmbiencePrompt.Trim(),
                    Intensity = Math.Clamp(seg.Intensity <= 0 ? 0.6 : seg.Intensity, 0.1, 1.0),
                    Source = "prepared"
                });
            }

            if (!string.IsNullOrWhiteSpace(seg.OneShotPrompt))
            {
                var oneShotAt = Math.Clamp(seg.OneShotSeconds, start, end);
                var oneShotEnd = duration > 0
                    ? Math.Min(duration, oneShotAt + 1.2)
                    : oneShotAt + 1.2;
                cues.Add(new EnhanceCue
                {
                    Id = $"prep_one_{seg.Id}",
                    StartSeconds = oneShotAt,
                    EndSeconds = Math.Max(oneShotAt + 0.2, oneShotEnd),
                    CueType = "oneshot",
                    Prompt = seg.OneShotPrompt.Trim(),
                    Intensity = Math.Clamp(seg.Intensity <= 0 ? 0.8 : seg.Intensity, 0.1, 1.0),
                    Source = "prepared"
                });
            }
        }

        cues = cues
            .Where(c => c.EndSeconds > c.StartSeconds && !string.IsNullOrWhiteSpace(c.Prompt))
            .OrderBy(c => c.StartSeconds)
            .ToList();

        units = units
            .Where(u => u.EndSeconds > u.StartSeconds)
            .OrderBy(u => u.StartSeconds)
            .ToList();

        return cues.Count > 0;
    }

    private async void EnhanceGenerateButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isEnhancingAudio)
        {
            _enhanceAudioCts?.Cancel();
            return;
        }

        if (EnhanceTabEnableCheckBox?.IsChecked != true)
        {
            MessageBox.Show(this, "Enable enhancement first.", "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var selected = _enhanceInputFiles
            .Where(x => x.IsSelected)
            .Select(x => x.FullPath)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show(this, "Select at least one audio file to enhance.", "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var outputDir = ResolveEnhanceOutputDir();
        Directory.CreateDirectory(outputDir);

        _isEnhancingAudio = true;
        _enhanceAudioCts = new CancellationTokenSource();
        EnhanceGenerateButton.Content = "Stop";
        EnhanceGenerateButton.IsEnabled = true;
        EnhanceClearQueueButton.IsEnabled = false;

        try
        {
            foreach (var input in selected)
            {
                _enhanceAudioCts.Token.ThrowIfCancellationRequested();

                var row = _enhanceQueueRows.FirstOrDefault(x => string.Equals(x.SourcePath, input, StringComparison.OrdinalIgnoreCase));
                if (row is null)
                {
                    row = new QueueRow
                    {
                        FileName = Path.GetFileName(input),
                        SourcePath = input
                    };
                    _enhanceQueueRows.Add(row);
                }

                var inputName = Path.GetFileNameWithoutExtension(input);
                var outPath = Path.Combine(outputDir, $"{inputName}_enhanced.wav");
                var rowStart = DateTime.UtcNow;
                row.Status = "Enhancing...";
                row.ProgressLabel = "5%";
                row.ProgressValue = 5;
                row.ChunkInfo = "Preparing narration...";
                row.Eta = "Starting...";
                row.IsRunning = true;
                row.IsPaused = false;
                row.OutputPath = string.Empty;
                EnhanceQueueGrid.Items.Refresh();

                try
                {
                    var notes = (EnhanceSceneNotesTextBox?.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(notes))
                    {
                        notes = Path.GetFileNameWithoutExtension(input).Replace('_', ' ').Replace('-', ' ');
                    }

                    List<EnhanceCue>? preparedCues = null;
                    List<TimedTextUnit>? preparedUnits = null;
                    if (TryBuildPreparedEnhanceCues(input, out var timelineCues, out var timelineUnits))
                    {
                        preparedCues = timelineCues;
                        preparedUnits = timelineUnits;
                        row.ChunkInfo = "Using prepared timeline";
                    }

                    var progress = new Progress<AudioEnhanceProgress>(p =>
                    {
                        var total = Math.Max(p.Total, 1);
                        var done = Math.Clamp(p.Completed, 0, total);
                        var baseProgress = 10 + (int)Math.Round((done * 85.0) / total);
                        row.Status = p.Stage;
                        row.ProgressLabel = $"{Math.Clamp(baseProgress, 10, 99)}%";
                        row.ProgressValue = Math.Clamp(baseProgress, 10, 99);
                        row.ChunkInfo = total > 1 ? $"{done}/{total}" : p.Stage;
                        row.Eta = FormatDuration(DateTime.UtcNow - rowStart);
                        EnhanceQueueGrid.Items.Refresh();
                    });

                    var result = await RunAudioEnhanceForExternalAudioAsync(
                        input,
                        outPath,
                        notes,
                        preparedCues,
                        preparedUnits,
                        _enhanceAudioCts.Token,
                        progress);

                    var noEnhanceCues = result.Success &&
                                        !string.IsNullOrWhiteSpace(result.Message) &&
                                        result.Message.Contains("narration-only", StringComparison.OrdinalIgnoreCase);
                    row.Status = noEnhanceCues ? "Done (no cues)" : "Done";
                    row.ProgressLabel = "100%";
                    row.ProgressValue = 100;
                    row.ChunkInfo = !result.Success
                        ? "Done (enhance warning)"
                        : noEnhanceCues
                            ? "Narration only"
                            : "Done";
                    row.Eta = FormatDuration(DateTime.UtcNow - rowStart);
                    row.OutputPath = outPath;
                    row.IsRunning = false;
                    row.IsPaused = false;
                    TraceGenerate($"Enhanced audio: {row.FileName} -> {outPath} | {result.Message}");
                }
                catch (OperationCanceledException)
                {
                    row.Status = "Stopped";
                    row.ChunkInfo = "Stopped";
                    row.Eta = "Stopped";
                    row.IsRunning = false;
                    row.IsPaused = false;
                    EnhanceQueueGrid.Items.Refresh();
                    break;
                }
                catch (Exception ex)
                {
                    row.Status = "Failed";
                    row.ProgressLabel = "0%";
                    row.ProgressValue = 0;
                    row.ChunkInfo = "Failed";
                    row.Eta = "Failed";
                    row.IsRunning = false;
                    row.IsPaused = false;
                    MessageBox.Show(this, NormalizeGenerationError(ex.Message), "Audio Enhancement Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    EnhanceQueueGrid.Items.Refresh();
                }
            }
        }
        finally
        {
            _enhanceAudioCts?.Dispose();
            _enhanceAudioCts = null;
            _isEnhancingAudio = false;
            EnhanceGenerateButton.Content = "Enhance Audio";
            EnhanceGenerateButton.IsEnabled = true;
            EnhanceClearQueueButton.IsEnabled = true;
            UpdateEnhanceQueueSummary();
        }
    }

    private void EnhanceClearQueueButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isEnhancingAudio)
        {
            MessageBox.Show(this, "Stop enhancement before clearing queue.", "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _enhanceQueueRows.Clear();
        UpdateEnhanceQueueSummary();
    }

    private async void UseAllInputFilesCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing || _isUpdatingInputSelection || UseAllInputFilesCheckBox is null)
        {
            return;
        }

        var useAll = UseAllInputFilesCheckBox.IsChecked == true;
        _isUpdatingInputSelection = true;
        try
        {
            foreach (var entry in _inputFiles)
            {
                entry.IsSelected = useAll;
            }
        }
        finally
        {
            _isUpdatingInputSelection = false;
        }

        await AutoSaveCurrentProjectAsync();
    }

    private async void InputFileSelectionCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing || _isUpdatingInputSelection)
        {
            return;
        }

        await AutoSaveCurrentProjectAsync();
    }

    private async void PrepareInputFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string path } || string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        await OpenPrepareWindowForPathAsync(path);
    }

    private async Task OpenPrepareWindowForPathAsync(string sourcePath)
    {
        if (!IsPrepareScriptAvailableForCurrentSelection())
        {
            MessageBox.Show(this,
                "Select Narrator Voice = Mixed (Prepare Chapter) to use Prepare Script for this model.",
                "Prepare Script",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        if (!File.Exists(sourcePath))
        {
            MessageBox.Show(this, "Source file not found:\n" + sourcePath, "Prepare Script", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var sourceText = await File.ReadAllTextAsync(sourcePath);
        var sourceSig = BuildSourceSignature(sourcePath);
        var existing = TryGetPreparedScript(sourcePath);
        var allowInstructions = IsLocalQwenCustomVoiceMode();
        var voices = VoiceCombo.Items.OfType<VoiceItem>()
            .Select(v => (v.FullPath ?? string.Empty).Trim())
            .Where(v => !string.IsNullOrWhiteSpace(v) && !IsMixedVoiceToken(v))
            .Where(v => allowInstructions
                ? TryParseQwenCustomVoiceToken(v, out _)
                : IsActiveKittenModel()
                    ? TryParseKittenVoiceToken(v, out _)
                    : File.Exists(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var dlg = new ScriptPrepWindow(_config, sourcePath, sourceText, sourceSig, existing, voices, allowInstructions)
        {
            Owner = this
        };
        var ok = dlg.ShowDialog();
        if (dlg.Result is null)
        {
            return;
        }

        SetPreparedScript(sourcePath, dlg.Result);
        RefreshInputFilePreparedState();
        await AutoSaveCurrentProjectAsync();
        if (ok != true)
        {
            TraceGenerate($"Prepared script saved for '{Path.GetFileName(sourcePath)}' (window closed without Save & Close).");
        }
    }

    private void BrowseOutputButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        var baseOutput = ResolveOutputDir(_config.DefaultOutputDir);
        var selected = dialog.SelectedPath;
        OutputFolderTextBox.Text = string.Equals(selected, baseOutput, StringComparison.OrdinalIgnoreCase)
            ? ResolveProjectOutputDir(_config.DefaultOutputDir, _project.Name)
            : selected;
        _ = AutoSaveCurrentProjectAsync();
    }

    private async void OutputFolderTextBox_OnLostFocus(object sender, RoutedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }
        await AutoSaveCurrentProjectAsync();
    }

    private async void VoiceCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing || _isRefreshingProjectSelector || _isNormalizingVoiceSelection)
        {
            return;
        }

        var selected = GetSelectedVoicePath();
        if (string.IsNullOrWhiteSpace(selected))
        {
            return;
        }

        if (IsApiOpenAiMode() && TryParseOpenAiVoiceToken(selected, out var apiVoice))
        {
            _config.ApiVoice = apiVoice;
        }

        if (!IsApiOpenAiMode() &&
            !IsActiveKittenModel() &&
            !IsLocalQwenCustomVoiceMode() &&
            !IsMixedVoiceToken(selected) &&
            File.Exists(selected) &&
            !selected.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
        {
            if (!TryNormalizeLocalVoiceFile(selected, out var normalizedPath, out var error))
            {
                MessageBox.Show(this, error, "Voice Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!string.Equals(selected, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                _isNormalizingVoiceSelection = true;
                try
                {
                    LoadVoiceList();
                    SelectVoiceByPath(normalizedPath);
                }
                finally
                {
                    _isNormalizingVoiceSelection = false;
                }

                selected = normalizedPath;
            }
        }

        _project.VoicePath = selected;
        await AutoSaveCurrentProjectAsync();
        UpdateVoiceSelectionMode();
        RefreshInputFilePreparedState();

    }

    private async void PlaySampleButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (IsApiOpenAiMode())
        {
            MessageBox.Show(this,
                "OpenAI API voices do not have local sample files. Use Generate to preview the selected API voice.",
                "OpenAI API",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        if (IsActiveKittenModel())
        {
            MessageBox.Show(this,
                "Kitten TTS built-in voices do not have local sample files. Generate a short test line to preview the selected voice.",
                "Kitten TTS",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        if (IsLocalQwenCustomVoiceMode())
        {
            var selected = GetSelectedVoicePath();
            if (!TryParseQwenCustomVoiceToken(selected, out var speaker))
            {
                MessageBox.Show(this, "Select a valid CustomVoice speaker first.", "Qwen CustomVoice", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var existingSample = TryResolveExistingQwenCustomVoiceSamplePath(speaker);
            if (!string.IsNullOrWhiteSpace(existingSample))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = existingSample,
                    UseShellExecute = true
                });
                return;
            }

            var oldContent = PlaySampleButton.Content;
            PlaySampleButton.IsEnabled = false;
            PlaySampleButton.Content = "Preparing...";
            try
            {
                var samplePath = await EnsureQwenCustomVoiceSampleAsync(speaker);
                Process.Start(new ProcessStartInfo
                {
                    FileName = samplePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Failed to prepare sample for '{speaker}'.\n\n{UserMessageFormatter.FormatOperationError("Voice sample", ex.Message)}",
                    "Qwen CustomVoice",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                PlaySampleButton.Content = oldContent;
                PlaySampleButton.IsEnabled = true;
            }
            return;
        }

        var selectedVoice = GetSelectedVoicePath();
        if (string.IsNullOrWhiteSpace(selectedVoice) || !File.Exists(selectedVoice))
        {
            MessageBox.Show(this, "Select a valid voice file first.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = selectedVoice,
            UseShellExecute = true
        });
    }

    private void PortfolioHyperlink_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://elvlin.store",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Open website", ex.Message),
                "Open Link",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private async void GenerateButton_OnClick(object sender, RoutedEventArgs e)
    {
        TraceGenerate("Generate button clicked.");
        var hasActiveRow = ReconcileGenerationState();
        if (_isGenerating)
        {
            if (!hasActiveRow)
            {
                TraceGenerate("Detected stale generation lock (no active running rows). Recovering generation state.");
                _pauseGenerationRequested = false;
                _generationCts?.Dispose();
                _generationCts = null;
                _isGenerating = false;
                GenerateButton.IsEnabled = true;
            }
            else
            {
                TraceGenerate("Generate ignored because another run is active.");
                MessageBox.Show(this, "Generation is already running.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
        }

        _isGenerating = true;
        GenerateButton.IsEnabled = false;
        _pauseGenerationRequested = false;
        _generationCts?.Dispose();
        _generationCts = new CancellationTokenSource();
        var generationCt = _generationCts.Token;
        ITtsBackend? backend = null;
        try
        {
            var projectInputBackup = (_project.SourceTextFiles ?? new List<string>())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            CollectProjectFromUi();

            var selectedInputs = GetSelectedInputFilesFromUi();
            var unfinishedQueueInputs = _queueRows
                .Where(r => !string.IsNullOrWhiteSpace(r.SourcePath))
                .Where(r =>
                    !string.Equals(r.Status, "Done", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(r.Status, "Missing file", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.SourcePath)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (unfinishedQueueInputs.Count > 0)
            {
                // Preserve unfinished queue jobs across stop/restart even if input checkbox selection changed.
                selectedInputs = selectedInputs
                    .Concat(unfinishedQueueInputs)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            if (selectedInputs.Count == 0 && _inputFiles.Count == 0 && projectInputBackup.Count > 0)
            {
                selectedInputs = projectInputBackup;
                foreach (var input in selectedInputs)
                {
                    if (!_inputFiles.Any(x => string.Equals(x.FullPath, input, StringComparison.OrdinalIgnoreCase)))
                    {
                        _inputFiles.Add(new InputFileEntry(input, true));
                    }
                }

                _project.SourceTextFiles = selectedInputs.ToList();
                _project.SourceTextPath = _project.SourceTextFiles.FirstOrDefault() ?? string.Empty;
                TraceGenerate($"Recovered {selectedInputs.Count} source file(s) from project file.");
            }

            if (selectedInputs.Count == 0)
            {
                TraceGenerate("Generate blocked: no source text files.");
                MessageBox.Show(this, "Select at least one text file to generate.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                backend = CreateBackend();
                TraceGenerate($"Local device preference: {NormalizePreferDevice(_config.PreferDevice)}");
            }
            catch (Exception ex)
            {
                TraceGenerate($"Backend creation failed: {ex.Message}");
                MessageBox.Show(this,
                    UserMessageFormatter.FormatOperationError("Backend", ex.Message),
                    "Backend Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }
            var isApi = string.Equals(_config.BackendMode, "api", StringComparison.OrdinalIgnoreCase);
            var effectiveLocalSettings = isApi ? new SynthesisSettings() : GetEffectiveLocalSettings();
            if (!isApi)
            {
                effectiveLocalSettings.StylePresetKey = GetSelectedStylePresetKey();
                ApplySelectedStyleRuntimeOverrides(effectiveLocalSettings);
            }
            var mixedPrepareMode = !isApi && IsMixedVoiceToken(_project.VoicePath);
            if (mixedPrepareMode)
            {
                var missingPrepared = selectedInputs
                    .Where(p => !HasUsablePreparedScript(p))
                    .Select(Path.GetFileName)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (missingPrepared.Count > 0)
                {
                    var preview = string.Join(", ", missingPrepared.Take(5));
                    TraceGenerate($"Generate blocked: mixed mode selected but missing prepared scripts. Missing={preview}");
                    MessageBox.Show(
                        this,
                        $"Mixed mode requires prepared parts for each chapter.\nMissing prepared script: {preview}",
                        "Prepare Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }
            var requiresVoiceFile = !isApi && !mixedPrepareMode && !IsActiveKittenModel() && !IsLocalQwenCustomVoiceMode();

            if (!isApi && (string.IsNullOrWhiteSpace(_project.VoicePath) || (requiresVoiceFile && !File.Exists(_project.VoicePath))))
            {
                TraceGenerate("Generate blocked: selected voice is missing.");
                MessageBox.Show(this, "Select a valid voice.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (requiresVoiceFile && !TryEnsureProjectVoiceIsWav(out var generationPrepareError))
            {
                TraceGenerate($"Generate blocked by voice normalization: {generationPrepareError}");
                MessageBox.Show(this, generationPrepareError, "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (requiresVoiceFile && !ValidateLocalVoiceForGeneration(_project.VoicePath, out var generationVoiceError))
            {
                TraceGenerate($"Generate blocked by voice validation: {generationVoiceError}");
                MessageBox.Show(this, generationVoiceError, "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Directory.CreateDirectory(_project.OutputDir);
            TraceGenerate($"Generate started. Inputs={selectedInputs.Count}, OutputDir={_project.OutputDir}");

            var runRows = new List<(string InputPath, QueueRow Row)>(selectedInputs.Count);
            var missingInputs = new List<string>();
            foreach (var input in selectedInputs)
            {
                var row = _queueRows.FirstOrDefault(r => string.Equals(r.SourcePath, input, StringComparison.OrdinalIgnoreCase));
                if (row is null)
                {
                    row = new QueueRow
                    {
                        FileName = Path.GetFileName(input),
                        SourcePath = input
                    };
                    _queueRows.Add(row);
                }

                row.FileName = Path.GetFileName(input);
                row.SourcePath = input;
                row.Status = "Queued";
                row.ProgressLabel = "0%";
                row.ProgressValue = 0;
                row.ChunkInfo = "Queued";
                row.Eta = "Queued";
                row.OutputPath = string.Empty;
                row.IsRunning = false;
                row.IsPaused = false;
                runRows.Add((input, row));
            }

            QueueGrid.Items.Refresh();
            TraceGenerate($"Queue rows prepared: {runRows.Count}");
            if (!ReferenceEquals(QueueGrid.ItemsSource, _queueRows))
            {
                QueueGrid.ItemsSource = _queueRows;
            }
            CollectionViewSource.GetDefaultView(QueueGrid.ItemsSource)?.Refresh();
            QueueGrid.UpdateLayout();
            TraceGenerate($"Queue visual state: collection={_queueRows.Count}, gridItems={QueueGrid.Items.Count}");
            await Dispatcher.Yield(DispatcherPriority.Background);

            if (runRows.Count == 0)
            {
                MessageBox.Show(this, "No valid source files were found. Fix rows marked 'Missing file' and try again.", "Audiobook Creator", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            for (var i = 0; i < runRows.Count; i++)
            {
                var (input, row) = runRows[i];
                if (generationCt.IsCancellationRequested)
                {
                    break;
                }

                await WaitForGenerationResumeAsync(row, generationCt);

                if (!File.Exists(input))
                {
                    row.Status = "Missing file";
                    row.ProgressLabel = "0%";
                    row.ProgressValue = 0;
                    row.ChunkInfo = "-";
                    row.Eta = "Fix path";
                    row.IsRunning = false;
                    row.IsPaused = false;
                    missingInputs.Add(input);
                    TraceGenerate($"Missing file: {input}");
                    QueueGrid.Items.Refresh();
                    continue;
                }

                row.Status = "Generating narration...";
                row.ProgressLabel = "10%";
                row.ProgressValue = 10;
                row.ChunkInfo = "--";
                row.Eta = "Starting...";
                row.IsRunning = true;
                row.IsPaused = false;
                QueueGrid.Items.Refresh();
                TraceGenerate($"Processing: {row.FileName}");
                var rowStart = DateTime.UtcNow;

                var text = await File.ReadAllTextAsync(input);
                var ext = isApi
                    ? "wav"
                    : NormalizeLocalOutputExtension(effectiveLocalSettings.OutputFormat);
                var outputPath = Path.Combine(_project.OutputDir, $"{Path.GetFileNameWithoutExtension(input)}.{ext}");
                var preparedScript = TryGetPreparedScript(input);
                var usePrepared = !isApi &&
                                  IsMixedVoiceToken(_project.VoicePath) &&
                                  preparedScript is { Parts.Count: > 0 };
                TraceGenerate(
                    $"Prepared pipeline for {Path.GetFileName(input)}: " +
                    $"exists={(preparedScript is { Parts.Count: > 0 })}, " +
                    $"mixedVoice={IsMixedVoiceToken(_project.VoicePath)}, modeCustomVoice={IsLocalQwenCustomVoiceMode()}, isApi={isApi}, usePrepared={usePrepared}");
                if (usePrepared && preparedScript is not null)
                {
                    var stale = !string.Equals(preparedScript.SourceSignature ?? string.Empty, BuildSourceSignature(input), StringComparison.Ordinal);
                    if (stale)
                    {
                        if (IsMixedVoiceToken(_project.VoicePath))
                        {
                            TraceGenerate($"Prepared script for '{Path.GetFileName(input)}' is stale, but mixed mode requires prepared pipeline. Continuing with prepared parts.");
                        }
                        else
                        {
                            var choice = MessageBox.Show(
                                this,
                                $"Prepared script for '{Path.GetFileName(input)}' looks stale (source file changed).\n\nYes = Use prepared script anyway\nNo = Use raw text this run",
                                "Prepared Script Stale",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Question);
                            usePrepared = choice == MessageBoxResult.Yes;
                        }
                    }
                }

                try
                {
                    IReadOnlyList<SubtitleChunkTiming>? subtitleChunkTimings = null;
                    EnhanceResult? enhanceResult = null;
                    if (isApi)
                    {
                        row.ProgressLabel = "35%";
                        row.ProgressValue = 35;
                        row.ChunkInfo = "--";
                        row.Eta = "Running...";
                        QueueGrid.Items.Refresh();
                        await backend.SynthesizeAsync(CreateLocalTtsRequest(
                            text,
                            _project.VoicePath,
                            outputPath,
                            (float)SpeedSlider.Value,
                            effectiveLocalSettings), generationCt);
                    }
                    else
                    {
                        if (usePrepared && preparedScript is not null)
                        {
                            subtitleChunkTimings = await GeneratePreparedScriptLocalAsync(
                                backend,
                                preparedScript,
                                outputPath,
                                (float)SpeedSlider.Value,
                                effectiveLocalSettings,
                                generationCt,
                                (doneParts, totalParts) =>
                                {
                                    var total = Math.Max(totalParts, 1);
                                    var boundedDone = Math.Clamp(doneParts, 0, total);
                                    var progress = boundedDone == 0
                                        ? 10
                                        : 10 + (int)Math.Round((boundedDone * 85.0) / total);
                                    progress = Math.Clamp(progress, 10, 99);
                                    row.ProgressLabel = $"{progress}%";
                                    row.ProgressValue = progress;
                                    if (boundedDone == 0)
                                    {
                                        row.ChunkInfo = $"Part 0/{total}";
                                        row.Eta = "Starting...";
                                    }
                                    else
                                    {
                                        var elapsed = DateTime.UtcNow - rowStart;
                                        var remainingSeconds = (elapsed.TotalSeconds / boundedDone) * (total - boundedDone);
                                        var eta = TimeSpan.FromSeconds(Math.Max(0, remainingSeconds));
                                        row.ChunkInfo = $"Part {boundedDone}/{total}";
                                        row.Eta = FormatDuration(eta);
                                    }
                                    QueueGrid.Items.Refresh();
                                });
                        }
                        else
                        {
                            subtitleChunkTimings = await GenerateLocalWithChunkingAsync(
                                backend,
                                text,
                                _project.VoicePath,
                                outputPath,
                                (float)SpeedSlider.Value,
                                effectiveLocalSettings,
                                generationCt,
                                (doneChunks, totalChunks) =>
                                {
                                    var total = Math.Max(totalChunks, 1);
                                    var boundedDone = Math.Clamp(doneChunks, 0, total);
                                    var progress = boundedDone == 0
                                        ? 10
                                        : 10 + (int)Math.Round((boundedDone * 85.0) / total);
                                    progress = Math.Clamp(progress, 10, 99);
                                    row.ProgressLabel = $"{progress}%";
                                    row.ProgressValue = progress;
                                    if (boundedDone == 0)
                                    {
                                        row.ChunkInfo = $"Chunk 0/{total}";
                                        row.Eta = "Starting...";
                                    }
                                    else
                                    {
                                        var elapsed = DateTime.UtcNow - rowStart;
                                        var remainingSeconds = (elapsed.TotalSeconds / boundedDone) * (total - boundedDone);
                                        var eta = TimeSpan.FromSeconds(Math.Max(0, remainingSeconds));
                                        row.ChunkInfo = $"Chunk {boundedDone}/{total}";
                                        row.Eta = FormatDuration(eta);
                                    }
                                    QueueGrid.Items.Refresh();
                                });
                        }
                    }

                    if (IsEnhanceEnabledForRun())
                    {
                        var enhanceProgress = new Progress<AudioEnhanceProgress>(p =>
                        {
                            var total = Math.Max(p.Total, 1);
                            var done = Math.Clamp(p.Completed, 0, total);
                            var progress = 90 + (int)Math.Round((done * 9.0) / total);
                            row.Status = p.Stage;
                            row.ProgressLabel = $"{Math.Clamp(progress, 90, 99)}%";
                            row.ProgressValue = Math.Clamp(progress, 90, 99);
                            row.ChunkInfo = total > 1 ? $"{done}/{total}" : p.Stage;
                            var elapsed = DateTime.UtcNow - rowStart;
                            row.Eta = FormatDuration(elapsed);
                            QueueGrid.Items.Refresh();
                        });

                        enhanceResult = await RunAudioEnhanceForChapterAsync(
                            text,
                            outputPath,
                            subtitleChunkTimings,
                            preparedScript,
                            generationCt,
                            enhanceProgress);

                        if (!enhanceResult.Success)
                        {
                            TraceGenerate($"Enhance warning for {row.FileName}: {enhanceResult.Message}");
                        }
                    }

                    await GenerateSubtitlesIfEnabledAsync(text, outputPath, subtitleChunkTimings);
                    var noEnhanceCues = enhanceResult is { Success: true } &&
                                        !string.IsNullOrWhiteSpace(enhanceResult.Message) &&
                                        enhanceResult.Message.Contains("narration-only", StringComparison.OrdinalIgnoreCase);
                    row.Status = noEnhanceCues ? "Done (no cues)" : "Done";
                    row.ProgressLabel = "100%";
                    row.ProgressValue = 100;
                    row.ChunkInfo = enhanceResult is { Success: false }
                        ? "Done (enhance warning)"
                        : noEnhanceCues
                            ? "Narration only"
                            : "Done";
                    row.Eta = FormatDuration(DateTime.UtcNow - rowStart);
                    row.OutputPath = outputPath;
                    row.IsRunning = false;
                    row.IsPaused = false;
                    AddHistoryEntryForCompletedRow(row, isApi, effectiveLocalSettings);
                    TryAutoClearCompletedQueueRow(row);
                    if (backend is ChatterboxOnnxBackend chatterboxBackend)
                        TraceGenerate($"ONNX provider used: {chatterboxBackend.ActiveExecutionProvider}");
                    else if (backend is Qwen3OnnxSplitBackend qwenSplitBackend)
                        TraceGenerate($"ONNX provider used: {qwenSplitBackend.ActiveExecutionProvider}");
                    TraceGenerate($"Generated: {row.FileName}");
                }
                catch (OperationCanceledException)
                {
                    row.Status = "Stopped";
                    row.ChunkInfo = "Stopped";
                    row.Eta = "Stopped";
                    row.IsRunning = false;
                    row.IsPaused = false;
                    TraceGenerate($"Stopped: {row.FileName}");
                    QueueGrid.Items.Refresh();
                    break;
                }
                catch (Exception ex)
                {
                    row.Status = "Failed";
                    row.ProgressLabel = "0%";
                    row.ProgressValue = 0;
                    row.ChunkInfo = "Failed";
                    row.Eta = "Failed";
                    row.IsRunning = false;
                    row.IsPaused = false;
                    TraceGenerate($"Failed: {row.FileName} -> {ex.Message}");
                    MessageBox.Show(this, NormalizeGenerationError(ex.Message), "Generation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                QueueGrid.Items.Refresh();
            }

            if (generationCt.IsCancellationRequested)
            {
                foreach (var (_, row) in runRows.Where(r => r.Row.Status == "Queued"))
                {
                    row.Status = "Stopped";
                    row.ChunkInfo = "Stopped";
                    row.Eta = "Stopped";
                    row.IsRunning = false;
                    row.IsPaused = false;
                }
            }

            if (IsAutoClearCompletedEnabled())
            {
                var doneRows = _queueRows.Where(r => IsCompletedQueueStatus(r.Status)).ToList();
                foreach (var row in doneRows)
                {
                    EnsureHistoryEntryForRestoredDoneRow(row);
                    _queueRows.Remove(row);
                }
            }

            if (_config.AutoRemoveCompletedInputFiles)
            {
                var completedInputPaths = runRows
                    .Where(r => IsCompletedQueueStatus(r.Row.Status))
                    .Select(r => r.InputPath)
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (completedInputPaths.Count > 0)
                {
                    for (var idx = _inputFiles.Count - 1; idx >= 0; idx--)
                    {
                        var fullPath = _inputFiles[idx].FullPath;
                        if (completedInputPaths.Any(p => string.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase)))
                        {
                            _inputFiles.RemoveAt(idx);
                        }
                    }
                }
            }

            var completed = runRows.Count(r => IsCompletedQueueStatus(r.Row.Status));
            var failed = runRows.Count(r => string.Equals(r.Row.Status, "Failed", StringComparison.OrdinalIgnoreCase));
            var stopped = runRows.Count(r => string.Equals(r.Row.Status, "Stopped", StringComparison.OrdinalIgnoreCase));
            if (!generationCt.IsCancellationRequested)
            {
                MessageBox.Show(
                    this,
                    $"Generation finished. Done: {completed}, Failed: {failed}, Stopped: {stopped}, Missing: {missingInputs.Count}.",
                    "Audiobook Creator",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            TraceGenerate($"Generation finished. Done={completed}, Failed={failed}, Stopped={stopped}, Missing={missingInputs.Count}");
            _ = AutoSaveCurrentProjectAsync();
        }
        catch (Exception ex)
        {
            TraceGenerate($"Unhandled generate exception: {ex.Message}");
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Generation", ex.Message),
                "Audiobook Creator",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            if (backend is IDisposable backendDisposable)
            {
                backendDisposable.Dispose();
            }
            _generationCts?.Dispose();
            _generationCts = null;
            _pauseGenerationRequested = false;
            GenerateButton.IsEnabled = true;
            _isGenerating = false;
        }
    }

    private async Task WaitForGenerationResumeAsync(QueueRow row, CancellationToken ct)
    {
        while (_pauseGenerationRequested && !ct.IsCancellationRequested)
        {
            row.Status = "Paused";
            row.ChunkInfo = "Paused";
            row.Eta = "Paused";
            row.IsPaused = true;
            row.IsRunning = true;
            QueueGrid.Items.Refresh();
            await Task.Delay(120, ct);
        }

        if (!ct.IsCancellationRequested && row.IsPaused)
        {
            row.IsPaused = false;
                        row.Status = "Generating narration...";
            row.ChunkInfo = "--";
            row.Eta = "Resuming...";
            QueueGrid.Items.Refresh();
        }
    }

    private static string FormatDuration(TimeSpan value)
    {
        if (value.TotalHours >= 1)
        {
            return $"{(int)value.TotalHours:00}:{value.Minutes:00}:{value.Seconds:00}";
        }

        return $"{value.Minutes:00}:{value.Seconds:00}";
    }

    private bool IsAutoClearCompletedEnabled() => AutoClearCheckBox?.IsChecked == true;

    private static bool IsCompletedQueueStatus(string? status)
    {
        return !string.IsNullOrWhiteSpace(status) &&
               status.StartsWith("Done", StringComparison.OrdinalIgnoreCase);
    }

    private void TryAutoClearCompletedQueueRow(QueueRow row)
    {
        if (!IsAutoClearCompletedEnabled())
        {
            return;
        }

        if (!string.Equals(row.Status, "Done", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        EnsureHistoryEntryForRestoredDoneRow(row);
        _queueRows.Remove(row);
    }

    private bool ReconcileGenerationState()
    {
        var realActive = false;
        var staleRunningRows = new List<QueueRow>();

        foreach (var row in _queueRows)
        {
            if (!row.IsRunning)
            {
                continue;
            }

            var status = row.Status ?? string.Empty;
            var isRealRunning =
                status.StartsWith("Generating", StringComparison.OrdinalIgnoreCase) ||
                status.StartsWith("Paused", StringComparison.OrdinalIgnoreCase) ||
                status.StartsWith("Stopping", StringComparison.OrdinalIgnoreCase);

            if (isRealRunning)
            {
                realActive = true;
                continue;
            }

            staleRunningRows.Add(row);
        }

        if (staleRunningRows.Count > 0)
        {
            foreach (var row in staleRunningRows)
            {
                row.IsRunning = false;
                row.IsPaused = false;
                if (string.IsNullOrWhiteSpace(row.Status))
                {
                    row.Status = "Queued";
                }
            }

            QueueGrid.Items.Refresh();
            TraceGenerate($"Recovered {staleRunningRows.Count} stale running row(s).");
        }

        if (_isGenerating && !realActive)
        {
            _pauseGenerationRequested = false;
            _generationCts?.Dispose();
            _generationCts = null;
            _isGenerating = false;
            if (GenerateButton is not null)
            {
                GenerateButton.IsEnabled = true;
            }

            TraceGenerate("Recovered stale generation lock (no real active row).");
        }

        return realActive;
    }

    private void ClearFinishedButton_OnClick(object sender, RoutedEventArgs e)
    {
        var doneRows = _queueRows
            .Where(r => string.Equals(r.Status, "Done", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (doneRows.Count == 0)
        {
            return;
        }

        foreach (var row in doneRows)
        {
            EnsureHistoryEntryForRestoredDoneRow(row);
            _queueRows.Remove(row);
        }

        QueueGrid.Items.Refresh();
        _ = AutoSaveCurrentProjectAsync();
    }

    private void ClearQueueButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_queueRows.Count == 0)
        {
            return;
        }

        var hasActiveRow = ReconcileGenerationState();

        if (_isGenerating || hasActiveRow)
        {
            MessageBox.Show(
                this,
                "Stop the current generation before clearing the entire queue.",
                "Clear Queue",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var confirm = MessageBox.Show(
            this,
            "Remove all rows from Job Queue?",
            "Clear Queue",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        _queueRows.Clear();
        QueueGrid.Items.Refresh();
        _ = AutoSaveCurrentProjectAsync();
    }

    private void AddHistoryEntryForCompletedRow(QueueRow row, bool isApi, SynthesisSettings? settingsOverride = null)
    {
        if (row is null || !string.Equals(row.Status, "Done", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var voiceLabel = ResolveHistoryVoiceLabel();
        var deviceLabel = ResolveHistoryDeviceLabel(isApi);
        var modelLabel = ResolveHistoryModelLabel(isApi, settingsOverride);

        _historyRows.Insert(0, new HistoryRow
        {
            FileName = row.FileName ?? Path.GetFileName(row.SourcePath ?? string.Empty),
            SourcePath = row.SourcePath ?? string.Empty,
            OutputPath = row.OutputPath ?? string.Empty,
            ModelLabel = modelLabel,
            DeviceLabel = deviceLabel,
            VoiceLabel = voiceLabel,
            CompletedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            DurationLabel = string.IsNullOrWhiteSpace(row.Eta) ? string.Empty : row.Eta
        });

        while (_historyRows.Count > 1000)
        {
            _historyRows.RemoveAt(_historyRows.Count - 1);
        }
    }

    private void BackfillHistoryFromRestoredDoneQueueRows()
    {
        foreach (var row in _queueRows.Where(r => string.Equals(r.Status, "Done", StringComparison.OrdinalIgnoreCase)))
        {
            EnsureHistoryEntryForRestoredDoneRow(row);
        }
    }

    private void EnsureHistoryEntryForRestoredDoneRow(QueueRow row)
    {
        if (row is null || !string.Equals(row.Status, "Done", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var outputPath = row.OutputPath ?? string.Empty;
        var sourcePath = row.SourcePath ?? string.Empty;
        var fileName = row.FileName ?? Path.GetFileName(sourcePath);

        var alreadyExists = _historyRows.Any(h =>
            (!string.IsNullOrWhiteSpace(outputPath) &&
             string.Equals(h.OutputPath, outputPath, StringComparison.OrdinalIgnoreCase)) ||
            (string.Equals(h.FileName, fileName, StringComparison.OrdinalIgnoreCase) &&
             string.Equals(h.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase)));

        if (alreadyExists)
        {
            return;
        }

        var completedAt = string.Empty;
        if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath))
        {
            completedAt = File.GetLastWriteTime(outputPath).ToString("yyyy-MM-dd HH:mm:ss");
        }

        _historyRows.Add(new HistoryRow
        {
            FileName = fileName,
            SourcePath = sourcePath,
            OutputPath = outputPath,
            ModelLabel = "Restored (unknown)",
            DeviceLabel = "-",
            VoiceLabel = "-",
            CompletedAt = completedAt,
            DurationLabel = row.Eta ?? string.Empty
        });

        while (_historyRows.Count > 1000)
        {
            _historyRows.RemoveAt(_historyRows.Count - 1);
        }
    }

    private string ResolveHistoryVoiceLabel()
    {
        if (IsApiOpenAiMode())
        {
            if (TryParseOpenAiVoiceToken(GetSelectedVoicePath(), out var openAiVoice))
            {
                return openAiVoice;
            }

            return string.IsNullOrWhiteSpace(_config.ApiVoice) ? "-" : _config.ApiVoice.Trim();
        }

        if (IsMixedVoiceToken(GetSelectedVoicePath() ?? string.Empty))
        {
            return "Mixed";
        }

        if (IsActiveKittenModel() && TryParseKittenVoiceToken(GetSelectedVoicePath() ?? string.Empty, out var kittenVoice))
        {
            return kittenVoice;
        }
        if (IsLocalQwenCustomVoiceMode() && TryParseQwenCustomVoiceToken(GetSelectedVoicePath() ?? string.Empty, out var qwenSpeaker))
        {
            return qwenSpeaker;
        }

        var voicePath = GetSelectedVoicePath();
        return string.IsNullOrWhiteSpace(voicePath) ? "-" : Path.GetFileName(voicePath);
    }

    private string ResolveHistoryDeviceLabel(bool isApi)
    {
        return isApi ? "api" : NormalizePreferDevice(_config.PreferDevice);
    }

    private string ResolveHistoryModelLabel(bool isApi, SynthesisSettings? settingsOverride)
    {
        if (isApi)
        {
            var provider = string.IsNullOrWhiteSpace(_config.ApiProvider) ? "api" : _config.ApiProvider.Trim();
            var model = string.IsNullOrWhiteSpace(_config.ApiModelId) ? "default" : _config.ApiModelId.Trim();
            return $"API | {provider} | {model}";
        }

        var preset = (_config.LocalModelPreset ?? string.Empty).Trim();
        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        var profile = string.IsNullOrWhiteSpace(preset) ? "local" : preset;
        var style = (settingsOverride?.StylePresetKey ?? GetSelectedStylePresetKey() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(style))
        {
            return string.IsNullOrWhiteSpace(repo) ? profile : $"{profile} | {repo}";
        }

        return string.IsNullOrWhiteSpace(repo)
            ? $"{profile} | style:{style}"
            : $"{profile} | {repo} | style:{style}";
    }

    private void SpeedSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (SpeedLabel is null)
        {
            return;
        }
        SpeedLabel.Text = $"{SpeedSlider.Value:0.0}x";
    }

    private async void StylePresetCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing || _isUpdatingStylePreset || _isRefreshingProjectSelector)
        {
            return;
        }

        var key = GetSelectedStylePresetKey();
        if (string.IsNullOrWhiteSpace(key))
        {
            return;
        }

        _project.Settings ??= new SynthesisSettings();
        _project.Settings.StylePresetKey = key;

        try
        {
            var seed = _project.Settings is null ? new SynthesisSettings() : CloneSynthesisSettings(_project.Settings);
            seed.StylePresetKey = key;
            OverwriteCurrentModelSettingsEverywhere(seed);
            await _configStore.SaveAsync(_config);
        }
        catch
        {
            // Non-fatal persistence failure.
        }

        await AutoSaveCurrentProjectAsync();
    }

    private async void SettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var dlg = new SettingsWindow(_config, GetSelectedVoicePath())
            {
                Owner = this
            };
            var result = dlg.ShowDialog();
            if (result != true || dlg.UpdatedConfig is null)
            {
                return;
            }

            _config = dlg.UpdatedConfig;
            await _configStore.SaveAsync(_config);
            if (TryGetBestActiveModelProfile(out var syncedProfile, out _))
            {
                _project.Settings = CloneSynthesisSettings(syncedProfile);
                await AutoSaveCurrentProjectAsync();
            }
            if (string.IsNullOrWhiteSpace(_project.OutputDir))
            {
                _project.OutputDir = ResolveProjectOutputDir(_config.DefaultOutputDir, _project.Name);
                OutputFolderTextBox.Text = _project.OutputDir;
            }
            LoadVoiceList();
            RefreshStylePresetOptions();
            RefreshModelSummary();
            RefreshModelControlSummary();
            RefreshVoiceDesignReadiness();
            SyncMainDeviceComboFromConfig();
            RenderSystemInfo();
            RefreshInputFilePreparedState();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"Failed to open settings:\n{ex.Message}",
                "Settings Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }


    private async void MainDeviceCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        var selected = (MainDeviceCombo.SelectedItem as ComboBoxItem)?.Content?.ToString();
        if (string.IsNullOrWhiteSpace(selected))
        {
            selected = MainDeviceCombo.Text;
        }

        var normalized = NormalizePreferDevice(selected);
        if (string.Equals(_config.PreferDevice, normalized, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (string.Equals(normalized, "gpu", StringComparison.OrdinalIgnoreCase))
        {
            var profile = SystemProbe.Detect();
            if (!profile.CudaAvailable)
            {
                var hint = string.IsNullOrWhiteSpace(profile.CudaInstallHint)
                    ? "Install CUDA runtime: https://developer.nvidia.com/cuda-downloads"
                    : profile.CudaInstallHint;
                MessageBox.Show(
                    this,
                    $"CUDA is not ready for native ONNX GPU.\n\nStatus: {profile.CudaStatus}\n\n{hint}",
                    "GPU Setup Needed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        _config.PreferDevice = normalized;
        RefreshModelControlSummary();
        try
        {
            await _configStore.SaveAsync(_config);
        }
        catch
        {
            // Non-fatal if config save fails.
        }
    }

    private void ModelControlsButton_OnClick(object sender, RoutedEventArgs e)
    {
        // Use the same effective settings object that generation uses, so UI reflects runtime behavior.
        var settings = GetEffectiveLocalSettings();
        var isQwenModel = IsQwenRepo(_config.ModelRepoId);
        var pauseProfile = ResolveChunkProfile();
        var (defaultClauseRange, defaultSentenceRange, defaultParagraphRange, defaultEllipsisRange) = GetDefaultPauseRanges(pauseProfile);
        var dlg = new Window
        {
            Title = "Model Controls",
            Width = 560,
            Height = 720,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.CanResize,
            Background = Brushes.White
        };

        var root = new Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = new TextBlock
        {
            Text = "Local Model Behavior",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold
        };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        var panel = new Grid { Margin = new Thickness(0, 12, 0, 0) };
        panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
        panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var i = 0; i < 23; i++)
        {
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        var chunkModeLabel = new TextBlock { Text = "Chunk Mode", VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(chunkModeLabel, 0);
        panel.Children.Add(chunkModeLabel);
        var chunkModeCombo = new ComboBox { Height = 32 };
        chunkModeCombo.Items.Add("auto");
        chunkModeCombo.Items.Add("manual");
        chunkModeCombo.Text = string.IsNullOrWhiteSpace(settings.ChunkMode) ? "auto" : settings.ChunkMode;
        Grid.SetRow(chunkModeCombo, 0);
        Grid.SetColumn(chunkModeCombo, 1);
        panel.Children.Add(chunkModeCombo);

        var minCharsLabel = new TextBlock { Text = "Min Chars", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(minCharsLabel, 1);
        panel.Children.Add(minCharsLabel);
        var minCharsBox = new TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = settings.MinChars.ToString() };
        Grid.SetRow(minCharsBox, 1);
        Grid.SetColumn(minCharsBox, 1);
        panel.Children.Add(minCharsBox);

        var maxCharsLabel = new TextBlock { Text = "Max Chars (Auto)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(maxCharsLabel, 2);
        panel.Children.Add(maxCharsLabel);
        var maxCharsBox = new TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = settings.MaxChars.ToString() };
        Grid.SetRow(maxCharsBox, 2);
        Grid.SetColumn(maxCharsBox, 1);
        panel.Children.Add(maxCharsBox);

        var manualCharsLabel = new TextBlock { Text = "Max Chars (Manual)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(manualCharsLabel, 3);
        panel.Children.Add(manualCharsLabel);
        var manualCharsBox = new TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = settings.ManualMaxChars.ToString() };
        Grid.SetRow(manualCharsBox, 3);
        Grid.SetColumn(manualCharsBox, 1);
        panel.Children.Add(manualCharsBox);

        var chunkPauseLabel = new TextBlock { Text = "Chunk Pause (ms)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(chunkPauseLabel, 4);
        panel.Children.Add(chunkPauseLabel);
        var chunkPauseBox = new TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = settings.ChunkPauseMs.ToString() };
        Grid.SetRow(chunkPauseBox, 4);
        Grid.SetColumn(chunkPauseBox, 1);
        panel.Children.Add(chunkPauseBox);

        var paragraphPauseLabel = new TextBlock { Text = "Paragraph Pause (ms)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(paragraphPauseLabel, 5);
        panel.Children.Add(paragraphPauseLabel);
        var paragraphPauseBox = new TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = settings.ParagraphPauseMs.ToString() };
        Grid.SetRow(paragraphPauseBox, 5);
        Grid.SetColumn(paragraphPauseBox, 1);
        panel.Children.Add(paragraphPauseBox);

        var effectiveClauseRange = ResolvePauseRange(settings.ClausePauseMinMs, settings.ClausePauseMaxMs, defaultClauseRange);
        var effectiveSentenceRange = ResolvePauseRange(settings.SentencePauseMinMs, settings.SentencePauseMaxMs, defaultSentenceRange);
        var effectiveEllipsisRange = ResolvePauseRange(settings.EllipsisPauseMinMs, settings.EllipsisPauseMaxMs, defaultEllipsisRange);
        var effectiveParagraphRange = ResolvePauseRange(settings.ParagraphPauseMinMs, settings.ParagraphPauseMaxMs, defaultParagraphRange);

        AddPauseRangeRow(panel, 6, "Clause Pause Range", effectiveClauseRange, out var clausePauseMinBox, out var clausePauseMaxBox);
        AddPauseRangeRow(panel, 7, "Sentence Pause Range", effectiveSentenceRange, out var sentencePauseMinBox, out var sentencePauseMaxBox);
        AddPauseRangeRow(panel, 8, "Ellipsis Pause Range", effectiveEllipsisRange, out var ellipsisPauseMinBox, out var ellipsisPauseMaxBox);
        AddPauseRangeRow(panel, 9, "Paragraph Pause Range", effectiveParagraphRange, out var paragraphPauseMinBox, out var paragraphPauseMaxBox);

        var outFmtLabel = new TextBlock { Text = "Output Format", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(outFmtLabel, 10);
        panel.Children.Add(outFmtLabel);
        var outFmtCombo = new ComboBox { Height = 32, Margin = new Thickness(0, 10, 0, 0) };
        outFmtCombo.Items.Add("wav");
        outFmtCombo.Items.Add("mp3");
        outFmtCombo.Text = NormalizeLocalOutputExtension(settings.OutputFormat);
        Grid.SetRow(outFmtCombo, 10);
        Grid.SetColumn(outFmtCombo, 1);
        panel.Children.Add(outFmtCombo);

        var protectBracket = new CheckBox
        {
            Content = "Preserve [ ... ] directives when chunking (Qwen-style)",
            Margin = new Thickness(0, 12, 0, 0),
            IsChecked = settings.ProtectBracketDirectives
        };
        Grid.SetRow(protectBracket, 11);
        Grid.SetColumnSpan(protectBracket, 2);
        panel.Children.Add(protectBracket);

        var hintPanel = new Grid { Margin = new Thickness(0, 10, 0, 0) };
        hintPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
        hintPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var hintLabel = new TextBlock { Text = "Qwen Hint (optional)", VerticalAlignment = VerticalAlignment.Center };
        var hintBox = new TextBox
        {
            Height = 32,
            Text = settings.LocalInstructionHint ?? string.Empty,
            ToolTip = "If set and model repo is Qwen3-TTS, each chunk is prefixed like [hint] ..."
        };
        Grid.SetColumn(hintBox, 1);
        hintPanel.Children.Add(hintLabel);
        hintPanel.Children.Add(hintBox);
        Grid.SetRow(hintPanel, 12);
        Grid.SetColumnSpan(hintPanel, 2);
        panel.Children.Add(hintPanel);

        var qwenStablePresetCheck = new CheckBox
        {
            Content = "Qwen Audiobook Stable Preset (recommended)",
            Margin = new Thickness(0, 12, 0, 0),
            IsChecked = settings.QwenStableAudiobookPreset,
            IsEnabled = isQwenModel
        };
        Grid.SetRow(qwenStablePresetCheck, 13);
        Grid.SetColumnSpan(qwenStablePresetCheck, 2);
        panel.Children.Add(qwenStablePresetCheck);

        var qwenUseRefTextCheck = new CheckBox
        {
            Content = "Qwen: Use ref_text (ICL clone mode)",
            Margin = new Thickness(0, 8, 0, 0),
            IsChecked = settings.QwenUseRefText,
            IsEnabled = isQwenModel
        };
        Grid.SetRow(qwenUseRefTextCheck, 14);
        Grid.SetColumnSpan(qwenUseRefTextCheck, 2);
        panel.Children.Add(qwenUseRefTextCheck);

        var qwenRefTextWarning = new TextBlock
        {
            Text = "Turn OFF if voice sample transcript leaks into output. OFF uses x-vector-only clone (safer for long generation).",
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            Visibility = isQwenModel ? Visibility.Visible : Visibility.Collapsed
        };
        Grid.SetRow(qwenRefTextWarning, 15);
        Grid.SetColumnSpan(qwenRefTextWarning, 2);
        panel.Children.Add(qwenRefTextWarning);

        var qwenDoSampleCheck = new CheckBox
        {
            Content = "Qwen do_sample (turn off for most stable narration)",
            Margin = new Thickness(0, 8, 0, 0),
            IsChecked = settings.QwenDoSample,
            IsEnabled = isQwenModel
        };
        Grid.SetRow(qwenDoSampleCheck, 16);
        Grid.SetColumnSpan(qwenDoSampleCheck, 2);
        panel.Children.Add(qwenDoSampleCheck);

        AddQwenNumberField(panel, 17, "Qwen Temperature", settings.QwenTemperature.ToString("0.##", CultureInfo.InvariantCulture), isQwenModel, out var qwenTemperatureBox);
        AddQwenNumberField(panel, 18, "Qwen Top-K", settings.QwenTopK.ToString(CultureInfo.InvariantCulture), isQwenModel, out var qwenTopKBox);
        AddQwenNumberField(panel, 19, "Qwen Top-P", settings.QwenTopP.ToString("0.##", CultureInfo.InvariantCulture), isQwenModel, out var qwenTopPBox);
        AddQwenNumberField(panel, 20, "Qwen Repetition Penalty", settings.QwenRepetitionPenalty.ToString("0.##", CultureInfo.InvariantCulture), isQwenModel, out var qwenRepetitionBox);

        var qwenRetryCheck = new CheckBox
        {
            Content = "Qwen auto-retry suspicious chunks",
            Margin = new Thickness(0, 8, 0, 0),
            IsChecked = settings.QwenAutoRetryBadChunks,
            IsEnabled = isQwenModel
        };
        Grid.SetRow(qwenRetryCheck, 21);
        Grid.SetColumnSpan(qwenRetryCheck, 2);
        panel.Children.Add(qwenRetryCheck);
        AddQwenNumberField(panel, 22, "Qwen Retry Count", settings.QwenBadChunkRetryCount.ToString(CultureInfo.InvariantCulture), isQwenModel, out var qwenRetryCountBox);

        Grid.SetRow(panel, 1);
        root.Children.Add(panel);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        var cancelBtn = new Button { Content = "Cancel", Width = 90, Height = 32, Margin = new Thickness(0, 0, 8, 0) };
        cancelBtn.Click += (_, _) => { dlg.DialogResult = false; dlg.Close(); };
        var applyBtn = new Button { Content = "Apply", Width = 100, Height = 34 };
        applyBtn.Click += (_, _) =>
        {
            settings.ChunkMode = (chunkModeCombo.Text ?? "auto").Trim().ToLowerInvariant() == "manual" ? "manual" : "auto";
            settings.MinChars = Math.Clamp(ParseIntOrFallback(minCharsBox.Text, settings.MinChars), 40, 3000);
            settings.MaxChars = Math.Clamp(ParseIntOrFallback(maxCharsBox.Text, settings.MaxChars), 80, 4000);
            settings.ManualMaxChars = Math.Clamp(ParseIntOrFallback(manualCharsBox.Text, settings.ManualMaxChars), 80, 6000);
            settings.ChunkPauseMs = Math.Clamp(ParseIntOrFallback(chunkPauseBox.Text, settings.ChunkPauseMs), 0, 5000);
            settings.ParagraphPauseMs = Math.Clamp(ParseIntOrFallback(paragraphPauseBox.Text, settings.ParagraphPauseMs), 0, 10000);
            settings.ClausePauseMinMs = Math.Clamp(ParseIntOrFallback(clausePauseMinBox.Text, effectiveClauseRange.Min), 0, 5000);
            settings.ClausePauseMaxMs = Math.Clamp(ParseIntOrFallback(clausePauseMaxBox.Text, effectiveClauseRange.Max), 0, 5000);
            settings.SentencePauseMinMs = Math.Clamp(ParseIntOrFallback(sentencePauseMinBox.Text, effectiveSentenceRange.Min), 0, 10000);
            settings.SentencePauseMaxMs = Math.Clamp(ParseIntOrFallback(sentencePauseMaxBox.Text, effectiveSentenceRange.Max), 0, 10000);
            settings.EllipsisPauseMinMs = Math.Clamp(ParseIntOrFallback(ellipsisPauseMinBox.Text, effectiveEllipsisRange.Min), 0, 10000);
            settings.EllipsisPauseMaxMs = Math.Clamp(ParseIntOrFallback(ellipsisPauseMaxBox.Text, effectiveEllipsisRange.Max), 0, 10000);
            settings.ParagraphPauseMinMs = Math.Clamp(ParseIntOrFallback(paragraphPauseMinBox.Text, effectiveParagraphRange.Min), 0, 15000);
            settings.ParagraphPauseMaxMs = Math.Clamp(ParseIntOrFallback(paragraphPauseMaxBox.Text, effectiveParagraphRange.Max), 0, 15000);
            settings.OutputFormat = NormalizeLocalOutputExtension(outFmtCombo.Text);
            settings.ProtectBracketDirectives = protectBracket.IsChecked == true;
            settings.LocalInstructionHint = (hintBox.Text ?? string.Empty).Trim();
            settings.QwenStableAudiobookPreset = qwenStablePresetCheck.IsChecked == true;
            settings.QwenUseRefText = qwenUseRefTextCheck.IsChecked != false;
            settings.QwenDoSample = qwenDoSampleCheck.IsChecked == true;
            settings.QwenTemperature = Math.Clamp(ParseDoubleOrFallback(qwenTemperatureBox.Text, settings.QwenTemperature), 0.05, 2.0);
            settings.QwenTopK = Math.Clamp(ParseIntOrFallback(qwenTopKBox.Text, settings.QwenTopK), 1, 200);
            settings.QwenTopP = Math.Clamp(ParseDoubleOrFallback(qwenTopPBox.Text, settings.QwenTopP), 0.05, 1.0);
            settings.QwenRepetitionPenalty = Math.Clamp(ParseDoubleOrFallback(qwenRepetitionBox.Text, settings.QwenRepetitionPenalty), 1.0, 2.0);
            settings.QwenAutoRetryBadChunks = qwenRetryCheck.IsChecked == true;
            settings.QwenBadChunkRetryCount = Math.Clamp(ParseIntOrFallback(qwenRetryCountBox.Text, settings.QwenBadChunkRetryCount), 0, 4);
            if (settings.MinChars > settings.MaxChars)
            {
                settings.MinChars = settings.MaxChars;
            }
            NormalizePauseRangeOrder(settings);
            if (isQwenModel)
            {
                ApplyQwenStableAudiobookPreset(settings);
            }
            OverwriteCurrentModelSettingsEverywhere(settings);
            try
            {
                Task.Run(() => _configStore.SaveAsync(_config)).GetAwaiter().GetResult();
            }
            catch
            {
                // Keep dialog flow even if config profile save fails.
            }
            dlg.DialogResult = true;
            dlg.Close();
        };
        buttons.Children.Add(cancelBtn);
        buttons.Children.Add(applyBtn);
        Grid.SetRow(buttons, 2);
        root.Children.Add(buttons);

        dlg.Content = root;
        if (dlg.ShowDialog() == true)
        {
            RefreshModelControlSummary();
            _ = AutoSaveCurrentProjectAsync();
        }
    }

    private void CloseButton_OnClick(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private ITtsBackend CreateBackend()
    {
        var backendMode = (_config.BackendMode ?? "auto").Trim().ToLowerInvariant();
        if (backendMode == "api")
        {
            return new ApiBackend(new InferenceOptions
            {
                Provider = _config.ApiProvider,
                ApiKey = ResolveApiKeyForProvider(_config.ApiProvider),
                ModelId = _config.ApiModelId,
                Voice = _config.ApiVoice,
                BaseUrl = _config.ApiBaseUrl,
                LanguageType = _config.ApiLanguageType,
                VoiceDesignTargetModel = _config.ApiVoiceDesignTargetModel
            });
        }

        var repo = string.IsNullOrWhiteSpace(_config.ModelRepoId)
            ? "onnx-community/chatterbox-ONNX"
            : _config.ModelRepoId.Trim();
        var isChatterboxOnnx = repo.Contains("chatterbox", StringComparison.OrdinalIgnoreCase) &&
                               repo.Contains("onnx", StringComparison.OrdinalIgnoreCase);
        var usePythonBackend = UsesPythonLocalBackend();
        var isQwen3Tts = repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase);
        var isQwenOnnxDll = repo.Contains("qwen3-tts-onnx-dll", StringComparison.OrdinalIgnoreCase);
        var isQwenOnnxSplit = isQwen3Tts && repo.Contains("onnx", StringComparison.OrdinalIgnoreCase) && !isQwenOnnxDll;
        var isKittenTts = repo.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase);

        var localOptions = new LocalInferenceOptions
        {
            ModelCacheDir = _config.ModelCacheDir,
            ModelRepoId = repo,
            ModelBackend = _config.LocalModelBackend,
            PreferDevice = NormalizePreferDevice(_config.PreferDevice),
            MaxNewTokens = ResolveLocalMaxNewTokens(),
            ValidateOnnxRuntimeSessions = isChatterboxOnnx
        };

        if (string.Equals((_config.LocalModelPreset ?? string.Empty).Trim(), "chatterbox_onnx", StringComparison.OrdinalIgnoreCase) && usePythonBackend)
        {
            return new ChatterboxPythonBackend(localOptions);
        }

        if (isChatterboxOnnx)
        {
            return new ChatterboxOnnxBackend(localOptions);
        }

        if (isQwenOnnxDll)
        {
            return new QwenPythonBackend(localOptions);
        }

        if (isQwenOnnxSplit)
        {
            return new QwenPythonBackend(localOptions);
        }

        if (string.Equals((_config.LocalModelPreset ?? string.Empty).Trim(), "kitten_tts", StringComparison.OrdinalIgnoreCase) && usePythonBackend)
        {
            return new KittenPythonBackend(localOptions);
        }

        if (isKittenTts)
        {
            return new KittenTtsOnnxBackend(localOptions);
        }

        if (isQwen3Tts)
        {
            return new QwenPythonBackend(localOptions);
        }

        // For custom local repos, attempt native DLL backend first instead of hard-blocking.
        return new LocalDllBackend(localOptions);
    }

    private bool UsesPythonLocalBackend()
    {
        var preset = (_config.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        return preset is "chatterbox_onnx" or "kitten_tts" &&
               string.Equals(_config.LocalModelBackend, "python", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeLocalModelBackendChoice(string? value, string? preset, string? repoId)
    {
        var normalizedPreset = (preset ?? string.Empty).Trim().ToLowerInvariant();
        if (normalizedPreset == "chatterbox_onnx")
        {
            if (string.Equals((value ?? string.Empty).Trim(), "python", StringComparison.OrdinalIgnoreCase))
            {
                return "python";
            }

            return (repoId ?? string.Empty).Contains("onnx", StringComparison.OrdinalIgnoreCase) ? "onnx" : "python";
        }

        if (normalizedPreset == "kitten_tts")
        {
            return string.Equals((value ?? string.Empty).Trim(), "python", StringComparison.OrdinalIgnoreCase)
                ? "python"
                : "onnx";
        }

        return string.Empty;
    }

    private ITtsBackend CreateQwenBackendWithDevice(string preferDevice)
    {
        var repo = string.IsNullOrWhiteSpace(_config.ModelRepoId)
            ? "xkos/Qwen3-TTS-12Hz-1.7B-ONNX"
            : _config.ModelRepoId.Trim();
        var isQwen3Tts = repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase);

        if (!isQwen3Tts)
        {
            throw new InvalidOperationException("Qwen fallback backend requested for non-Qwen model.");
        }

        var localOptions = new LocalInferenceOptions
        {
            ModelCacheDir = _config.ModelCacheDir,
            ModelRepoId = repo,
            ModelBackend = _config.LocalModelBackend,
            PreferDevice = preferDevice,
            MaxNewTokens = ResolveLocalMaxNewTokens(),
            ValidateOnnxRuntimeSessions = false
        };

        return new QwenPythonBackend(localOptions);
    }

    private int ResolveLocalMaxNewTokens()
    {
        var s = GetEffectiveLocalSettings();
        var mode = (s.ChunkMode ?? "auto").Trim().ToLowerInvariant();
        var chunkChars = mode == "manual"
            ? Math.Clamp(Math.Max(s.MinChars, Math.Min(s.MaxChars, s.ManualMaxChars)), 80, 2400)
            : Math.Clamp(s.MaxChars, 120, 2400);

        var estimated = chunkChars * 2;
        // CustomVoice can produce long silent tails when max_new_tokens is too high.
        // Keep this capped lower for CustomVoice only; other model workflows remain unchanged.
        if (IsLocalQwenCustomVoiceMode())
        {
            return Math.Clamp(estimated, 220, 420);
        }

        return Math.Clamp(estimated, 320, 1024);
    }

    private async Task TryAutoDownloadModelAsync()
    {
        try
        {
            var progress = new Progress<string>(_ => { });
            await _modelDownloader.DownloadAsync(_config, progress, ct: CancellationToken.None);
        }
        catch
        {
            // Non-fatal on startup.
        }
    }

    private static string NormalizeLocalOutputExtension(string? format)
    {
        var normalized = (format ?? "wav").Trim().TrimStart('.').ToLowerInvariant();
        return normalized switch
        {
            "wav" => "wav",
            "mp3" => "mp3",
            _ => "wav"
        };
    }

    private static int ParseIntOrFallback(string? raw, int fallback)
    {
        return int.TryParse(raw?.Trim(), out var value) ? value : fallback;
    }

    private static double ParseDoubleOrFallback(string? raw, double fallback)
    {
        return double.TryParse(raw?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : fallback;
    }

    private static void AddPauseRangeRow(Grid panel, int row, string label, (int Min, int Max) range, out TextBox minBox, out TextBox maxBox)
    {
        var textLabel = new TextBlock { Text = $"{label} (ms)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(textLabel, row);
        panel.Children.Add(textLabel);

        var host = new Grid { Margin = new Thickness(0, 10, 0, 0) };
        host.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        host.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        host.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        minBox = new TextBox { Height = 32, Text = range.Min.ToString(CultureInfo.InvariantCulture) };
        maxBox = new TextBox { Height = 32, Text = range.Max.ToString(CultureInfo.InvariantCulture) };
        var sep = new TextBlock { Text = "to", Margin = new Thickness(8, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(minBox, 0);
        Grid.SetColumn(sep, 1);
        Grid.SetColumn(maxBox, 2);
        host.Children.Add(minBox);
        host.Children.Add(sep);
        host.Children.Add(maxBox);
        Grid.SetRow(host, row);
        Grid.SetColumn(host, 1);
        panel.Children.Add(host);
    }

    private static void AddQwenNumberField(Grid panel, int row, string label, string value, bool enabled, out TextBox textBox)
    {
        var textLabel = new TextBlock { Text = label, Margin = new Thickness(0, 8, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(textLabel, row);
        panel.Children.Add(textLabel);

        textBox = new TextBox { Height = 32, Margin = new Thickness(0, 8, 0, 0), Text = value, IsEnabled = enabled };
        Grid.SetRow(textBox, row);
        Grid.SetColumn(textBox, 1);
        panel.Children.Add(textBox);
    }

    private void RefreshModelControlSummary()
    {
        if (ModelControlsSummaryText is null)
        {
            return;
        }

        var s = GetEffectiveLocalSettings();
        var mode = string.IsNullOrWhiteSpace(s.ChunkMode) ? "auto" : s.ChunkMode.Trim().ToLowerInvariant();
        var size = mode == "manual" ? s.ManualMaxChars : s.MaxChars;
        var directive = s.ProtectBracketDirectives ? "bracket:on" : "bracket:off";
        var fmt = NormalizeLocalOutputExtension(s.OutputFormat);
        var device = NormalizePreferDevice(_config.PreferDevice);
        ModelControlsSummaryText.Text =
            $"Chunk {mode} ({s.MinChars}-{size}) | Pause {s.ChunkPauseMs}/{s.ParagraphPauseMs}ms | {fmt} | {directive} | device:{device}";
    }

    private static SynthesisSettings CloneSynthesisSettings(SynthesisSettings src)
    {
        return new SynthesisSettings
        {
            OutputFormat = src.OutputFormat,
            ChunkMode = src.ChunkMode,
            ManualMaxChars = src.ManualMaxChars,
            MinChars = src.MinChars,
            MaxChars = src.MaxChars,
            NarrationTargetSec = src.NarrationTargetSec,
            DialogueTargetSec = src.DialogueTargetSec,
            DialogueOverflow = src.DialogueOverflow,
            ChunkPauseMs = src.ChunkPauseMs,
            ParagraphPauseMs = src.ParagraphPauseMs,
            ClausePauseMinMs = src.ClausePauseMinMs,
            ClausePauseMaxMs = src.ClausePauseMaxMs,
            SentencePauseMinMs = src.SentencePauseMinMs,
            SentencePauseMaxMs = src.SentencePauseMaxMs,
            EllipsisPauseMinMs = src.EllipsisPauseMinMs,
            EllipsisPauseMaxMs = src.EllipsisPauseMaxMs,
            ParagraphPauseMinMs = src.ParagraphPauseMinMs,
            ParagraphPauseMaxMs = src.ParagraphPauseMaxMs,
            Atempo = src.Atempo,
            ReadSmallNumbersAsWords = src.ReadSmallNumbersAsWords,
            ProtectBracketDirectives = src.ProtectBracketDirectives,
            LocalInstructionHint = src.LocalInstructionHint,
            StylePresetKey = src.StylePresetKey,
            QwenStableAudiobookPreset = src.QwenStableAudiobookPreset,
            QwenUseRefText = src.QwenUseRefText,
            QwenDoSample = src.QwenDoSample,
            QwenTemperature = src.QwenTemperature,
            QwenTopK = src.QwenTopK,
            QwenTopP = src.QwenTopP,
            QwenRepetitionPenalty = src.QwenRepetitionPenalty,
            QwenAutoRetryBadChunks = src.QwenAutoRetryBadChunks,
            QwenBadChunkRetryCount = src.QwenBadChunkRetryCount
        };
    }

    private static bool IsQwenRepo(string? repoId)
        => (repoId ?? string.Empty).IndexOf("qwen3-tts", StringComparison.OrdinalIgnoreCase) >= 0;

    private static void ApplyQwenStableAudiobookPreset(SynthesisSettings settings)
    {
        if (!settings.QwenStableAudiobookPreset)
        {
            return;
        }

        // Prefer deterministic decoding first for long-form consistency.
        settings.QwenDoSample = false;
        settings.QwenTemperature = 0.7;
        settings.QwenTopK = 20;
        settings.QwenTopP = 0.9;
        settings.QwenRepetitionPenalty = Math.Clamp(settings.QwenRepetitionPenalty <= 0 ? 1.02 : settings.QwenRepetitionPenalty, 1.0, 1.2);
        settings.QwenAutoRetryBadChunks = true;
        settings.QwenBadChunkRetryCount = Math.Clamp(settings.QwenBadChunkRetryCount <= 0 ? 2 : settings.QwenBadChunkRetryCount, 0, 4);
    }

    private string BuildActiveModelProfileKey()
    {
        var repo = (_config.ModelRepoId ?? string.Empty).Trim().ToLowerInvariant();
        if (repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            return "qwen3_tts";
        }
        if (repo.Contains("chatterbox", StringComparison.OrdinalIgnoreCase) &&
            repo.Contains("onnx", StringComparison.OrdinalIgnoreCase))
        {
            return "chatterbox_onnx";
        }
        if (repo.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase))
        {
            return "kitten_tts";
        }

        if (!string.IsNullOrWhiteSpace(repo))
        {
            return repo;
        }

        var preset = (_config.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        return string.IsNullOrWhiteSpace(preset) ? "chatterbox_onnx" : preset;
    }

    private void ApplySelectedStyleRuntimeOverrides(SynthesisSettings settings)
    {
        if (settings is null)
        {
            return;
        }

        var style = (settings.StylePresetKey ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(style))
        {
            style = GetSelectedStylePresetKey();
            settings.StylePresetKey = style;
        }

        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        if (repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            // Style controls should actually affect Qwen runtime output, not just UI labels.
            switch (style)
            {
                case "expressive":
                    settings.QwenDoSample = true;
                    settings.QwenTemperature = 0.78;
                    settings.QwenTopK = 28;
                    settings.QwenTopP = 0.95;
                    settings.QwenRepetitionPenalty = Math.Clamp(settings.QwenRepetitionPenalty <= 1.0 ? 1.02 : settings.QwenRepetitionPenalty, 1.0, 1.2);
                    break;
                case "calm":
                    settings.QwenDoSample = false;
                    settings.QwenTemperature = 0.55;
                    settings.QwenTopK = 12;
                    settings.QwenTopP = 0.88;
                    settings.QwenRepetitionPenalty = Math.Clamp(Math.Max(settings.QwenRepetitionPenalty, 1.05), 1.0, 1.2);
                    break;
                case "dramatic":
                    settings.QwenDoSample = true;
                    settings.QwenTemperature = 0.86;
                    settings.QwenTopK = 36;
                    settings.QwenTopP = 0.97;
                    settings.QwenRepetitionPenalty = Math.Clamp(settings.QwenRepetitionPenalty <= 1.0 ? 1.01 : settings.QwenRepetitionPenalty, 1.0, 1.2);
                    break;
                case "energetic":
                    settings.QwenDoSample = true;
                    settings.QwenTemperature = 0.92;
                    settings.QwenTopK = 42;
                    settings.QwenTopP = 0.98;
                    settings.QwenRepetitionPenalty = Math.Clamp(settings.QwenRepetitionPenalty <= 1.0 ? 1.0 : settings.QwenRepetitionPenalty, 1.0, 1.2);
                    break;
                case "whispered":
                    settings.QwenDoSample = false;
                    settings.QwenTemperature = 0.45;
                    settings.QwenTopK = 8;
                    settings.QwenTopP = 0.82;
                    settings.QwenRepetitionPenalty = Math.Clamp(Math.Max(settings.QwenRepetitionPenalty, 1.08), 1.0, 1.2);
                    break;
                default:
                    // standard / narration-neutral: leave current profile values as-is.
                    break;
            }
        }
    }

    private (float SpeedScale, float? ChatterboxExaggeration) ResolveStyleRuntimeForCurrentModel(string styleKey)
    {
        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        if (!repo.Contains("chatterbox", StringComparison.OrdinalIgnoreCase))
        {
            return (1.0f, null);
        }

        return (styleKey ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "expressive" => (1.03f, 0.78f),
            "calm" => (0.95f, 0.32f),
            _ => (1.0f, 0.50f)
        };
    }

    private TtsRequest CreateLocalTtsRequest(
        string text,
        string voicePath,
        string outputPath,
        float baseSpeed,
        SynthesisSettings settings)
    {
        var styleKey = (settings.StylePresetKey ?? string.Empty).Trim().ToLowerInvariant();
        var isApiMode = string.Equals((_config.BackendMode ?? string.Empty).Trim(), "api", StringComparison.OrdinalIgnoreCase);
        var (speedScale, chatterboxExaggeration) = isApiMode
            ? (1.0f, (float?)null)
            : ResolveStyleRuntimeForCurrentModel(styleKey);
        var speed = Math.Clamp(baseSpeed * speedScale, 0.5f, 1.75f);
        var isQwenModel = IsQwenRepo(_config.ModelRepoId);
        var qwenUseRefText = !isQwenModel || ResolveEffectiveQwenUseRefText(settings);
        var refText = qwenUseRefText ? LoadRefTextForVoice(voicePath) : string.Empty;
        if (IsLocalQwenCustomVoiceMode())
        {
            var instructHint = TrimInstructionHintForQwen(settings.LocalInstructionHint);
            if (!string.IsNullOrWhiteSpace(instructHint))
            {
                refText = instructHint;
            }
        }

        return new TtsRequest
        {
            Text = text,
            VoicePath = voicePath,
            RefText = refText,
            OutputPath = outputPath,
            Speed = speed,
            StylePresetKey = styleKey,
            ChatterboxExaggeration = chatterboxExaggeration,
            QwenXVectorOnlyMode = isQwenModel ? !qwenUseRefText : null,
            QwenDoSample = settings.QwenDoSample,
            QwenTemperature = (float)Math.Clamp(settings.QwenTemperature, 0.05, 2.0),
            QwenTopK = Math.Clamp(settings.QwenTopK, 1, 200),
            QwenTopP = (float)Math.Clamp(settings.QwenTopP, 0.05, 1.0),
            QwenRepetitionPenalty = (float)Math.Clamp(settings.QwenRepetitionPenalty, 1.0, 2.0),
            QwenAutoRetryBadChunks = settings.QwenAutoRetryBadChunks,
            QwenBadChunkRetryCount = Math.Clamp(settings.QwenBadChunkRetryCount, 0, 4),
            ChunkPauseMs = Math.Clamp(settings.ChunkPauseMs, 0, 5000),
            ParagraphPauseMs = Math.Clamp(settings.ParagraphPauseMs, 0, 10000),
            ClausePauseMinMs = Math.Clamp(settings.ClausePauseMinMs, 0, 5000),
            ClausePauseMaxMs = Math.Clamp(settings.ClausePauseMaxMs, 0, 5000),
            SentencePauseMinMs = Math.Clamp(settings.SentencePauseMinMs, 0, 10000),
            SentencePauseMaxMs = Math.Clamp(settings.SentencePauseMaxMs, 0, 10000),
            EllipsisPauseMinMs = Math.Clamp(settings.EllipsisPauseMinMs, 0, 10000),
            EllipsisPauseMaxMs = Math.Clamp(settings.EllipsisPauseMaxMs, 0, 10000),
            ParagraphPauseMinMs = Math.Clamp(settings.ParagraphPauseMinMs, 0, 15000),
            ParagraphPauseMaxMs = Math.Clamp(settings.ParagraphPauseMaxMs, 0, 15000)
        };
    }

    private bool ResolveEffectiveQwenUseRefText(SynthesisSettings settings)
    {
        // Apply strict AND across all available sources so a single OFF always wins.
        // This prevents stale profile/project state from silently re-enabling ref_text.
        var runtimeSetting = settings.QwenUseRefText;
        var projectSetting = _project.Settings?.QwenUseRefText ?? true;
        var profileSetting = TryGetActiveProfileQwenUseRefText() ?? true;
        return runtimeSetting && projectSetting && profileSetting;
    }

    private bool? TryGetActiveProfileQwenUseRefText()
    {
        if (TryGetBestActiveModelProfile(out var profile, out _))
        {
            return profile.QwenUseRefText;
        }

        return null;
    }

    private bool TryGetBestActiveModelProfile(out SynthesisSettings profile, out string resolvedKey)
    {
        profile = new SynthesisSettings();
        resolvedKey = string.Empty;
        if (_config.ModelProfiles is null || _config.ModelProfiles.Count == 0)
        {
            return false;
        }

        foreach (var candidate in BuildActiveModelProfileKeys())
        {
            foreach (var kv in _config.ModelProfiles)
            {
                if (!string.Equals((kv.Key ?? string.Empty).Trim(), candidate, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (kv.Value is null)
                {
                    continue;
                }

                profile = kv.Value;
                resolvedKey = kv.Key ?? string.Empty;
                return true;
            }
        }

        return false;
    }

    private IEnumerable<string> BuildActiveModelProfileKeys()
    {
        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        var preset = (_config.LocalModelPreset ?? string.Empty).Trim();
        var keys = new List<string>();

        void Add(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (keys.Any(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            keys.Add(value);
        }

        Add(BuildActiveModelProfileKey());
        Add(repo.ToLowerInvariant());
        Add(preset.ToLowerInvariant());

        if (IsQwenRepo(repo))
        {
            Add("qwen3_tts");
            Add("qwen/qwen3-tts-12hz-1.7b-base");
            Add("xkos/qwen3-tts-12hz-1.7b-onnx");
            Add("zukky/qwen3-tts-onnx-dll");
        }
        else if (repo.IndexOf("chatterbox", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            Add("chatterbox_onnx");
            Add("onnx-community/chatterbox-onnx");
        }
        else if (repo.IndexOf("kitten-tts", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            Add("kitten_tts");
            Add("kittenml/kitten-tts-mini-0.8");
        }

        return keys;
    }

    private enum ModelFamily
    {
        Other,
        Qwen,
        Chatterbox,
        Kitten
    }

    private static ModelFamily DetectModelFamily(string? repoOrKey)
    {
        var value = (repoOrKey ?? string.Empty).Trim();
        if (value.IndexOf("qwen3-tts", StringComparison.OrdinalIgnoreCase) >= 0 || value.Equals("qwen3_tts", StringComparison.OrdinalIgnoreCase))
        {
            return ModelFamily.Qwen;
        }
        if (value.IndexOf("chatterbox", StringComparison.OrdinalIgnoreCase) >= 0 || value.Equals("chatterbox_onnx", StringComparison.OrdinalIgnoreCase))
        {
            return ModelFamily.Chatterbox;
        }
        if (value.IndexOf("kitten-tts", StringComparison.OrdinalIgnoreCase) >= 0 || value.Equals("kitten_tts", StringComparison.OrdinalIgnoreCase))
        {
            return ModelFamily.Kitten;
        }
        return ModelFamily.Other;
    }

    private static bool IsFamilyKey(ModelFamily family, string key)
    {
        return DetectModelFamily(key) == family;
    }

    private void OverwriteCurrentModelSettingsEverywhere(SynthesisSettings settings)
    {
        var snapshot = CloneSynthesisSettings(settings);
        _project.Settings = CloneSynthesisSettings(snapshot);
        _config.ModelProfiles ??= new Dictionary<string, SynthesisSettings>();

        var family = DetectModelFamily(_config.ModelRepoId);
        var keysToWrite = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in BuildActiveModelProfileKeys())
        {
            keysToWrite.Add(key);
        }

        foreach (var existingKey in _config.ModelProfiles.Keys.ToList())
        {
            if (IsFamilyKey(family, existingKey))
            {
                keysToWrite.Add(existingKey);
            }
        }

        foreach (var key in keysToWrite)
        {
            _config.ModelProfiles[key] = CloneSynthesisSettings(snapshot);
        }
    }

    private SynthesisSettings GetEffectiveLocalSettings()
    {
        var baseSettings = CloneSynthesisSettings(_project.Settings ?? new SynthesisSettings());
        if (_config.ModelProfiles is null || _config.ModelProfiles.Count == 0)
        {
            if (IsQwenRepo(_config.ModelRepoId))
            {
                ApplyQwenStableAudiobookPreset(baseSettings);
            }
            return baseSettings;
        }

        if (TryGetBestActiveModelProfile(out var perModel, out _))
        {
            var selected = CloneSynthesisSettings(perModel);
            // Leak-safety: if either project-level or profile-level setting disables ref_text, keep it disabled.
            // This avoids stale profile values silently forcing ICL mode back on.
            var projectWantsRefText = _project.Settings?.QwenUseRefText ?? true;
            selected.QwenUseRefText = selected.QwenUseRefText && projectWantsRefText;
            if (IsQwenRepo(_config.ModelRepoId))
            {
                ApplyQwenStableAudiobookPreset(selected);
            }
            return selected;
        }

        if (IsQwenRepo(_config.ModelRepoId))
        {
            ApplyQwenStableAudiobookPreset(baseSettings);
        }
        return baseSettings;
    }

    private async Task<IReadOnlyList<SubtitleChunkTiming>?> GenerateLocalWithChunkingAsync(
        ITtsBackend backend,
        string text,
        string voicePath,
        string outputPath,
        float speed,
        SynthesisSettings settings,
        CancellationToken ct,
        Action<int, int>? onChunkProgress = null)
    {
        var chunkProfile = ResolveChunkProfile();
        var preparedText = chunkProfile == ChunkProfile.Qwen3Tts
            ? NormalizeTextForQwenTts(text)
            : chunkProfile == ChunkProfile.KittenTts
                ? NormalizeTextForKittenTts(text)
            : text;
        var chunks = BuildTextChunks(preparedText, settings, chunkProfile);
        if (chunkProfile == ChunkProfile.Qwen3Tts)
        {
            TraceQwenChunkTrace("raw", chunks);
        }
        if (chunkProfile == ChunkProfile.Qwen3Tts && chunks.Count > 1)
        {
            chunks = RepackQwenChunksByTokenBudget(chunks, settings);
        }
        if (chunkProfile == ChunkProfile.Qwen3Tts)
        {
            TraceQwenChunkTrace("final", chunks);
        }
        if (chunks.Count == 0)
        {
            throw new InvalidOperationException("Input text is empty.");
        }

        var chunkPreview = string.Join(", ", chunks.Take(6).Select(c => c.Text.Length));
        if (chunks.Count > 6)
        {
            chunkPreview += ", ...";
        }
        TraceGenerate($"Chunking ({chunkProfile}) prepared {chunks.Count} chunk(s). Lengths: {chunkPreview}");

        // --- DIAG: Step 1 REPRODUCE ---
        var totalChars = chunks.Sum(c => c.Text.Length);
        TraceGenerate($"[DIAG REPRODUCE] Backend={backend.GetType().Name}, Profile={chunkProfile}, InputChars={preparedText.Length}, Chunks={chunks.Count}, CharsPerChunkAvg={totalChars / Math.Max(1, chunks.Count)}");
        // --- DIAG: Step 2 DESCRIBE FAILURE ---
        TraceGenerate("[DIAG EXPECTED] Each chunk = decent audio quality (clear speech, consistent voice). Full chapter = all chunks good, no cut-off.");
        TraceGenerate("[DIAG ACTUAL] Quality failure: even one chunk with bad/silent/choppy audio is the bug. Duration is a symptom.");
        // --- DIAG: Step 3 HYPOTHESES ---
        TraceGenerate("[DIAG HYPOTHESES] A=backend produces bad quality per chunk B=wrong input/decode to model C=sample rate/format mismatch D=merge/trim corrupts audio E=other");

        onChunkProgress?.Invoke(0, chunks.Count);

        var tempRoot = Path.Combine(Path.GetTempPath(), $"AudiobookCreator_chunks_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        try
        {
            var chunkWavs = new List<string>(chunks.Count);
            var chunkAudioSeconds = new List<double>(chunks.Count);
            string? previousChunkTail = null;
            bool? previousChunkWasDialogue = null;
            for (var i = 0; i < chunks.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                while (_pauseGenerationRequested && !ct.IsCancellationRequested)
                {
                    await Task.Delay(120, ct);
                }
                var chunkPath = Path.Combine(tempRoot, $"chunk_{i + 1:D4}.wav");
                var continuityTail = chunkProfile == ChunkProfile.Qwen3Tts &&
                                     previousChunkTail is not null &&
                                     previousChunkWasDialogue == chunks[i].IsDialogue
                    ? previousChunkTail
                    : null;
                var chunkText = ApplyLocalInstruction(chunks[i].Text, settings, continuityTail, chunks[i].IsDialogue);
                var request = CreateLocalTtsRequest(chunkText, voicePath, chunkPath, speed, settings);
                if (chunkProfile == ChunkProfile.Qwen3Tts)
                {
                    var qRefRuntime = settings.QwenUseRefText;
                    var qRefProject = _project.Settings?.QwenUseRefText ?? true;
                    var qRefProfile = TryGetActiveProfileQwenUseRefText();
                    TraceGenerate(
                        $"[QWEN CALL] Chunk {i + 1}/{chunks.Count} " +
                        $"chars={chunks[i].Text.Length}, words={CountWords(chunks[i].Text)}, estSec={EstimateSeconds(chunks[i].Text):0.0}, " +
                        $"dialogue={(chunks[i].IsDialogue ? 1 : 0)}, paraEnd={(chunks[i].ParagraphEnd ? 1 : 0)}, " +
                        $"hash={StableTextHashHex(chunks[i].Text)}, preview={QuoteForLog(ChunkPreview(chunks[i].Text))}, " +
                        $"q_useRefText={(!string.IsNullOrWhiteSpace(request.RefText) ? "1" : "0")}, " +
                        $"q_xvec={(request.QwenXVectorOnlyMode.HasValue ? (request.QwenXVectorOnlyMode.Value ? "1" : "0") : "null")}, " +
                        $"q_refChars={(request.RefText ?? string.Empty).Length}, " +
                        $"q_refCfgRuntime={(qRefRuntime ? 1 : 0)}, q_refCfgProject={(qRefProject ? 1 : 0)}, q_refCfgProfile={(qRefProfile.HasValue ? (qRefProfile.Value ? 1 : 0) : -1)}, " +
                        $"q_doSample={(request.QwenDoSample.HasValue ? request.QwenDoSample.Value.ToString() : "null")}, " +
                        $"q_temp={(request.QwenTemperature.HasValue ? request.QwenTemperature.Value.ToString("0.###", CultureInfo.InvariantCulture) : "null")}, " +
                        $"q_topK={(request.QwenTopK.HasValue ? request.QwenTopK.Value.ToString(CultureInfo.InvariantCulture) : "null")}, " +
                        $"q_topP={(request.QwenTopP.HasValue ? request.QwenTopP.Value.ToString("0.###", CultureInfo.InvariantCulture) : "null")}, " +
                        $"q_rep={(request.QwenRepetitionPenalty.HasValue ? request.QwenRepetitionPenalty.Value.ToString("0.###", CultureInfo.InvariantCulture) : "null")}, " +
                        $"q_maxNewEst={ResolveLocalMaxNewTokens()}");
                }

                try
                {
                    await backend.SynthesizeAsync(request, ct);
                }
                catch (Exception ex)
                {
                    TraceGenerate($"[DIAG EVIDENCE] Chunk {i + 1}/{chunks.Count} SYNTHESIZE EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                    throw;
                }

                if (chunkProfile == ChunkProfile.Qwen3Tts &&
                    TryDetectSuspiciousQwenChunkAudio(chunkPath, out var qwenGuardReason))
                {
                    TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: output quality guard ({qwenGuardReason}); retrying chunk.");

                    var recovered = false;
                    string? lastBadReason = qwenGuardReason;

                    // Retry once on the same backend first (transient generation glitches do happen).
                    try
                    {
                        await backend.SynthesizeAsync(request, ct);
                        if (!TryDetectSuspiciousQwenChunkAudio(chunkPath, out var retryReason))
                        {
                            recovered = true;
                        }
                        else
                        {
                            lastBadReason = retryReason;
                            TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: same-backend retry still suspicious ({retryReason}).");
                        }
                    }
                    catch (Exception ex)
                    {
                        TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: same-backend retry failed ({ex.GetType().Name}: {ex.Message}).");
                    }

                    var forceGpuOnlyRecovery = string.Equals(NormalizePreferDevice(_config.PreferDevice), "gpu", StringComparison.OrdinalIgnoreCase);
                    ITtsBackend? fallbackBackend = null;
                    if (!recovered && !forceGpuOnlyRecovery)
                    {
                        // For Qwen Python backend, CPU fallback is the reliable alternate path.
                        try
                        {
                            fallbackBackend = CreateQwenBackendWithDevice("cpu");
                            await fallbackBackend.SynthesizeAsync(request, ct);
                            if (!TryDetectSuspiciousQwenChunkAudio(chunkPath, out var fallbackReason))
                            {
                                recovered = true;
                                TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: CPU fallback succeeded.");
                            }
                            else
                            {
                                lastBadReason = fallbackReason;
                                TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: fallback also suspicious ({fallbackReason}); attempting sub-split.");
                            }
                        }
                        catch (Exception ex)
                        {
                            TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: CPU fallback failed ({ex.GetType().Name}: {ex.Message}); attempting sub-split.");
                        }
                    }
                    else if (!recovered && forceGpuOnlyRecovery)
                    {
                        TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: GPU is selected, skipping CPU fallback and attempting recursive sub-split on GPU.");
                    }

                    if (!recovered)
                    {
                        var subSplitBackend = fallbackBackend ?? backend;
                        recovered = await TryGenerateQwenSubSplitAsync(
                            chunks[i].Text,
                            chunkPath,
                            voicePath,
                            speed,
                            settings,
                            subSplitBackend,
                            tempRoot,
                            i + 1,
                            ct);
                        if (recovered)
                        {
                            TraceGenerate($"Qwen chunk {i + 1}/{chunks.Count}: sub-split succeeded.");
                        }
                    }

                    if (fallbackBackend is IDisposable disposableFallback)
                    {
                        disposableFallback.Dispose();
                    }

                    if (!recovered)
                    {
                        throw new InvalidOperationException(
                            $"Qwen quality guard: chunk output is mostly blank after retry/split recovery (chunk {i + 1}/{chunks.Count}, {lastBadReason}).");
                    }
                }

                chunkWavs.Add(chunkPath);
                // --- DIAG: Step 4 GATHER EVIDENCE (per chunk: quality + duration) ---
                if (TryInspectWav(chunkPath, out var chunkWavInfo, out _))
                {
                    var lowLevel = chunkWavInfo.Rms < 0.01;
                    chunkAudioSeconds.Add(Math.Max(0.04, chunkWavInfo.DurationSeconds));
                    TraceGenerate($"[DIAG EVIDENCE] Chunk {i + 1}/{chunks.Count} durationSec={chunkWavInfo.DurationSeconds:F2} rms={chunkWavInfo.Rms:F4} sampleRate={chunkWavInfo.SampleRate} bytes={new FileInfo(chunkPath).Length}{(lowLevel ? " LOW_LEVEL_BAD" : "")}");
                }
                else
                {
                    chunkAudioSeconds.Add(Math.Max(0.04, EstimateSeconds(chunks[i].Text)));
                    TraceGenerate($"[DIAG EVIDENCE] Chunk {i + 1}/{chunks.Count} WAV inspect failed or file missing.");
                }
                if (chunkProfile == ChunkProfile.Qwen3Tts)
                {
                    previousChunkTail = BuildChunkContinuityTail(chunks[i].Text);
                    previousChunkWasDialogue = chunks[i].IsDialogue;
                }
                onChunkProgress?.Invoke(i + 1, chunks.Count);
            }

            var mergedWavPath = Path.Combine(tempRoot, "merged.wav");
            MergeChunkWavs(chunkWavs, chunks, mergedWavPath, settings, chunkProfile,
                trimInteriorSilences: chunkProfile == ChunkProfile.Qwen3Tts);
            ct.ThrowIfCancellationRequested();
            // --- DIAG: Step 4 (merged) ---
            if (TryInspectWav(mergedWavPath, out var mergedWavInfo, out _))
                TraceGenerate($"[DIAG EVIDENCE] MergedWav durationSec={mergedWavInfo.DurationSeconds:F2} chunkCount={chunkWavs.Count}");
            else
                TraceGenerate("[DIAG EVIDENCE] MergedWav inspect failed.");
            await FinalizeLocalOutputAsync(mergedWavPath, outputPath);
            // --- DIAG: Step 4 (final output) ---
            if (File.Exists(outputPath))
            {
                var ext = Path.GetExtension(outputPath);
                if (string.Equals(ext, ".wav", StringComparison.OrdinalIgnoreCase) && TryInspectWav(outputPath, out var outWavInfo, out _))
                    TraceGenerate($"[DIAG EVIDENCE] Final output durationSec={outWavInfo.DurationSeconds:F2} path={outputPath}");
                else
                    TraceGenerate($"[DIAG EVIDENCE] Final output path={outputPath} sizeBytes={new FileInfo(outputPath).Length}");
            }
            else
                TraceGenerate("[DIAG EVIDENCE] Final output file not found.");

            var finalSeconds = TryGetAudioDurationSeconds(outputPath, out var measuredSeconds)
                ? measuredSeconds
                : Math.Max(0.2, chunkAudioSeconds.Sum());
            var subtitleTimings = BuildSubtitleChunkTimings(chunks, chunkAudioSeconds, settings, chunkProfile, finalSeconds);
            return subtitleTimings;
        }
        finally
        {
            try
            {
                Directory.Delete(tempRoot, recursive: true);
            }
            catch
            {
                // Best effort cleanup only.
            }
        }
    }

    private async Task<IReadOnlyList<SubtitleChunkTiming>?> GeneratePreparedScriptLocalAsync(
        ITtsBackend backend,
        PreparedScriptDocument prepared,
        string outputPath,
        float speed,
        SynthesisSettings settings,
        CancellationToken ct,
        Action<int, int>? onPartProgress = null)
    {
        var validParts = (prepared.Parts ?? new List<PreparedScriptPart>())
            .Where(p => !string.IsNullOrWhiteSpace((p.Text ?? string.Empty).Trim()))
            .OrderBy(p => p.Order)
            .ToList();
        if (validParts.Count == 0)
        {
            throw new InvalidOperationException("Prepared script has no usable parts.");
        }

        var tempRoot = Path.Combine(Path.GetTempPath(), $"AudiobookCreator_prepared_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        try
        {
            var partOutputWavs = new List<string>(validParts.Count);
            var mergeChunks = new List<TextChunk>(validParts.Count);
            var subtitleList = new List<SubtitleChunkTiming>(validParts.Count);
            var currentStart = 0.0;
            var chunkProfile = ResolveChunkProfile();

            onPartProgress?.Invoke(0, validParts.Count);
            for (var i = 0; i < validParts.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var part = validParts[i];
                var localSettings = CloneSynthesisSettings(settings);
                localSettings.LocalInstructionHint = IsLocalQwenCustomVoiceMode()
                    ? TrimInstructionHintForQwen(part.Instruction)
                    : string.Empty;
                var partVoice = ResolvePreparedPartVoicePath(part.VoicePath);
                var partWav = Path.Combine(tempRoot, $"part_{i + 1:D4}.wav");
                await GenerateLocalWithChunkingAsync(
                    backend,
                    part.Text,
                    partVoice,
                    partWav,
                    speed,
                    localSettings,
                    ct);

                if (!TryInspectWav(partWav, out var wavInfo, out _))
                {
                    throw new InvalidOperationException($"Prepared part {i + 1} produced invalid audio.");
                }

                partOutputWavs.Add(partWav);
                var isNarrator = string.Equals(part.SpeakerTag, "Narrator", StringComparison.OrdinalIgnoreCase);
                var chunk = new TextChunk(part.Text, i == validParts.Count - 1, !isNarrator);
                mergeChunks.Add(chunk);
                var end = currentStart + wavInfo.DurationSeconds;
                subtitleList.Add(new SubtitleChunkTiming(part.Text, currentStart, end));
                if (i < validParts.Count - 1)
                {
                    var pauseMs = ResolveChunkBoundaryPauseMs(chunk, i, settings, chunkProfile);
                    currentStart = end + (pauseMs / 1000.0);
                }
                else
                {
                    currentStart = end;
                }

                onPartProgress?.Invoke(i + 1, validParts.Count);
            }

            var mergedWav = Path.Combine(tempRoot, "prepared_merged.wav");
            MergeChunkWavs(partOutputWavs, mergeChunks, mergedWav, settings, chunkProfile, trimInteriorSilences: true);
            await FinalizeLocalOutputAsync(mergedWav, outputPath);
            return subtitleList;
        }
        finally
        {
            try
            {
                Directory.Delete(tempRoot, recursive: true);
            }
            catch
            {
                // Best effort cleanup only.
            }
        }
    }

    private string ResolvePreparedPartVoicePath(string? preferredVoicePath)
    {
        var mixedMode = IsMixedVoiceToken(_project.VoicePath);
        if (IsMixedVoiceToken(preferredVoicePath))
        {
            preferredVoicePath = string.Empty;
        }

        if (mixedMode)
        {
            if (IsLocalQwenCustomVoiceMode())
            {
                if (!string.IsNullOrWhiteSpace(preferredVoicePath) &&
                    TryParseQwenCustomVoiceToken(preferredVoicePath, out var speaker))
                {
                    return BuildQwenCustomVoiceToken(speaker);
                }

                throw new InvalidOperationException("Mixed mode requires each prepared part to use a valid Qwen CustomVoice speaker.");
            }

            if (IsActiveKittenModel())
            {
                if (!string.IsNullOrWhiteSpace(preferredVoicePath) &&
                    TryParseKittenVoiceToken(preferredVoicePath, out var kittenVoice))
                {
                    return BuildKittenVoiceToken(kittenVoice);
                }

                throw new InvalidOperationException("Mixed mode requires each prepared part to use a valid Kitten voice.");
            }

            if (!string.IsNullOrWhiteSpace(preferredVoicePath) && File.Exists(preferredVoicePath))
            {
                return preferredVoicePath.Trim();
            }

            throw new InvalidOperationException("Mixed mode requires each prepared part to have a valid voice file.");
        }

        if (IsLocalQwenCustomVoiceMode())
        {
            if (!string.IsNullOrWhiteSpace(preferredVoicePath) &&
                TryParseQwenCustomVoiceToken(preferredVoicePath, out var preferredSpeaker))
            {
                return BuildQwenCustomVoiceToken(preferredSpeaker);
            }

            if (!string.IsNullOrWhiteSpace(_project.VoicePath) &&
                TryParseQwenCustomVoiceToken(_project.VoicePath, out var projectSpeaker))
            {
                return BuildQwenCustomVoiceToken(projectSpeaker);
            }
        }

        if (!string.IsNullOrWhiteSpace(preferredVoicePath) && File.Exists(preferredVoicePath))
        {
            return preferredVoicePath.Trim();
        }

        if (!string.IsNullOrWhiteSpace(_project.VoicePath) && File.Exists(_project.VoicePath))
        {
            return _project.VoicePath.Trim();
        }

        throw new InvalidOperationException("Prepared script part needs a valid voice (part voice or project default voice).");
    }

    private static string TrimInstructionHintForQwen(string? value)
    {
        var clean = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        clean = clean.Replace("[", string.Empty, StringComparison.Ordinal)
                     .Replace("]", string.Empty, StringComparison.Ordinal);
        return clean.Length <= 220 ? clean : clean[..220].Trim();
    }

    private ChunkProfile ResolveChunkProfile()
    {
        var repo = (_config.ModelRepoId ?? string.Empty).Trim();
        if (repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            return ChunkProfile.Qwen3Tts;
        }
        if (repo.Contains("kitten-tts", StringComparison.OrdinalIgnoreCase))
        {
            return ChunkProfile.KittenTts;
        }

        return ChunkProfile.Chatterbox;
    }

    private async Task GenerateSubtitlesIfEnabledAsync(
        string sourceText,
        string outputPath,
        IReadOnlyList<SubtitleChunkTiming>? chunkTimings)
    {
        if (!_config.GenerateSrtSubtitles && !_config.GenerateAssSubtitles)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(outputPath) || !File.Exists(outputPath))
        {
            return;
        }

        if (!TryGetAudioDurationSeconds(outputPath, out var totalSeconds) || totalSeconds < 0.2)
        {
            return;
        }

        var cues = BuildSubtitleCues(sourceText, chunkTimings, totalSeconds);
        if (cues.Count == 0)
        {
            return;
        }

        var basePath = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? RuntimePaths.AppRoot,
            Path.GetFileNameWithoutExtension(outputPath));

        if (_config.GenerateSrtSubtitles)
        {
            var srtPath = basePath + ".srt";
            await File.WriteAllTextAsync(srtPath, BuildSrtContent(cues), Encoding.UTF8);
        }

        if (_config.GenerateAssSubtitles)
        {
            var assPath = basePath + ".ass";
            await File.WriteAllTextAsync(assPath, BuildAssContent(cues), Encoding.UTF8);
        }
    }

    private static IReadOnlyList<SubtitleChunkTiming> BuildSubtitleChunkTimings(
        IReadOnlyList<TextChunk> chunks,
        IReadOnlyList<double> chunkAudioSeconds,
        SynthesisSettings settings,
        ChunkProfile chunkProfile,
        double finalDurationSeconds)
    {
        if (chunks.Count == 0)
        {
            return Array.Empty<SubtitleChunkTiming>();
        }

        var unitDurations = new List<double>(chunks.Count);
        var estimatedTotal = 0.0;
        for (var i = 0; i < chunks.Count; i++)
        {
            var audioSec = i < chunkAudioSeconds.Count ? Math.Max(0.04, chunkAudioSeconds[i]) : Math.Max(0.04, EstimateSeconds(chunks[i].Text));
            var pauseSec = i < chunks.Count - 1
                ? ResolveChunkBoundaryPauseMs(chunks[i], i, settings, chunkProfile) / 1000.0
                : 0.0;
            var span = Math.Max(0.05, audioSec + pauseSec);
            unitDurations.Add(span);
            estimatedTotal += span;
        }

        var scale = estimatedTotal > 0.001 ? Math.Max(0.2, finalDurationSeconds / estimatedTotal) : 1.0;
        var timings = new List<SubtitleChunkTiming>(chunks.Count);
        var cursor = 0.0;
        for (var i = 0; i < chunks.Count; i++)
        {
            var span = unitDurations[i] * scale;
            var start = cursor;
            var end = i == chunks.Count - 1 ? finalDurationSeconds : Math.Min(finalDurationSeconds, cursor + span);
            if (end - start < 0.05)
            {
                end = Math.Min(finalDurationSeconds, start + 0.05);
            }
            timings.Add(new SubtitleChunkTiming(chunks[i].Text, start, end));
            cursor = end;
        }

        return timings;
    }

    private static List<SubtitleCue> BuildSubtitleCues(
        string sourceText,
        IReadOnlyList<SubtitleChunkTiming>? chunkTimings,
        double totalDurationSeconds)
    {
        var cues = new List<SubtitleCue>();
        if (chunkTimings is not null && chunkTimings.Count > 0)
        {
            foreach (var chunk in chunkTimings)
            {
                var lines = SplitSubtitleLines(chunk.Text);
                if (lines.Count == 0)
                {
                    continue;
                }

                var span = Math.Max(0.15, chunk.EndSeconds - chunk.StartSeconds);
                var totalChars = Math.Max(1, lines.Sum(l => l.Length));
                var cursor = chunk.StartSeconds;
                for (var i = 0; i < lines.Count; i++)
                {
                    var weight = lines[i].Length / (double)totalChars;
                    var lineSpan = i == lines.Count - 1
                        ? Math.Max(0.05, chunk.EndSeconds - cursor)
                        : Math.Max(0.05, span * weight);
                    var start = cursor;
                    var end = i == lines.Count - 1 ? chunk.EndSeconds : Math.Min(chunk.EndSeconds, cursor + lineSpan);
                    if (end <= start)
                    {
                        end = Math.Min(chunk.EndSeconds, start + 0.05);
                    }
                    cues.Add(new SubtitleCue(start, end, lines[i]));
                    cursor = end;
                }
            }

            if (cues.Count > 0)
            {
                cues[^1] = cues[^1] with { EndSeconds = Math.Max(cues[^1].StartSeconds + 0.05, totalDurationSeconds) };
            }
            return cues;
        }

        var fallbackLines = SplitSubtitleLines(sourceText);
        if (fallbackLines.Count == 0)
        {
            return cues;
        }

        var sumChars = Math.Max(1, fallbackLines.Sum(l => l.Length));
        var t = 0.0;
        for (var i = 0; i < fallbackLines.Count; i++)
        {
            var weight = fallbackLines[i].Length / (double)sumChars;
            var span = i == fallbackLines.Count - 1
                ? Math.Max(0.05, totalDurationSeconds - t)
                : Math.Max(0.05, totalDurationSeconds * weight);
            var start = t;
            var end = i == fallbackLines.Count - 1 ? totalDurationSeconds : Math.Min(totalDurationSeconds, t + span);
            cues.Add(new SubtitleCue(start, end, fallbackLines[i]));
            t = end;
        }

        return cues;
    }

    private static List<string> SplitSubtitleLines(string text, int maxChars = 72)
    {
        var normalized = Regex.Replace((text ?? string.Empty).Replace("\r\n", " "), @"\s+", " ").Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return new List<string>();
        }

        var sentenceParts = Regex.Split(normalized, @"(?<=[\.\!\?])\s+")
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
        if (sentenceParts.Count == 0)
        {
            sentenceParts.Add(normalized);
        }

        var lines = new List<string>();
        foreach (var part in sentenceParts)
        {
            if (part.Length <= maxChars)
            {
                lines.Add(part);
                continue;
            }

            var words = part.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            foreach (var w in words)
            {
                if (sb.Length > 0 && sb.Length + 1 + w.Length > maxChars)
                {
                    lines.Add(sb.ToString());
                    sb.Clear();
                }

                if (sb.Length > 0)
                {
                    sb.Append(' ');
                }
                sb.Append(w);
            }
            if (sb.Length > 0)
            {
                lines.Add(sb.ToString());
            }
        }

        return lines;
    }

    private static string BuildSrtContent(IReadOnlyList<SubtitleCue> cues)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < cues.Count; i++)
        {
            var cue = cues[i];
            sb.AppendLine((i + 1).ToString(CultureInfo.InvariantCulture));
            sb.AppendLine($"{FormatSrtTime(cue.StartSeconds)} --> {FormatSrtTime(cue.EndSeconds)}");
            sb.AppendLine(cue.Text);
            sb.AppendLine();
        }
        return sb.ToString();
    }

    private static string BuildAssContent(IReadOnlyList<SubtitleCue> cues)
    {
        var sb = new StringBuilder();
        sb.AppendLine("[Script Info]");
        sb.AppendLine("ScriptType: v4.00+");
        sb.AppendLine("WrapStyle: 0");
        sb.AppendLine("ScaledBorderAndShadow: yes");
        sb.AppendLine();
        sb.AppendLine("[V4+ Styles]");
        sb.AppendLine("Format: Name,Fontname,Fontsize,PrimaryColour,SecondaryColour,OutlineColour,BackColour,Bold,Italic,Underline,StrikeOut,ScaleX,ScaleY,Spacing,Angle,BorderStyle,Outline,Shadow,Alignment,MarginL,MarginR,MarginV,Encoding");
        sb.AppendLine("Style: Default,Segoe UI,42,&H00FFFFFF,&H000000FF,&H00111111,&H50000000,0,0,0,0,100,100,0,0,1,2,0,2,50,50,40,1");
        sb.AppendLine();
        sb.AppendLine("[Events]");
        sb.AppendLine("Format: Layer,Start,End,Style,Name,MarginL,MarginR,MarginV,Effect,Text");

        foreach (var cue in cues)
        {
            var text = cue.Text.Replace("\\", "\\\\").Replace("{", "\\{").Replace("}", "\\}");
            sb.AppendLine($"Dialogue: 0,{FormatAssTime(cue.StartSeconds)},{FormatAssTime(cue.EndSeconds)},Default,,0,0,0,,{text}");
        }

        return sb.ToString();
    }

    private static string FormatSrtTime(double seconds)
    {
        var ms = (long)Math.Round(Math.Max(0, seconds) * 1000.0);
        var ts = TimeSpan.FromMilliseconds(ms);
        return $"{(int)ts.TotalHours:00}:{ts.Minutes:00}:{ts.Seconds:00},{ts.Milliseconds:000}";
    }

    private static string FormatAssTime(double seconds)
    {
        var cs = (long)Math.Round(Math.Max(0, seconds) * 100.0);
        var hours = cs / 360000;
        var rem = cs % 360000;
        var mins = rem / 6000;
        rem %= 6000;
        var secs = rem / 100;
        var centis = rem % 100;
        return $"{hours}:{mins:00}:{secs:00}.{centis:00}";
    }

    private static bool TryGetAudioDurationSeconds(string path, out double durationSeconds)
    {
        durationSeconds = 0;
        try
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return false;
            }

            using var reader = new AudioFileReader(path);
            durationSeconds = Math.Max(0, reader.TotalTime.TotalSeconds);
            return durationSeconds > 0;
        }
        catch
        {
            return false;
        }
    }

    private static readonly HashSet<string> SentenceAbbreviations = new(StringComparer.OrdinalIgnoreCase)
    {
        "mr.", "mrs.", "ms.", "dr.", "st.", "sr.", "jr.", "vs.", "prof.", "inc.", "ltd.", "etc."
    };

    private static readonly Regex DialogueParagraphPattern = new(@"^\s*[""']", RegexOptions.Compiled);
    private static readonly Regex DialogueTagPattern = new(
        @"^\s*("".*?"")(?:\s*[,;]\s*(?:he|she|they|i|you|we|[A-Z][a-z]+)\s+(?:said|asked|replied|whispered|yelled|murmured|shouted|answered)\b[^.]*)?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
    private static readonly Regex InlineDialoguePattern = new(@"""[^""]+""|'[^']+'", RegexOptions.Compiled | RegexOptions.Singleline);

    private static List<TextChunk> BuildTextChunks(string text, SynthesisSettings settings, ChunkProfile profile)
    {
        var source = (text ?? string.Empty).Replace("\r\n", "\n");
        var paragraphs = Regex.Split(source, @"\n{2,}")
            .Select(p => Regex.Replace(p, @"\s+", " ").Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToList();

        if (paragraphs.Count == 0)
        {
            return new List<TextChunk>();
        }

        var mode = (settings.ChunkMode ?? "auto").Trim().ToLowerInvariant();
        return mode == "manual"
            ? BuildTextChunksByChars(paragraphs, settings)
            : BuildTextChunksByTime(paragraphs, settings, profile);
    }

    private List<TextChunk> RepackQwenChunksByTokenBudget(IReadOnlyList<TextChunk> chunks, SynthesisSettings settings)
    {
        var repo = string.IsNullOrWhiteSpace(_config.ModelRepoId) ? "xkos/Qwen3-TTS-12Hz-1.7B-ONNX" : _config.ModelRepoId.Trim();
        if (!repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            return chunks.ToList();
        }

        var tokenizerRepo = string.IsNullOrWhiteSpace(_config.AdditionalModelRepoId)
            ? "zukky/Qwen3-TTS-ONNX-DLL"
            : _config.AdditionalModelRepoId.Trim();

        if (!QwenTokenCounter.TryCreate(_config.ModelCacheDir, repo, tokenizerRepo, ResolveLocalMaxNewTokens(), out var counter, out var error) ||
            counter is null)
        {
            TraceGenerate($"Qwen token-aware repack unavailable: {error}");
            return chunks.ToList();
        }

        using (counter)
        {
            var budget = Math.Max(700, counter.RecommendedPromptTokenBudget);
            var preferDevice = NormalizePreferDevice(_config.PreferDevice);
            var isGpuPreferred = string.Equals(preferDevice, "gpu", StringComparison.OrdinalIgnoreCase);
            // Token budget alone is too loose for stable TTS audio duration. Keep a practical cap too.
            // GPU (especially CUDA on some ORT/Qwen graphs) is more sensitive to long dialogue chunks.
            var secondsCap = isGpuPreferred
                ? Math.Clamp(Math.Max(settings.NarrationTargetSec * 1.9, 11.5), 9.0, 16.0)
                : Math.Clamp(Math.Max(settings.NarrationTargetSec * 2.2, 14.0), 10.0, 22.0);
            var wordCap = isGpuPreferred
                ? (int)Math.Clamp(Math.Round(secondsCap * 3.0), 44, 62)
                : (int)Math.Clamp(Math.Round(secondsCap * 3.0), 45, 80);
            var repacked = new List<TextChunk>(chunks.Count);
            var softParagraphMerges = 0;

            var currentText = chunks[0].Text.Trim();
            var currentParagraphEnd = chunks[0].ParagraphEnd;
            var currentIsDialogue = chunks[0].IsDialogue;
            var currentTokens = SafeCountPromptTokens(counter, currentText);
            var currentSeconds = EstimateSeconds(currentText);
            var currentWords = CountWords(currentText);
            var overBudgetSingles = currentTokens > budget ? 1 : 0;
            var maxPromptTokens = currentTokens;
            var maxSeconds = currentSeconds;
            var maxWords = currentWords;

            for (var i = 1; i < chunks.Count; i++)
            {
                var next = chunks[i];
                var nextText = next.Text.Trim();
                if (string.IsNullOrWhiteSpace(nextText))
                {
                    continue;
                }

                // Preserve paragraph boundaries by default, but allow merging tiny same-mode chunks
                // across paragraph breaks to avoid excessive fragmentation in dialogue-heavy chapters.
                if (currentParagraphEnd)
                {
                    var nextTokensForBoundary = SafeCountPromptTokens(counter, nextText);
                    var nextSecondsForBoundary = EstimateSeconds(nextText);
                    var nextWordsForBoundary = CountWords(nextText);
                    var candidateAcrossParagraph = $"{currentText} {nextText}".Trim();
                    var candidateAcrossParagraphTokens = SafeCountPromptTokens(counter, candidateAcrossParagraph);
                    var candidateAcrossParagraphSeconds = EstimateSeconds(candidateAcrossParagraph);
                    var candidateAcrossParagraphWords = CountWords(candidateAcrossParagraph);
                    var tinyCurrent = currentWords < Math.Max(12, wordCap / 3) || currentSeconds < Math.Max(2.6, secondsCap * 0.28);
                    var tinyNext = nextWordsForBoundary < Math.Max(12, wordCap / 3) || nextSecondsForBoundary < Math.Max(2.6, secondsCap * 0.28);
                    var canSoftMergeParagraph =
                        currentIsDialogue == next.IsDialogue &&
                        (tinyCurrent || tinyNext) &&
                        candidateAcrossParagraphTokens <= budget &&
                        candidateAcrossParagraphSeconds <= secondsCap &&
                        candidateAcrossParagraphWords <= wordCap;

                    if (canSoftMergeParagraph)
                    {
                        currentText = candidateAcrossParagraph;
                        currentParagraphEnd = next.ParagraphEnd;
                        currentIsDialogue = next.IsDialogue;
                        currentTokens = candidateAcrossParagraphTokens;
                        currentSeconds = candidateAcrossParagraphSeconds;
                        currentWords = candidateAcrossParagraphWords;
                        maxPromptTokens = Math.Max(maxPromptTokens, currentTokens);
                        maxSeconds = Math.Max(maxSeconds, currentSeconds);
                        maxWords = Math.Max(maxWords, currentWords);
                        softParagraphMerges++;
                        continue;
                    }

                    repacked.Add(new TextChunk(currentText, true, currentIsDialogue));
                    currentText = nextText;
                    currentParagraphEnd = next.ParagraphEnd;
                    currentIsDialogue = next.IsDialogue;
                    currentTokens = nextTokensForBoundary;
                    currentSeconds = nextSecondsForBoundary;
                    currentWords = nextWordsForBoundary;
                    if (currentTokens > budget)
                    {
                        overBudgetSingles++;
                    }
                    maxPromptTokens = Math.Max(maxPromptTokens, currentTokens);
                    maxSeconds = Math.Max(maxSeconds, currentSeconds);
                    maxWords = Math.Max(maxWords, currentWords);
                    continue;
                }

                if (currentIsDialogue != next.IsDialogue)
                {
                    repacked.Add(new TextChunk(currentText, currentParagraphEnd, currentIsDialogue));
                    currentText = nextText;
                    currentParagraphEnd = next.ParagraphEnd;
                    currentIsDialogue = next.IsDialogue;
                    currentTokens = SafeCountPromptTokens(counter, currentText);
                    currentSeconds = EstimateSeconds(currentText);
                    currentWords = CountWords(currentText);
                    if (currentTokens > budget)
                    {
                        overBudgetSingles++;
                    }
                    maxPromptTokens = Math.Max(maxPromptTokens, currentTokens);
                    maxSeconds = Math.Max(maxSeconds, currentSeconds);
                    maxWords = Math.Max(maxWords, currentWords);
                    continue;
                }

                var candidate = $"{currentText} {nextText}".Trim();
                var candidateTokens = SafeCountPromptTokens(counter, candidate);
                var candidateSeconds = EstimateSeconds(candidate);
                var candidateWords = CountWords(candidate);
                if (candidateTokens <= budget && candidateSeconds <= secondsCap && candidateWords <= wordCap)
                {
                    currentText = candidate;
                    currentParagraphEnd = next.ParagraphEnd;
                    currentIsDialogue = next.IsDialogue;
                    currentTokens = candidateTokens;
                    currentSeconds = candidateSeconds;
                    currentWords = candidateWords;
                    maxPromptTokens = Math.Max(maxPromptTokens, currentTokens);
                    maxSeconds = Math.Max(maxSeconds, currentSeconds);
                    maxWords = Math.Max(maxWords, currentWords);
                    continue;
                }

                repacked.Add(new TextChunk(currentText, currentParagraphEnd, currentIsDialogue));
                currentText = nextText;
                currentParagraphEnd = next.ParagraphEnd;
                currentIsDialogue = next.IsDialogue;
                currentTokens = SafeCountPromptTokens(counter, currentText);
                currentSeconds = EstimateSeconds(currentText);
                currentWords = CountWords(currentText);
                if (currentTokens > budget)
                {
                    overBudgetSingles++;
                }
                maxPromptTokens = Math.Max(maxPromptTokens, currentTokens);
                maxSeconds = Math.Max(maxSeconds, currentSeconds);
                maxWords = Math.Max(maxWords, currentWords);
            }

            if (!string.IsNullOrWhiteSpace(currentText))
            {
                repacked.Add(new TextChunk(currentText, currentParagraphEnd, currentIsDialogue));
            }

            var minPromptTokens = counter.PromptOverheadTokens + 3;
            var minChunkWords = isGpuPreferred ? 8 : 10;
            var minChunkSeconds = isGpuPreferred ? 2.0 : 2.4;
            var beforeTinyMerge = repacked.Count;
            repacked = MergeQwenTooShortChunks(repacked, counter, minPromptTokens, minChunkWords, minChunkSeconds);

            TraceGenerate(
                $"Qwen token-aware repack: chunks {chunks.Count}->{repacked.Count}, " +
                $"promptBudget={budget}, promptOverhead={counter.PromptOverheadTokens}, " +
                $"device={preferDevice}, " +
                $"maxPromptTokens={maxPromptTokens}, secCap={secondsCap:0.#}, wordCap={wordCap}, " +
                $"maxChunkSec={maxSeconds:0.#}, maxChunkWords={maxWords}, " +
                $"maxPos={counter.MaxPositionEmbeddings}, maxNew={counter.MaxNewTokens}, overBudgetSingles={overBudgetSingles}, " +
                $"minPromptTokens={minPromptTokens}, minChunkWords={minChunkWords}, minChunkSec={minChunkSeconds:0.#}, " +
                $"softParagraphMerges={softParagraphMerges}, tinyMerge={beforeTinyMerge}->{repacked.Count}");

            return repacked;
        }
    }

    private static int SafeCountPromptTokens(QwenTokenCounter counter, string text)
    {
        try
        {
            return counter.CountAssistantPromptTokens(text);
        }
        catch
        {
            // If token counting fails mid-run, keep chunk rather than crashing generation.
            return int.MaxValue / 4;
        }
    }

    private static List<TextChunk> MergeQwenTooShortChunks(
        List<TextChunk> chunks,
        QwenTokenCounter counter,
        int minPromptTokens,
        int minChunkWords,
        double minChunkSeconds)
    {
        if (chunks.Count <= 1)
        {
            return chunks;
        }

        var list = new List<TextChunk>(chunks);
        var changed = true;
        while (changed && list.Count > 1)
        {
            changed = false;
            for (var i = 0; i < list.Count; i++)
            {
                var tokens = SafeCountPromptTokens(counter, list[i].Text);
                var words = CountWords(list[i].Text);
                var secs = EstimateSeconds(list[i].Text);
                if (tokens > minPromptTokens && words >= minChunkWords && secs >= minChunkSeconds)
                {
                    continue;
                }

                if (i + 1 < list.Count && list[i + 1].IsDialogue == list[i].IsDialogue)
                {
                    var merged = $"{list[i].Text} {list[i + 1].Text}".Trim();
                    list[i + 1] = new TextChunk(merged, list[i + 1].ParagraphEnd, list[i].IsDialogue);
                    list.RemoveAt(i);
                    changed = true;
                    break;
                }

                if (i > 0 && list[i - 1].IsDialogue == list[i].IsDialogue)
                {
                    var merged = $"{list[i - 1].Text} {list[i].Text}".Trim();
                    list[i - 1] = new TextChunk(merged, list[i].ParagraphEnd, list[i].IsDialogue);
                    list.RemoveAt(i);
                    changed = true;
                    break;
                }

                // Emergency fallback only if no same-mode neighbor exists (avoids tokenizer-too-short hard failures).
                if (i + 1 < list.Count)
                {
                    var merged = $"{list[i].Text} {list[i + 1].Text}".Trim();
                    list[i + 1] = new TextChunk(merged, list[i + 1].ParagraphEnd, list[i + 1].IsDialogue);
                    list.RemoveAt(i);
                    changed = true;
                    break;
                }

                if (i > 0)
                {
                    var merged = $"{list[i - 1].Text} {list[i].Text}".Trim();
                    list[i - 1] = new TextChunk(merged, list[i].ParagraphEnd, list[i - 1].IsDialogue);
                    list.RemoveAt(i);
                    changed = true;
                    break;
                }
            }
        }

        return list;
    }

    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return 0;
        }

        return Regex.Matches(text, @"\b\w+\b").Count;
    }

    private void TraceQwenChunkTrace(string phase, IReadOnlyList<TextChunk> chunks)
    {
        TraceGenerate($"[QWEN CHUNK TRACE] phase={phase} count={chunks.Count}");
        for (var i = 0; i < chunks.Count; i++)
        {
            var c = chunks[i];
            TraceGenerate(
                $"[QWEN CHUNK TRACE] phase={phase} idx={i + 1}/{chunks.Count} " +
                $"chars={c.Text.Length}, words={CountWords(c.Text)}, estSec={EstimateSeconds(c.Text):0.0}, " +
                $"dialogue={(c.IsDialogue ? 1 : 0)}, paraEnd={(c.ParagraphEnd ? 1 : 0)}, " +
                $"hash={StableTextHashHex(c.Text)}, preview={QuoteForLog(ChunkPreview(c.Text))}");
        }
    }

    private static string ChunkPreview(string text, int edgeChars = 80)
    {
        var s = Regex.Replace((text ?? string.Empty).Replace("\r\n", " ").Replace('\n', ' '), @"\s+", " ").Trim();
        if (s.Length <= edgeChars * 2 + 8)
        {
            return s;
        }

        return $"{s[..edgeChars]} ... {s[^edgeChars..]}";
    }

    private static string QuoteForLog(string text)
    {
        var s = (text ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        return "\"" + s + "\"";
    }

    private static string StableTextHashHex(string text)
    {
        unchecked
        {
            ulong hash = 1469598103934665603UL; // FNV-1a 64
            foreach (var ch in text ?? string.Empty)
            {
                hash ^= ch;
                hash *= 1099511628211UL;
            }
            return hash.ToString("X16", CultureInfo.InvariantCulture);
        }
    }

    private static List<TextChunk> BuildTextChunksByTime(IReadOnlyList<string> paragraphs, SynthesisSettings settings, ChunkProfile profile)
    {
        var narrationTarget = Math.Clamp(settings.NarrationTargetSec, 2.0, 12.0);
        var dialogueTarget = Math.Clamp(settings.DialogueTargetSec, 2.0, 10.0);
        var overflow = Math.Clamp(settings.DialogueOverflow, 1.1, 2.5);
        var hardMaxChars = Math.Clamp(settings.MaxChars, 120, 640);

        if (profile == ChunkProfile.Qwen3Tts)
        {
            // Qwen can handle longer segments than chatterbox; fewer boundaries reduces dialogue artifacts.
            // Max capped at 2200 to avoid long silent/murmur generation on the 0.68B split model.
            narrationTarget = Math.Clamp(Math.Max(narrationTarget * 3.0, 14.0), 8.0, 34.0);
            dialogueTarget = Math.Clamp(Math.Max(dialogueTarget * 3.0, 10.0), 6.0, 26.0);
            overflow = Math.Clamp(Math.Max(overflow, 2.8), 2.0, 4.2);
            hardMaxChars = Math.Clamp(Math.Max(settings.MaxChars * 4, 1600), 800, 2200);
        }
        else if (profile == ChunkProfile.KittenTts)
        {
            // Kitten wrapper uses ~400-char chunking; keep chunks moderate and punctuation-friendly.
            narrationTarget = Math.Clamp(Math.Max(narrationTarget * 1.8, 8.5), 5.0, 16.0);
            dialogueTarget = Math.Clamp(Math.Max(dialogueTarget * 1.6, 6.0), 4.0, 12.0);
            overflow = Math.Clamp(Math.Max(overflow, 2.0), 1.4, 3.2);
            hardMaxChars = Math.Clamp(Math.Max(settings.MaxChars, 380), 260, 460);
        }

        var chunks = new List<TextChunk>();
        foreach (var paragraph in paragraphs)
        {
            var isDialogue = DialogueParagraphPattern.IsMatch(paragraph);
            var target = isDialogue ? dialogueTarget : narrationTarget;
            var sentences = SplitSentences(paragraph, settings.ProtectBracketDirectives);
            if (sentences.Count == 0)
            {
                sentences.Add(paragraph);
            }
            var timedUnits = BuildModeAwareTimedUnits(sentences, isDialogue, narrationTarget, dialogueTarget, overflow, hardMaxChars);
            if (timedUnits.Count == 0)
            {
                timedUnits.Add(new TimedUnit(paragraph.Trim(), isDialogue));
            }

            var packed = PackTimedUnitsByTimeAndChars(timedUnits, narrationTarget, dialogueTarget, hardMaxChars);
            for (var i = 0; i < packed.Count; i++)
            {
                chunks.Add(new TextChunk(packed[i].Text, i == packed.Count - 1, packed[i].IsDialogue));
            }
        }

        return chunks;
    }

    private static bool IsDialogueSentenceUnit(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var s = text.Trim();
        if (DialogueParagraphPattern.IsMatch(s) || DialogueTagPattern.IsMatch(s))
        {
            return true;
        }

        // Inline quoted phrases often appear inside narration sentences; treating the whole sentence
        // as dialogue causes over-fragmentation and many tiny Qwen chunks.
        return false;
    }

    private static List<TimedUnit> PackTimedUnitsByTimeAndChars(
        IReadOnlyList<TimedUnit> units,
        double narrationTargetSec,
        double dialogueTargetSec,
        int hardMaxChars)
    {
        var packed = new List<TimedUnit>();
        var current = string.Empty;
        var currentSec = 0.0;
        var currentIsDialogue = false;

        foreach (var unit in units)
        {
            var text = unit.Text.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(current))
            {
                current = text;
                currentSec = EstimateSeconds(text);
                currentIsDialogue = unit.IsDialogue;
                if (unit.ForceBoundaryAfter)
                {
                    packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
                    current = string.Empty;
                    currentSec = 0.0;
                }
                continue;
            }

            var targetSec = currentIsDialogue ? dialogueTargetSec : narrationTargetSec;
            var candidate = $"{current} {text}".Trim();
            var candidateSec = currentSec + EstimateSeconds(text);
            var modeBoundary = currentIsDialogue != unit.IsDialogue;
            var wouldOverflow = candidateSec > targetSec || candidate.Length > hardMaxChars;

            if (modeBoundary || wouldOverflow)
            {
                packed.Add(new TimedUnit(current.Trim(), currentIsDialogue));
                current = text;
                currentSec = EstimateSeconds(text);
                currentIsDialogue = unit.IsDialogue;
                continue;
            }

            current = candidate;
            currentSec = candidateSec;
            if (unit.ForceBoundaryAfter)
            {
                packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
                current = string.Empty;
                currentSec = 0.0;
            }
        }

        if (!string.IsNullOrWhiteSpace(current))
        {
            packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
        }

        return packed;
    }

    private static List<TimedUnit> PackTimedUnitsByChars(IReadOnlyList<TimedUnit> units, int maxChars)
    {
        var packed = new List<TimedUnit>();
        var current = string.Empty;
        var currentIsDialogue = false;

        foreach (var unit in units)
        {
            var text = unit.Text.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(current))
            {
                current = text;
                currentIsDialogue = unit.IsDialogue;
                if (unit.ForceBoundaryAfter)
                {
                    packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
                    current = string.Empty;
                }
                continue;
            }

            var candidate = $"{current} {text}".Trim();
            var modeBoundary = currentIsDialogue != unit.IsDialogue;
            var wouldOverflow = candidate.Length > maxChars;

            if (modeBoundary || wouldOverflow)
            {
                packed.Add(new TimedUnit(current.Trim(), currentIsDialogue));
                current = text;
                currentIsDialogue = unit.IsDialogue;
                continue;
            }

            current = candidate;
            if (unit.ForceBoundaryAfter)
            {
                packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
                current = string.Empty;
            }
        }

        if (!string.IsNullOrWhiteSpace(current))
        {
            packed.Add(new TimedUnit(current.Trim(), currentIsDialogue, true));
        }

        return packed;
    }

    private static List<TimedUnit> BuildModeAwareTimedUnits(
        IReadOnlyList<string> sentences,
        bool paragraphFallbackIsDialogue,
        double narrationTarget,
        double dialogueTarget,
        double overflow,
        int hardMaxChars)
    {
        var timedUnits = new List<TimedUnit>();
        foreach (var sentence in sentences)
        {
            var sentenceTrim = sentence.Trim();
            if (string.IsNullOrWhiteSpace(sentenceTrim))
            {
                continue;
            }

            var sentenceIsDialogue = IsDialogueSentenceUnit(sentenceTrim);
            var sentenceTarget = sentenceIsDialogue ? dialogueTarget : narrationTarget;
            var wrapped = sentenceIsDialogue
                ? SplitDialogueAtomic(sentenceTrim, sentenceTarget, overflow)
                : WrapLongNarrative(sentenceTrim, sentenceTarget, overflow);

            foreach (var raw in wrapped)
            {
                var clean = raw.Trim();
                if (string.IsNullOrWhiteSpace(clean))
                {
                    continue;
                }

                if (clean.Length > hardMaxChars)
                {
                    foreach (var split in SplitLongUnit(clean, hardMaxChars))
                    {
                        var s = split.Trim();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            foreach (var clause in SplitPauseAwareClauses(s))
                            {
                                timedUnits.Add(new TimedUnit(clause.Text, sentenceIsDialogue, clause.ForceBoundaryAfter));
                            }
                        }
                    }
                }
                else
                {
                    foreach (var clause in SplitPauseAwareClauses(clean))
                    {
                        timedUnits.Add(new TimedUnit(clause.Text, sentenceIsDialogue, clause.ForceBoundaryAfter));
                    }
                }
            }
        }

        if (timedUnits.Count == 0)
        {
            timedUnits.Add(new TimedUnit(string.Join(" ", sentences).Trim(), paragraphFallbackIsDialogue, true));
        }

        return timedUnits;
    }

    private static List<TextChunk> BuildTextChunksByChars(IReadOnlyList<string> paragraphs, SynthesisSettings settings)
    {
        var minChars = Math.Clamp(settings.MinChars, 40, 2400);
        var maxChars = Math.Clamp(settings.MaxChars, minChars, 2400);
        var manualMax = Math.Clamp(settings.ManualMaxChars, 80, 6000);
        var maxPerChunk = Math.Max(minChars, Math.Min(maxChars, manualMax));
        var narrationTarget = Math.Clamp(settings.NarrationTargetSec, 2.0, 12.0);
        var overflow = Math.Clamp(settings.DialogueOverflow, 1.1, 2.5);

        var chunks = new List<TextChunk>();
        foreach (var paragraph in paragraphs)
        {
            var isDialogue = DialogueParagraphPattern.IsMatch(paragraph);
            var sentences = SplitSentences(paragraph, settings.ProtectBracketDirectives);
            if (sentences.Count == 0)
            {
                sentences.Add(paragraph);
            }

            var timedUnits = BuildModeAwareTimedUnits(sentences, isDialogue, narrationTarget, narrationTarget, overflow, maxPerChunk);
            var packed = PackTimedUnitsByChars(timedUnits, maxPerChunk);
            for (var i = 0; i < packed.Count; i++)
            {
                chunks.Add(new TextChunk(packed[i].Text, i == packed.Count - 1, packed[i].IsDialogue));
            }
        }

        return chunks;
    }

    private string ApplyLocalInstruction(string text, SynthesisSettings settings, string? continuityTail = null, bool isDialogueChunk = false)
    {
        var repo = _config.ModelRepoId ?? string.Empty;
        if (!repo.Contains("qwen3-tts", StringComparison.OrdinalIgnoreCase))
        {
            return text;
        }

        if (IsLocalQwenCustomVoiceMode())
        {
            return text;
        }

        var manualHint = (settings.LocalInstructionHint ?? string.Empty).Trim();
        // Do not inject automatic style/continuity bracket prompts into Qwen ONNX text.
        // In practice this can be spoken aloud by the model and can destabilize chunk generation.
        if (string.IsNullOrWhiteSpace(manualHint))
        {
            return text;
        }

        return $"[instruction: {manualHint}] {text}";
    }

    private static string? BuildChunkContinuityTail(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var s = Regex.Replace(text, @"\s+", " ").Trim();
        if (s.Length <= 140)
        {
            return s;
        }

        var tail = s[^140..];
        var split = tail.IndexOf(' ');
        if (split > 0 && split < tail.Length - 1)
        {
            tail = tail[(split + 1)..];
        }

        return tail.Trim();
    }

    private static List<string> SplitSentences(string paragraph, bool bracketAware)
    {
        if (string.IsNullOrWhiteSpace(paragraph))
        {
            return new List<string>();
        }

        var results = new List<string>();
        var sb = new StringBuilder();
        var bracketDepth = 0;
        for (var i = 0; i < paragraph.Length; i++)
        {
            var ch = paragraph[i];
            sb.Append(ch);
            if (ch == '[')
            {
                bracketDepth++;
            }
            else if (ch == ']' && bracketDepth > 0)
            {
                bracketDepth--;
            }

            if (ch is '.' or '!' or '?')
            {
                if (bracketAware && bracketDepth > 0)
                {
                    continue;
                }

                var hasBoundary = i + 1 >= paragraph.Length || char.IsWhiteSpace(paragraph[i + 1]);
                if (!hasBoundary && i + 2 < paragraph.Length && (paragraph[i + 1] is '"' or '\'' or ')' or ']'))
                {
                    hasBoundary = char.IsWhiteSpace(paragraph[i + 2]);
                    if (hasBoundary)
                    {
                        sb.Append(paragraph[i + 1]);
                        i++;
                    }
                }

                if (hasBoundary)
                {
                    var sentence = sb.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(sentence) && !EndsWithSentenceAbbreviation(sentence))
                    {
                        results.Add(sentence);
                        sb.Clear();
                    }
                }
            }
        }

        var tail = sb.ToString().Trim();
        if (!string.IsNullOrWhiteSpace(tail))
        {
            results.Add(tail);
        }

        return results;
    }

    private static List<(string Text, bool ForceBoundaryAfter)> SplitPauseAwareClauses(string sentence)
    {
        var results = new List<(string Text, bool ForceBoundaryAfter)>();
        if (string.IsNullOrWhiteSpace(sentence))
        {
            return results;
        }

        var sb = new StringBuilder();
        var bracketDepth = 0;
        for (var i = 0; i < sentence.Length; i++)
        {
            var ch = sentence[i];
            sb.Append(ch);

            if (ch == '[')
            {
                bracketDepth++;
            }
            else if (ch == ']' && bracketDepth > 0)
            {
                bracketDepth--;
            }

            if (bracketDepth > 0)
            {
                continue;
            }

            var isEllipsisChar = ch == '\u2026';
            var isThreeDots = ch == '.' && i + 2 < sentence.Length && sentence[i + 1] == '.' && sentence[i + 2] == '.';
            if (isThreeDots)
            {
                sb.Append('.');
                sb.Append('.');
                i += 2;
            }

            var isSentencePunc = ch is '.' or '!' or '?';
            if (!(ch is ',' or ';' or ':' || isEllipsisChar || isThreeDots || isSentencePunc))
            {
                continue;
            }

            var nextIndex = i + 1;
            var hasBoundary = nextIndex >= sentence.Length || char.IsWhiteSpace(sentence[nextIndex]);
            if (!hasBoundary && nextIndex + 1 < sentence.Length && (sentence[nextIndex] is '"' or '\'' or ')' or ']'))
            {
                hasBoundary = char.IsWhiteSpace(sentence[nextIndex + 1]);
                if (hasBoundary)
                {
                    sb.Append(sentence[nextIndex]);
                    i++;
                }
            }

            if (!hasBoundary)
            {
                continue;
            }

            if (isSentencePunc)
            {
                // Avoid splitting on common abbreviations (e.g. "Mr.", "Dr.", "etc.") in this
                // fallback clause splitter. The sentence splitter handles most cases already, but
                // this catches formatting edge cases where sentence punctuation survives here.
                var candidateSentence = sb.ToString().Trim();
                if (EndsWithSentenceAbbreviation(candidateSentence))
                {
                    continue;
                }
            }

            var part = sb.ToString().Trim();
            if (!string.IsNullOrWhiteSpace(part))
            {
                results.Add((part, true));
            }
            sb.Clear();
        }

        var tail = sb.ToString().Trim();
        if (!string.IsNullOrWhiteSpace(tail))
        {
            var forceBoundary = GetTrailingSpeechPunctuation(tail).HasValue || EndsWithEllipsis(tail);
            results.Add((tail, forceBoundary));
        }

        if (results.Count == 0)
        {
            var clean = sentence.Trim();
            results.Add((clean, GetTrailingSpeechPunctuation(clean).HasValue || EndsWithEllipsis(clean)));
        }

        return results;
    }

    private static bool EndsWithSentenceAbbreviation(string sentence)
    {
        var trimmed = sentence.Trim().TrimEnd('"', '\'', ')', ']');
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return false;
        }

        var lastSpace = trimmed.LastIndexOf(' ');
        var tail = (lastSpace >= 0 ? trimmed[(lastSpace + 1)..] : trimmed).ToLowerInvariant();
        return SentenceAbbreviations.Contains(tail);
    }

    private static List<string> SplitDialogueAtomic(string sentence, double targetSec, double overflow)
    {
        var sec = EstimateSeconds(sentence);
        if (DialogueTagPattern.IsMatch(sentence))
        {
            return sec <= targetSec * overflow
                ? new List<string> { sentence.Trim() }
                : SoftWrapByTime(sentence, targetSec, dialogueMode: true).ToList();
        }

        return sec <= targetSec * overflow
            ? new List<string> { sentence.Trim() }
            : SoftWrapByTime(sentence, targetSec, dialogueMode: true).ToList();
    }

    private static List<string> WrapLongNarrative(string sentence, double targetSec, double overflow)
    {
        if (EstimateSeconds(sentence) <= targetSec * overflow)
        {
            return new List<string> { sentence.Trim() };
        }

        var clauses = Regex.Split(sentence, @"([;:]\s+)");
        var output = new List<string>();
        var current = string.Empty;
        for (var i = 0; i < clauses.Length; i += 2)
        {
            var segment = clauses[i] + (i + 1 < clauses.Length ? clauses[i + 1] : string.Empty);
            segment = segment.Trim();
            if (string.IsNullOrWhiteSpace(segment))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(current))
            {
                current = segment;
                continue;
            }

            var combined = $"{current} {segment}".Trim();
            if (EstimateSeconds(combined) <= targetSec)
            {
                current = combined;
            }
            else
            {
                output.Add(current);
                current = segment;
            }
        }

        if (!string.IsNullOrWhiteSpace(current))
        {
            output.Add(current.Trim());
        }

        if (output.Any(part => EstimateSeconds(part) > targetSec * 1.2))
        {
            return SoftWrapByTime(sentence, targetSec).ToList();
        }

        return output;
    }

    private static IEnumerable<string> SoftWrapByTime(string text, double targetSec, bool dialogueMode = false)
    {
        var clean = Regex.Replace((text ?? string.Empty).Trim(), @"\s+", " ");
        if (string.IsNullOrWhiteSpace(clean))
        {
            yield break;
        }

        var secs = Math.Max(EstimateSeconds(clean), 0.1);
        var cps = clean.Length / secs;
        var width = Math.Max(40, (int)Math.Round(cps * Math.Max(1.0, targetSec)));
        foreach (var part in SplitLongUnit(clean, width, preferCommaBreaks: !dialogueMode))
        {
            yield return part;
        }
    }

    private static double EstimateSeconds(string text)
    {
        var source = text ?? string.Empty;
        var words = Regex.Matches(source, @"\b\w+\b").Count;
        if (words == 0)
        {
            return 0.15;
        }

        var wps = 3.8;
        wps -= 0.08 * Regex.Matches(source, @"\d").Count;
        wps -= 0.12 * (source.Count(c => c is ':' or '/' or '(' or ')'));
        wps -= 0.06 * Regex.Matches(source, @"\b[A-Z]{2,}\b").Count;
        wps = Math.Clamp(wps, 3.0, 4.4);
        return Math.Max(0.15, words / wps);
    }

    private static IEnumerable<string> SplitLongUnit(string text, int maxChars, bool preferCommaBreaks = true)
    {
        var remaining = text.Trim();
        while (remaining.Length > maxChars)
        {
            var split = FindSplitIndex(remaining, maxChars, preferCommaBreaks);
            if (split <= 0 || split >= remaining.Length)
            {
                split = maxChars;
            }

            var part = remaining[..split].Trim();
            if (!string.IsNullOrWhiteSpace(part))
            {
                yield return part;
            }

            remaining = remaining[split..].TrimStart();
        }

        if (!string.IsNullOrWhiteSpace(remaining))
        {
            yield return remaining;
        }
    }

    private static int FindSplitIndex(string text, int maxChars, bool preferCommaBreaks = true)
    {
        var limit = Math.Min(maxChars, text.Length - 1);
        var punctuation = preferCommaBreaks ? new[] { ';', ':', ',' } : new[] { ';', ':' };
        foreach (var ch in punctuation)
        {
            var idx = text.LastIndexOf(ch, limit);
            if (idx >= maxChars / 2)
            {
                return idx + 1;
            }
        }

        for (var i = limit; i >= maxChars / 2; i--)
        {
            if ((text[i] == '"' || text[i] == '\'') &&
                i + 1 < text.Length &&
                char.IsWhiteSpace(text[i + 1]))
            {
                return i + 1;
            }
        }

        var space = text.LastIndexOf(' ', limit);
        if (space >= maxChars / 2)
        {
            return space + 1;
        }

        return limit;
    }

    private static string NormalizeTextForQwenTts(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return text ?? string.Empty;
        }

        var s = text;

        // Repair common mojibake from scraped/transcoded chapters.
        s = s.Replace("Ã¢â‚¬Å“", "\"", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬Â", "\"", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬Ëœ", "'", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬â„¢", "'", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬Â¦", "\u2026", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬â€", "-", StringComparison.Ordinal)
             .Replace("Ã¢â‚¬â€œ", "-", StringComparison.Ordinal)
             .Replace("Ã‚ ", " ", StringComparison.Ordinal)
             .Replace("Ã‚", string.Empty, StringComparison.Ordinal);

        // Normalize punctuation for more stable dialogue tokenization.
        s = s.Replace("\u201C", "\"", StringComparison.Ordinal)
             .Replace("\u201D", "\"", StringComparison.Ordinal)
             .Replace("\u2018", "'", StringComparison.Ordinal)
             .Replace("\u2019", "'", StringComparison.Ordinal)
             .Replace("\u2014", "-", StringComparison.Ordinal)
             .Replace("\u2013", "-", StringComparison.Ordinal)
             .Replace("\u00A0", " ", StringComparison.Ordinal);

        // Keep whitespace stable but avoid bizarre spacing around quotes.
        s = s.Replace("\r\n", "\n", StringComparison.Ordinal)
             .Replace("\r", "\n", StringComparison.Ordinal);
        s = Regex.Replace(s, @"[ \t]+", " ");
        s = Regex.Replace(s, @" ?([""']) ", " $1");
        s = Regex.Replace(s, @"([""'])\s{2,}", "$1 ");

        return s;
    }

    private static string NormalizeTextForKittenTts(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return text ?? string.Empty;
        }

        var s = text.Replace("\r\n", "\n", StringComparison.Ordinal)
                    .Replace("\r", "\n", StringComparison.Ordinal);
        s = s.Replace("\u201C", "\"", StringComparison.Ordinal)
             .Replace("\u201D", "\"", StringComparison.Ordinal)
             .Replace("\u2018", "'", StringComparison.Ordinal)
             .Replace("\u2019", "'", StringComparison.Ordinal)
             .Replace("\u2014", "-", StringComparison.Ordinal)
             .Replace("\u2013", "-", StringComparison.Ordinal)
             .Replace("\u00A0", " ", StringComparison.Ordinal);
        s = Regex.Replace(s, @"[ \t]+", " ");
        s = Regex.Replace(s, @"\n{3,}", "\n\n");
        return s.Trim();
    }

    private static void MergeChunkWavs(
        IReadOnlyList<string> chunkWavs,
        IReadOnlyList<TextChunk> chunks,
        string outputPath,
        SynthesisSettings settings,
        ChunkProfile chunkProfile,
        bool trimInteriorSilences = false)
    {
        if (chunkWavs.Count == 0 || chunkWavs.Count != chunks.Count)
        {
            throw new InvalidOperationException("Chunk synthesis output mismatch.");
        }

        WaveFormat? format = null;
        using var writerHolder = new WriterHolder();
        for (var i = 0; i < chunkWavs.Count; i++)
        {
            using var reader = new WaveFileReader(chunkWavs[i]);
            if (format is null)
            {
                format = reader.WaveFormat;
                writerHolder.Writer = new WaveFileWriter(outputPath, format);
            }
            else if (!WaveFormatEquals(format, reader.WaveFormat))
            {
                throw new InvalidOperationException("Generated chunks have incompatible WAV formats.");
            }

            var data = ReadWaveBytes(reader);
            data = TrimWaveBoundarySilenceIfSupported(
                data,
                format,
                trimLeading: i > 0,
                trimTrailing: i < chunkWavs.Count - 1,
                maxLeadingTrimMs: chunks[i].IsDialogue ? 180 : 240,
                maxTrailingTrimMs: chunks[i].ParagraphEnd ? 550 : 320);
            // Avoid touching the first chunk's interior timing. Qwen sometimes starts with a soft/quiet onset,
            // and interior-silence capping can trim/compress that opening in a way that sounds like missing words.
            if (trimInteriorSilences && i > 0)
            {
                data = TrimLongInteriorSilences(data, format);
            }

            var fadeInMs = ResolveBoundaryFadeInMs(chunks, i);
            var fadeOutMs = ResolveBoundaryFadeOutMs(chunks, i);
            ApplyBoundaryFadesIfSupported(data, format, fadeInMs, fadeOutMs);
            writerHolder.Writer!.Write(data, 0, data.Length);

            if (i < chunkWavs.Count - 1)
            {
                var pauseMs = ResolveChunkBoundaryPauseMs(chunks[i], i, settings, chunkProfile);
                WriteSilence(writerHolder.Writer!, pauseMs);
            }
        }
    }

    private static int ResolveChunkBoundaryPauseMs(TextChunk chunk, int chunkIndex, SynthesisSettings settings, ChunkProfile chunkProfile)
    {
        var (defaultClauseRange, defaultSentenceRange, defaultParagraphRange, defaultEllipsisRange) = chunkProfile switch
        {
            // Qwen can carry natural prosody better than Chatterbox, so use moderate floor pauses.
            ChunkProfile.Qwen3Tts => ((180, 260), (420, 620), (820, 1200), (520, 760)),
            // Kitten benefits from slightly stronger punctuation pauses to improve readability.
            ChunkProfile.KittenTts => ((170, 240), (400, 580), (780, 1120), (500, 720)),
            // Chatterbox often reads punctuation too tightly inside chunked narration. Raise floors
            // so joins sound like real reading instead of near-continuous speech.
            _ => ((190, 280), (480, 700), (900, 1350), (560, 820))
        };
        var clauseRange = ResolvePauseRange(settings.ClausePauseMinMs, settings.ClausePauseMaxMs, defaultClauseRange);
        var sentenceRange = ResolvePauseRange(settings.SentencePauseMinMs, settings.SentencePauseMaxMs, defaultSentenceRange);
        var paragraphRange = ResolvePauseRange(settings.ParagraphPauseMinMs, settings.ParagraphPauseMaxMs, defaultParagraphRange);
        var ellipsisRange = ResolvePauseRange(settings.EllipsisPauseMinMs, settings.EllipsisPauseMaxMs, defaultEllipsisRange);
        var chunkPauseMs = settings.ChunkPauseMs;
        var paragraphPauseMs = settings.ParagraphPauseMs;

        if (chunk.ParagraphEnd)
        {
            var sampled = ResolveRandomizedPauseMs(chunk, chunkIndex, chunkProfile, paragraphRange, "paragraph");
            return Math.Max(Math.Max(0, paragraphPauseMs), sampled);
        }

        if (EndsWithEllipsis(chunk.Text))
        {
            var sampled = ResolveRandomizedPauseMs(chunk, chunkIndex, chunkProfile, ellipsisRange, "ellipsis");
            return Math.Max(Math.Max(0, chunkPauseMs), sampled);
        }

        var ending = GetTrailingSpeechPunctuation(chunk.Text);
        return ending switch
        {
            '.' or '!' or '?' => Math.Max(Math.Max(0, chunkPauseMs), ResolveRandomizedPauseMs(chunk, chunkIndex, chunkProfile, sentenceRange, "sentence")),
            ',' or ';' or ':' => Math.Max(Math.Max(0, chunkPauseMs), ResolveRandomizedPauseMs(chunk, chunkIndex, chunkProfile, clauseRange, "clause")),
            _ => 0
        };
    }

    private static int ResolveRandomizedPauseMs(TextChunk chunk, int chunkIndex, ChunkProfile chunkProfile, (int Min, int Max) range, string category)
    {
        var min = Math.Max(0, range.Min);
        var max = Math.Max(min, range.Max);
        if (max <= min)
        {
            return min;
        }

        var seed = ComputeStablePauseSeed(chunk, chunkIndex, chunkProfile, category);
        var span = max - min + 1;
        return min + (int)(seed % (uint)span);
    }

    private static ((int Min, int Max) Clause, (int Min, int Max) Sentence, (int Min, int Max) Paragraph, (int Min, int Max) Ellipsis)
        GetDefaultPauseRanges(ChunkProfile chunkProfile)
    {
        return chunkProfile switch
        {
            ChunkProfile.Qwen3Tts => ((180, 260), (420, 620), (820, 1200), (520, 760)),
            ChunkProfile.KittenTts => ((170, 240), (400, 580), (780, 1120), (500, 720)),
            _ => ((190, 280), (480, 700), (900, 1350), (560, 820))
        };
    }

    private static (int Min, int Max) ResolvePauseRange(int configuredMin, int configuredMax, (int Min, int Max) fallback)
    {
        var min = configuredMin > 0 ? configuredMin : fallback.Min;
        var max = configuredMax > 0 ? configuredMax : fallback.Max;
        min = Math.Max(0, min);
        max = Math.Max(min, max);
        return (min, max);
    }

    private static void NormalizePauseRangeOrder(SynthesisSettings settings)
    {
        var cmin = settings.ClausePauseMinMs; var cmax = settings.ClausePauseMaxMs; NormalizePair(ref cmin, ref cmax); settings.ClausePauseMinMs = cmin; settings.ClausePauseMaxMs = cmax;
        var smin = settings.SentencePauseMinMs; var smax = settings.SentencePauseMaxMs; NormalizePair(ref smin, ref smax); settings.SentencePauseMinMs = smin; settings.SentencePauseMaxMs = smax;
        var emin = settings.EllipsisPauseMinMs; var emax = settings.EllipsisPauseMaxMs; NormalizePair(ref emin, ref emax); settings.EllipsisPauseMinMs = emin; settings.EllipsisPauseMaxMs = emax;
        var pmin = settings.ParagraphPauseMinMs; var pmax = settings.ParagraphPauseMaxMs; NormalizePair(ref pmin, ref pmax); settings.ParagraphPauseMinMs = pmin; settings.ParagraphPauseMaxMs = pmax;

        static void NormalizePair(ref int min, ref int max)
        {
            min = Math.Max(0, min);
            max = Math.Max(0, max);
            if (max > 0 && min == 0)
            {
                min = max;
            }
            if (min > 0 && max == 0)
            {
                max = min;
            }
            if (max < min)
            {
                (min, max) = (max, min);
            }
        }
    }

    // Deterministic jitter keeps pauses human-like without making the same project sound different every rerun.
    private static uint ComputeStablePauseSeed(TextChunk chunk, int chunkIndex, ChunkProfile chunkProfile, string category)
    {
        unchecked
        {
            uint hash = 2166136261;

            static uint Mix(uint h, uint value)
            {
                unchecked
                {
                    h ^= value;
                    return h * 16777619;
                }
            }

            hash = Mix(hash, (uint)chunkIndex);
            hash = Mix(hash, (uint)chunkProfile);
            hash = Mix(hash, chunk.ParagraphEnd ? 1u : 0u);
            hash = Mix(hash, chunk.IsDialogue ? 1u : 0u);

            for (var i = 0; i < category.Length; i++)
            {
                hash = Mix(hash, category[i]);
            }

            var text = chunk.Text ?? string.Empty;
            for (var i = 0; i < text.Length; i++)
            {
                hash = Mix(hash, text[i]);
            }

            return hash;
        }
    }

    private static bool EndsWithEllipsis(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var trimmed = text.TrimEnd();
        while (trimmed.Length > 0 && (trimmed[^1] == '"' || trimmed[^1] == '\'' || trimmed[^1] == ')' || trimmed[^1] == ']'))
        {
            trimmed = trimmed[..^1].TrimEnd();
        }

        return trimmed.EndsWith("...", StringComparison.Ordinal) || trimmed.EndsWith('\u2026');
    }

    private static char? GetTrailingSpeechPunctuation(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var trimmed = text.TrimEnd();
        while (trimmed.Length > 0 && (trimmed[^1] == '"' || trimmed[^1] == '\'' || trimmed[^1] == ')' || trimmed[^1] == ']'))
        {
            trimmed = trimmed[..^1].TrimEnd();
        }

        if (trimmed.Length == 0)
        {
            return null;
        }

        var last = trimmed[^1];
        return last is '.' or '!' or '?' or ',' or ';' or ':' ? last : null;
    }

    private static int ResolveBoundaryFadeInMs(IReadOnlyList<TextChunk> chunks, int index)
    {
        if (index <= 0 || index >= chunks.Count)
        {
            return 0;
        }

        // Keep joins smoother, especially when chunking forces boundaries mid-flow.
        return chunks[index - 1].IsDialogue == chunks[index].IsDialogue ? 14 : 10;
    }

    private static int ResolveBoundaryFadeOutMs(IReadOnlyList<TextChunk> chunks, int index)
    {
        if (index < 0 || index >= chunks.Count - 1)
        {
            return 0;
        }

        return chunks[index].IsDialogue == chunks[index + 1].IsDialogue ? 14 : 10;
    }

    private static void CopyWave(WaveFileReader reader, WaveFileWriter writer)
    {
        var buffer = new byte[32768];
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            writer.Write(buffer, 0, read);
        }
    }

    private static byte[] ReadWaveBytes(WaveFileReader reader)
    {
        var ms = new MemoryStream();
        var buffer = new byte[32768];
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            ms.Write(buffer, 0, read);
        }

        return ms.ToArray();
    }

    private static bool TryDetectSuspiciousQwenChunkAudio(string wavPath, out string reason)
    {
        reason = string.Empty;
        try
        {
            if (!File.Exists(wavPath))
            {
                reason = "missing chunk wav";
                return true;
            }

            using var reader = new WaveFileReader(wavPath);
            if (reader.Length <= 0)
            {
                reason = "empty wav";
                return true;
            }

            var format = reader.WaveFormat;
            if (format.Encoding != WaveFormatEncoding.Pcm || format.BitsPerSample != 16 || format.BlockAlign <= 0)
            {
                return false; // skip guard on unsupported formats
            }

            var sampleRate = format.SampleRate;
            var channels = Math.Max(1, format.Channels);
            var totalFrames = (int)Math.Min(int.MaxValue, reader.SampleCount / channels);
            if (totalFrames <= 0)
            {
                reason = "no audio frames";
                return true;
            }

            var pcm = ReadWaveBytes(reader);
            var usable = pcm.Length - (pcm.Length % format.BlockAlign);
            if (usable <= 0)
            {
                reason = "no pcm payload";
                return true;
            }

            var windowFrames = Math.Max(1, sampleRate / 10); // 100ms
            var windowBytes = AlignDown(windowFrames * format.BlockAlign, format.BlockAlign);
            if (windowBytes <= 0)
            {
                return false;
            }

            const int silentThreshold = 180;
            const int activeThreshold = 420;
            var windows = 0;
            var silentWindows = 0;
            var activeWindows = 0;
            var longestSilentRun = 0;
            var currentSilentRun = 0;
            var peakAbs = 0;

            for (var offset = 0; offset < usable; offset += windowBytes)
            {
                var len = Math.Min(windowBytes, usable - offset);
                if (len < format.BlockAlign)
                {
                    break;
                }

                long sumAbs = 0;
                var sampleCount = 0;
                for (var i = offset; i + 1 < offset + len; i += 2)
                {
                    short s = (short)(pcm[i] | (pcm[i + 1] << 8));
                    var a = Math.Abs((int)s);
                    sumAbs += a;
                    sampleCount++;
                    if (a > peakAbs)
                    {
                        peakAbs = a;
                    }
                }

                if (sampleCount == 0)
                {
                    continue;
                }

                var meanAbs = (double)sumAbs / sampleCount;
                windows++;
                if (meanAbs < silentThreshold)
                {
                    silentWindows++;
                    currentSilentRun++;
                    if (currentSilentRun > longestSilentRun)
                    {
                        longestSilentRun = currentSilentRun;
                    }
                }
                else
                {
                    currentSilentRun = 0;
                }

                if (meanAbs >= activeThreshold)
                {
                    activeWindows++;
                }
            }

            if (windows == 0)
            {
                reason = "no analysis windows";
                return true;
            }

            var durationSec = totalFrames / (double)sampleRate;
            var silentRatio = silentWindows / (double)windows;
            var activeRatio = activeWindows / (double)windows;
            var longestSilentRunSec = longestSilentRun * 0.1;

            // Only flag clearly bad chunks. Short chunks (e.g. 2â€“5s) often have high silence %
            // from natural pauses; avoid triggering retry/sub-split on those.
            var suspicious =
                peakAbs < 600 ||
                (durationSec >= 8.0 && silentRatio >= 0.88) ||
                (durationSec >= 10.0 && activeRatio <= 0.08) ||
                (durationSec >= 8.0 && longestSilentRunSec >= 5.0 && silentRatio >= 0.70);

            if (!suspicious)
            {
                return false;
            }

            reason = $"dur={durationSec:0.0}s, silent={silentRatio:P0}, active={activeRatio:P0}, longestSilent={longestSilentRunSec:0.0}s, peak={peakAbs}";
            return true;
        }
        catch (Exception ex)
        {
            // Never let diagnostic guard crash generation.
            reason = $"guard-scan-error: {ex.GetType().Name}";
            return false;
        }
    }

    /// <summary>
    /// Last-resort recovery: splits <paramref name="chunkText"/> at a sentence boundary near
    /// the midpoint, generates each half separately, then stitches the two WAVs together into
    /// <paramref name="outputPath"/>.  Returns false only if both halves also look suspicious.
    /// </summary>
    private async Task<bool> TryGenerateQwenSubSplitAsync(
        string chunkText,
        string outputPath,
        string voicePath,
        float speed,
        SynthesisSettings settings,
        ITtsBackend backend,
        string tempRoot,
        int chunkIndex,
        CancellationToken ct)
        => await TryGenerateQwenSubSplitAsync(chunkText, outputPath, voicePath, speed, settings, backend, tempRoot, chunkIndex, ct, depth: 0);

    private async Task<bool> TryGenerateQwenSubSplitAsync(
        string chunkText,
        string outputPath,
        string voicePath,
        float speed,
        SynthesisSettings settings,
        ITtsBackend backend,
        string tempRoot,
        int chunkIndex,
        CancellationToken ct,
        int depth)
    {
        if (depth >= 6 || string.IsNullOrWhiteSpace(chunkText) || chunkText.Length < 24)
        {
            return false;
        }

        // --- find a sentence-boundary split near the midpoint ---
        var mid = chunkText.Length / 2;
        var splitPos = -1;

        // Search outward from mid for a sentence-end punctuation followed by whitespace
        for (var radius = 0; radius < mid; radius++)
        {
            foreach (var delta in new[] { radius, -radius })
            {
                var pos = mid + delta;
                if (pos < 1 || pos >= chunkText.Length) continue;
                var ch = chunkText[pos - 1];
                if ((ch == '.' || ch == '!' || ch == '?') &&
                    pos < chunkText.Length && char.IsWhiteSpace(chunkText[pos]))
                {
                    splitPos = pos;
                    break;
                }
            }
            if (splitPos >= 0) break;
        }

        // Fallback: split at whitespace nearest to midpoint
        if (splitPos < 0)
        {
            for (var radius = 0; radius < mid; radius++)
            {
                foreach (var delta in new[] { radius, -radius })
                {
                    var pos = mid + delta;
                    if (pos >= 1 && pos < chunkText.Length && char.IsWhiteSpace(chunkText[pos]))
                    {
                        splitPos = pos;
                        break;
                    }
                }
                if (splitPos >= 0) break;
            }
        }

        if (splitPos <= 0 || splitPos >= chunkText.Length)
        {
            return false; // can't split sensibly
        }

        var partA = chunkText[..splitPos].Trim();
        var partB = chunkText[splitPos..].Trim();
        if (string.IsNullOrWhiteSpace(partA) || string.IsNullOrWhiteSpace(partB))
        {
            return false;
        }

        var pathA = Path.Combine(tempRoot, $"subsplit_{chunkIndex:D4}_a.wav");
        var pathB = Path.Combine(tempRoot, $"subsplit_{chunkIndex:D4}_b.wav");

        async Task<bool> SynthesizeOrSplitPartAsync(string partText, string partPath, string suffix)
        {
            try
            {
                var req = CreateLocalTtsRequest(
                    ApplyLocalInstruction(partText, settings),
                    voicePath, partPath, speed, settings);
                await backend.SynthesizeAsync(req, ct);
            }
            catch
            {
                return false;
            }

            if (!File.Exists(partPath))
            {
                return false;
            }

            if (!TryDetectSuspiciousQwenChunkAudio(partPath, out var partReason))
            {
                return true;
            }

            TraceGenerate($"Qwen chunk {chunkIndex}: recursive sub-split depth {depth + 1} part {suffix} still suspicious ({partReason}); splitting again.");
            return await TryGenerateQwenSubSplitAsync(
                partText,
                partPath,
                voicePath,
                speed,
                settings,
                backend,
                tempRoot,
                chunkIndex,
                ct,
                depth + 1);
        }

        var okA = await SynthesizeOrSplitPartAsync(partA, pathA, "A");
        var okB = await SynthesizeOrSplitPartAsync(partB, pathB, "B");
        if (!okA || !okB || !File.Exists(pathA) || !File.Exists(pathB))
        {
            return false;
        }

        // Stitch both halves so we never drop content. Previously we only kept "non-suspicious"
        // halves, which dropped one half and produced missing words / mess in the final audio.
        try
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);

            WaveFormat? fmt = null;
            using (var rFmt = new WaveFileReader(pathA))
            {
                fmt = rFmt.WaveFormat;
            }

            using var writer = new WaveFileWriter(outputPath, fmt!);
            foreach (var part in new[] { pathA, pathB })
            {
                using var r = new WaveFileReader(part);
                CopyWave(r, writer);
            }
        }
        catch
        {
            return false;
        }

        return true;
    }

    private static byte[] TrimWaveBoundarySilenceIfSupported(
        byte[] pcmData,
        WaveFormat format,
        bool trimLeading,
        bool trimTrailing,
        int maxLeadingTrimMs,
        int maxTrailingTrimMs)
    {
        if ((!trimLeading && !trimTrailing) || pcmData.Length == 0)
        {
            return pcmData;
        }

        if (format.Encoding != WaveFormatEncoding.Pcm || format.BitsPerSample != 16 || format.BlockAlign <= 0)
        {
            return pcmData;
        }

        var blockAlign = format.BlockAlign;
        var usableLength = pcmData.Length - (pcmData.Length % blockAlign);
        if (usableLength <= blockAlign)
        {
            return pcmData;
        }

        var bytesPerMs = format.AverageBytesPerSecond / 1000.0;
        var maxLeadBytes = trimLeading ? AlignDown((int)Math.Round(Math.Max(0, maxLeadingTrimMs) * bytesPerMs), blockAlign) : 0;
        var maxTrailBytes = trimTrailing ? AlignDown((int)Math.Round(Math.Max(0, maxTrailingTrimMs) * bytesPerMs), blockAlign) : 0;
        maxLeadBytes = Math.Min(maxLeadBytes, usableLength - blockAlign);
        maxTrailBytes = Math.Min(maxTrailBytes, usableLength - blockAlign);

        const int silenceThreshold = 140; // ~ -47 dBFS on int16
        var start = 0;
        if (maxLeadBytes > 0)
        {
            var leadLimit = Math.Min(maxLeadBytes, usableLength - blockAlign);
            while (start + blockAlign <= leadLimit && IsPcm16FrameSilent(pcmData, start, blockAlign, silenceThreshold))
            {
                start += blockAlign;
            }
        }

        var end = usableLength;
        if (maxTrailBytes > 0)
        {
            var trailStartLimit = Math.Max(start + blockAlign, usableLength - maxTrailBytes);
            while (end - blockAlign >= trailStartLimit && IsPcm16FrameSilent(pcmData, end - blockAlign, blockAlign, silenceThreshold))
            {
                end -= blockAlign;
            }
        }

        if (start <= 0 && end >= usableLength)
        {
            return pcmData;
        }

        if (end <= start + blockAlign)
        {
            // Keep at least one frame rather than producing an empty chunk.
            start = 0;
            end = Math.Min(blockAlign, usableLength);
        }

        var trimmed = new byte[end - start];
        Buffer.BlockCopy(pcmData, start, trimmed, 0, trimmed.Length);
        return trimmed;
    }

    private static bool IsPcm16FrameSilent(byte[] data, int offset, int blockAlign, int absThreshold)
    {
        var end = Math.Min(data.Length, offset + blockAlign);
        for (var i = offset; i + 1 < end; i += 2)
        {
            short sample = (short)(data[i] | (data[i + 1] << 8));
            if (Math.Abs(sample) > absThreshold)
            {
                return false;
            }
        }

        return true;
    }

    private static int AlignDown(int value, int align)
    {
        if (align <= 1)
        {
            return Math.Max(0, value);
        }

        var v = Math.Max(0, value);
        return v - (v % align);
    }

    /// <summary>
    /// Scans PCM data for silent runs longer than <paramref name="maxSilenceMs"/> and caps each such run
    /// to <paramref name="capSilenceMs"/>. Designed to remove multi-second blank stretches that the
    /// Qwen 0.68B split model occasionally generates.
    /// </summary>
    private static byte[] TrimLongInteriorSilences(byte[] pcmData, WaveFormat format, int maxSilenceMs = 1200, int capSilenceMs = 450)
    {
        if (pcmData.Length == 0 ||
            format.Encoding != WaveFormatEncoding.Pcm ||
            format.BitsPerSample != 16 ||
            format.BlockAlign <= 0 ||
            format.AverageBytesPerSecond <= 0)
        {
            return pcmData;
        }

        var blockAlign = format.BlockAlign;
        var usableLength = pcmData.Length - (pcmData.Length % blockAlign);
        if (usableLength <= blockAlign)
        {
            return pcmData;
        }

        var bytesPerMs = format.AverageBytesPerSecond / 1000.0;
        var maxSilenceBytes = AlignDown((int)Math.Round(maxSilenceMs * bytesPerMs), blockAlign);
        var capBytes = AlignDown((int)Math.Round(capSilenceMs * bytesPerMs), blockAlign);
        if (maxSilenceBytes <= 0 || capBytes <= 0)
        {
            return pcmData;
        }

        // Use 100 ms windows to classify each segment as silent or active.
        const int silenceThreshold = 200;
        var windowBytes = AlignDown(Math.Max(blockAlign, format.AverageBytesPerSecond / 10), blockAlign);

        // Build run-length list: (isSilent, startByte, byteCount)
        var runs = new List<(bool Silent, int Start, int Length)>();
        var curSilent = false;
        var curStart = 0;
        var curLength = 0;

        for (var offset = 0; offset < usableLength; offset += windowBytes)
        {
            var len = Math.Min(windowBytes, usableLength - offset);
            if (len < blockAlign)
            {
                break;
            }

            long sumAbs = 0;
            var n = 0;
            for (var i = offset; i + 1 < offset + len; i += 2)
            {
                sumAbs += Math.Abs((int)(short)(pcmData[i] | (pcmData[i + 1] << 8)));
                n++;
            }

            var silent = n > 0 && (double)sumAbs / n < silenceThreshold;
            if (runs.Count == 0 || silent != curSilent)
            {
                if (runs.Count > 0)
                {
                    runs.Add((curSilent, curStart, curLength));
                }

                curSilent = silent;
                curStart = offset;
                curLength = len;
            }
            else
            {
                curLength += len;
            }
        }

        if (curLength > 0)
        {
            runs.Add((curSilent, curStart, curLength));
        }

        // If no interior silent run exceeds the threshold, return unchanged.
        if (!runs.Any(r => r.Silent && r.Length > maxSilenceBytes))
        {
            return pcmData;
        }

        // Build output, capping oversized silent runs.
        var outputParts = new List<(int Start, int Count)>(runs.Count);
        foreach (var (silent, start, length) in runs)
        {
            outputParts.Add(silent && length > maxSilenceBytes
                ? (start, Math.Max(capBytes, Math.Min(capBytes, length)))
                : (start, length));
        }

        var totalOut = outputParts.Sum(p => p.Count);
        if (totalOut >= pcmData.Length)
        {
            return pcmData;
        }

        var output = new byte[totalOut];
        var pos = 0;
        foreach (var (start, count) in outputParts)
        {
            var take = Math.Min(count, pcmData.Length - start);
            if (take > 0)
            {
                Array.Copy(pcmData, start, output, pos, take);
                pos += take;
            }
        }

        return output;
    }

    private static void ApplyBoundaryFadesIfSupported(byte[] pcmData, WaveFormat format, int fadeInMs, int fadeOutMs)
    {
        if ((fadeInMs <= 0 && fadeOutMs <= 0) || pcmData.Length == 0)
        {
            return;
        }

        if (format.Encoding != WaveFormatEncoding.Pcm || format.BitsPerSample != 16 || format.BlockAlign <= 0)
        {
            return;
        }

        var bytesPerMillisecond = format.AverageBytesPerSecond / 1000.0;
        var maxUsable = pcmData.Length - (pcmData.Length % format.BlockAlign);
        if (maxUsable <= 0)
        {
            return;
        }

        if (fadeInMs > 0)
        {
            var fadeBytes = (int)Math.Round(bytesPerMillisecond * fadeInMs);
            fadeBytes -= fadeBytes % format.BlockAlign;
            fadeBytes = Math.Min(fadeBytes, maxUsable);
            ApplyPcm16FadeIn(pcmData, fadeBytes, format.BlockAlign);
        }

        if (fadeOutMs > 0)
        {
            var fadeBytes = (int)Math.Round(bytesPerMillisecond * fadeOutMs);
            fadeBytes -= fadeBytes % format.BlockAlign;
            fadeBytes = Math.Min(fadeBytes, maxUsable);
            ApplyPcm16FadeOut(pcmData, fadeBytes, format.BlockAlign);
        }
    }

    private static void ApplyPcm16FadeIn(byte[] data, int fadeBytes, int blockAlign)
    {
        if (fadeBytes <= 0 || blockAlign <= 0)
        {
            return;
        }

        var steps = Math.Max(1, fadeBytes / blockAlign);
        for (var frame = 0; frame < steps; frame++)
        {
            var gain = (float)(frame + 1) / steps;
            var frameOffset = frame * blockAlign;
            for (var o = frameOffset; o + 1 < frameOffset + blockAlign && o + 1 < data.Length; o += 2)
            {
                short sample = (short)(data[o] | (data[o + 1] << 8));
                var scaled = (int)Math.Round(sample * gain);
                scaled = Math.Clamp(scaled, short.MinValue, short.MaxValue);
                data[o] = (byte)(scaled & 0xFF);
                data[o + 1] = (byte)((scaled >> 8) & 0xFF);
            }
        }
    }

    private static void ApplyPcm16FadeOut(byte[] data, int fadeBytes, int blockAlign)
    {
        if (fadeBytes <= 0 || blockAlign <= 0 || data.Length < blockAlign)
        {
            return;
        }

        var usableLength = data.Length - (data.Length % blockAlign);
        var start = Math.Max(0, usableLength - fadeBytes);
        var steps = Math.Max(1, (usableLength - start) / blockAlign);
        for (var frame = 0; frame < steps; frame++)
        {
            var gain = (float)(steps - frame) / steps;
            var frameOffset = start + (frame * blockAlign);
            for (var o = frameOffset; o + 1 < frameOffset + blockAlign && o + 1 < data.Length; o += 2)
            {
                short sample = (short)(data[o] | (data[o + 1] << 8));
                var scaled = (int)Math.Round(sample * gain);
                scaled = Math.Clamp(scaled, short.MinValue, short.MaxValue);
                data[o] = (byte)(scaled & 0xFF);
                data[o + 1] = (byte)((scaled >> 8) & 0xFF);
            }
        }
    }

    private static bool WaveFormatEquals(WaveFormat a, WaveFormat b)
    {
        return a.Encoding == b.Encoding
            && a.SampleRate == b.SampleRate
            && a.BitsPerSample == b.BitsPerSample
            && a.Channels == b.Channels
            && a.BlockAlign == b.BlockAlign;
    }

    private static void WriteSilence(WaveFileWriter writer, int milliseconds)
    {
        if (milliseconds <= 0)
        {
            return;
        }

        var bytes = (int)((long)writer.WaveFormat.AverageBytesPerSecond * milliseconds / 1000);
        var align = Math.Max(1, writer.WaveFormat.BlockAlign);
        bytes -= bytes % align;
        if (bytes <= 0)
        {
            return;
        }

        var silence = new byte[Math.Min(8192, bytes)];
        var remaining = bytes;
        while (remaining > 0)
        {
            var write = Math.Min(silence.Length, remaining);
            writer.Write(silence, 0, write);
            remaining -= write;
        }
    }

    private async Task FinalizeLocalOutputAsync(string sourceWav, string outputPath)
    {
        var ext = Path.GetExtension(outputPath).Trim().TrimStart('.').ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(ext) || ext == "wav")
        {
            File.Copy(sourceWav, outputPath, overwrite: true);
            return;
        }

        if (ext != "mp3")
        {
            throw new InvalidOperationException($"Unsupported local output format: {ext}. Use wav or mp3.");
        }

        var ffmpeg = ResolveFfmpegExecutable();
        if (string.IsNullOrWhiteSpace(ffmpeg))
        {
            throw new InvalidOperationException("ffmpeg not found for mp3 export. Expected tools/ffmpeg/ffmpeg.exe.");
        }

        var psi = new ProcessStartInfo
        {
            FileName = ffmpeg,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("-y");
        psi.ArgumentList.Add("-i");
        psi.ArgumentList.Add(sourceWav);
        psi.ArgumentList.Add("-vn");
        psi.ArgumentList.Add("-c:a");
        psi.ArgumentList.Add("libmp3lame");
        psi.ArgumentList.Add("-q:a");
        psi.ArgumentList.Add("2");
        psi.ArgumentList.Add(outputPath);

        using var proc = Process.Start(psi);
        if (proc is null)
        {
            throw new InvalidOperationException("Failed to start ffmpeg process.");
        }

        var stderr = await proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync();
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException($"ffmpeg failed: {stderr}");
        }
    }

    private static string? ResolveFfmpegExecutable()
    {
        var localTool = Path.Combine(RuntimePaths.AppRoot, "tools", "ffmpeg", "ffmpeg.exe");
        if (File.Exists(localTool))
        {
            return localTool;
        }

        const string fallback = "ffmpeg";
        return fallback;
    }

    private sealed class WriterHolder : IDisposable
    {
        public WaveFileWriter? Writer { get; set; }

        public void Dispose()
        {
            Writer?.Dispose();
        }
    }

    private void QueueGrid_OnMouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (QueueGrid.SelectedItem is not QueueRow row || string.IsNullOrWhiteSpace(row.OutputPath))
        {
            return;
        }
        if (!File.Exists(row.OutputPath))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = row.OutputPath,
            UseShellExecute = true
        });
    }

    private async void RetryQueueRowButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: QueueRow row })
        {
            return;
        }

        if (row.IsRunning)
        {
            _pauseGenerationRequested = !_pauseGenerationRequested;
            row.IsPaused = _pauseGenerationRequested;
            row.Status = _pauseGenerationRequested ? "Paused" : "Generating...";
            row.ChunkInfo = _pauseGenerationRequested ? "Paused" : row.ChunkInfo;
            row.Eta = _pauseGenerationRequested ? "Paused" : "Resuming...";
            QueueGrid.Items.Refresh();
            return;
        }

        if (row.CanRemoveQueued)
        {
            _queueRows.Remove(row);
            QueueGrid.Items.Refresh();
            return;
        }

        if (!row.CanRetry)
        {
            return;
        }

        if (_isGenerating)
        {
            return;
        }
        if (string.IsNullOrWhiteSpace(row.SourcePath) || !File.Exists(row.SourcePath))
        {
            MessageBox.Show(this, "Original source file not found for retry.", "Retry", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        CollectProjectFromUi();
        var isApi = string.Equals(_config.BackendMode, "api", StringComparison.OrdinalIgnoreCase);
        var effectiveLocalSettings = isApi ? new SynthesisSettings() : GetEffectiveLocalSettings();
        if (!isApi)
        {
            effectiveLocalSettings.StylePresetKey = GetSelectedStylePresetKey();
            ApplySelectedStyleRuntimeOverrides(effectiveLocalSettings);
        }
        var mixedPrepareMode = !isApi && IsMixedVoiceToken(_project.VoicePath);
        if (mixedPrepareMode && !HasUsablePreparedScript(row.SourcePath))
        {
            MessageBox.Show(this,
                "Mixed mode requires prepared parts for this chapter. Click Prepare first.",
                "Prepare Required",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }
        var requiresVoiceFile = !isApi && !mixedPrepareMode && !IsActiveKittenModel() && !IsLocalQwenCustomVoiceMode();
        if (!isApi && (string.IsNullOrWhiteSpace(_project.VoicePath) || (requiresVoiceFile && !File.Exists(_project.VoicePath))))
        {
            MessageBox.Show(this, "Select a valid voice.", "Retry", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (requiresVoiceFile && !TryEnsureProjectVoiceIsWav(out var retryPrepareError))
        {
            MessageBox.Show(this, retryPrepareError, "Retry", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (requiresVoiceFile && !ValidateLocalVoiceForGeneration(_project.VoicePath, out var retryVoiceError))
        {
            MessageBox.Show(this, retryVoiceError, "Retry", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Directory.CreateDirectory(_project.OutputDir);
        var ext = isApi ? "mp3" : NormalizeLocalOutputExtension(effectiveLocalSettings.OutputFormat);
        var outputPath = Path.Combine(_project.OutputDir, $"{Path.GetFileNameWithoutExtension(row.SourcePath)}.{ext}");

        ITtsBackend? backend = null;
        try
        {
            _isGenerating = true;
            GenerateButton.IsEnabled = false;
            _pauseGenerationRequested = false;
            _generationCts?.Dispose();
            _generationCts = new CancellationTokenSource();
            var generationCt = _generationCts.Token;

        row.Status = "Generating narration...";
            row.ProgressLabel = "25%";
            row.ProgressValue = 25;
            row.ChunkInfo = "--";
            row.Eta = "Starting...";
            row.IsRunning = true;
            row.IsPaused = false;
            QueueGrid.Items.Refresh();
            var rowStart = DateTime.UtcNow;

            backend = CreateBackend();
            TraceGenerate($"Local device preference: {NormalizePreferDevice(_config.PreferDevice)}");
            var text = await File.ReadAllTextAsync(row.SourcePath);
            var preparedScript = TryGetPreparedScript(row.SourcePath);
            var usePrepared = !isApi && IsMixedVoiceToken(_project.VoicePath) && preparedScript is { Parts.Count: > 0 };
            IReadOnlyList<SubtitleChunkTiming>? subtitleChunkTimings = null;
            EnhanceResult? enhanceResult = null;
            if (isApi)
            {
                await backend.SynthesizeAsync(CreateLocalTtsRequest(
                    text,
                    _project.VoicePath,
                    outputPath,
                    (float)SpeedSlider.Value,
                    effectiveLocalSettings), generationCt);
            }
            else
            {
                if (usePrepared && preparedScript is not null)
                {
                    subtitleChunkTimings = await GeneratePreparedScriptLocalAsync(
                        backend,
                        preparedScript,
                        outputPath,
                        (float)SpeedSlider.Value,
                        effectiveLocalSettings,
                        generationCt,
                        (doneParts, totalParts) =>
                        {
                            var total = Math.Max(totalParts, 1);
                            var boundedDone = Math.Clamp(doneParts, 0, total);
                            var progress = boundedDone == 0
                                ? 10
                                : 10 + (int)Math.Round((boundedDone * 85.0) / total);
                            progress = Math.Clamp(progress, 10, 99);
                            row.ProgressLabel = $"{progress}%";
                            row.ProgressValue = progress;
                            if (boundedDone == 0)
                            {
                                row.ChunkInfo = $"Part 0/{total}";
                                row.Eta = "Starting...";
                            }
                            else
                            {
                                var elapsed = DateTime.UtcNow - rowStart;
                                var remainingSeconds = (elapsed.TotalSeconds / boundedDone) * (total - boundedDone);
                                row.ChunkInfo = $"Part {boundedDone}/{total}";
                                row.Eta = FormatDuration(TimeSpan.FromSeconds(Math.Max(0, remainingSeconds)));
                            }
                            QueueGrid.Items.Refresh();
                        });
                }
                else
                {
                    subtitleChunkTimings = await GenerateLocalWithChunkingAsync(
                        backend,
                        text,
                        _project.VoicePath,
                        outputPath,
                        (float)SpeedSlider.Value,
                        effectiveLocalSettings,
                        generationCt,
                        (doneChunks, totalChunks) =>
                        {
                            var total = Math.Max(totalChunks, 1);
                            var boundedDone = Math.Clamp(doneChunks, 0, total);
                            var progress = boundedDone == 0
                                ? 10
                                : 10 + (int)Math.Round((boundedDone * 85.0) / total);
                            progress = Math.Clamp(progress, 10, 99);
                            row.ProgressLabel = $"{progress}%";
                            row.ProgressValue = progress;
                            if (boundedDone == 0)
                            {
                                row.ChunkInfo = $"Chunk 0/{total}";
                                row.Eta = "Starting...";
                            }
                            else
                            {
                                var elapsed = DateTime.UtcNow - rowStart;
                                var remainingSeconds = (elapsed.TotalSeconds / boundedDone) * (total - boundedDone);
                                row.ChunkInfo = $"Chunk {boundedDone}/{total}";
                                row.Eta = FormatDuration(TimeSpan.FromSeconds(Math.Max(0, remainingSeconds)));
                            }
                            QueueGrid.Items.Refresh();
                    });
                }
            }

            if (IsEnhanceEnabledForRun())
            {
                var enhanceProgress = new Progress<AudioEnhanceProgress>(p =>
                {
                    var total = Math.Max(p.Total, 1);
                    var done = Math.Clamp(p.Completed, 0, total);
                    var progress = 90 + (int)Math.Round((done * 9.0) / total);
                    row.Status = p.Stage;
                    row.ProgressLabel = $"{Math.Clamp(progress, 90, 99)}%";
                    row.ProgressValue = Math.Clamp(progress, 90, 99);
                    row.ChunkInfo = total > 1 ? $"{done}/{total}" : p.Stage;
                    var elapsed = DateTime.UtcNow - rowStart;
                    row.Eta = FormatDuration(elapsed);
                    QueueGrid.Items.Refresh();
                });

                enhanceResult = await RunAudioEnhanceForChapterAsync(
                    text,
                    outputPath,
                    subtitleChunkTimings,
                    preparedScript,
                    generationCt,
                    enhanceProgress);

                if (!enhanceResult.Success)
                {
                    TraceGenerate($"Enhance warning for retry {row.FileName}: {enhanceResult.Message}");
                }
            }
            await GenerateSubtitlesIfEnabledAsync(text, outputPath, subtitleChunkTimings);

            var noEnhanceCues = enhanceResult is { Success: true } &&
                                !string.IsNullOrWhiteSpace(enhanceResult.Message) &&
                                enhanceResult.Message.Contains("narration-only", StringComparison.OrdinalIgnoreCase);
            row.Status = noEnhanceCues ? "Done (no cues)" : "Done";
            row.ProgressLabel = "100%";
            row.ProgressValue = 100;
            row.ChunkInfo = enhanceResult is { Success: false }
                ? "Done (enhance warning)"
                : noEnhanceCues
                    ? "Narration only"
                    : "Done";
            row.Eta = FormatDuration(DateTime.UtcNow - rowStart);
            row.OutputPath = outputPath;
            row.IsRunning = false;
            row.IsPaused = false;
            AddHistoryEntryForCompletedRow(row, isApi, effectiveLocalSettings);
            TryAutoClearCompletedQueueRow(row);
            if (backend is ChatterboxOnnxBackend chatterboxBackend)
                TraceGenerate($"ONNX provider used: {chatterboxBackend.ActiveExecutionProvider}");
            else if (backend is Qwen3OnnxSplitBackend qwenSplitBackend)
                TraceGenerate($"ONNX provider used: {qwenSplitBackend.ActiveExecutionProvider}");
        }
        catch (OperationCanceledException)
        {
            row.Status = "Stopped";
            row.ChunkInfo = "Stopped";
            row.Eta = "Stopped";
            row.IsRunning = false;
            row.IsPaused = false;
        }
        catch (Exception ex)
        {
            row.Status = "Failed";
            row.ProgressLabel = "0%";
            row.ProgressValue = 0;
            row.ChunkInfo = "Failed";
            row.Eta = "Failed";
            row.IsRunning = false;
            row.IsPaused = false;
            MessageBox.Show(this, NormalizeGenerationError(ex.Message), "Retry Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (backend is IDisposable backendDisposable)
            {
                backendDisposable.Dispose();
            }
            QueueGrid.Items.Refresh();
            GenerateButton.IsEnabled = true;
            _isGenerating = false;
            _pauseGenerationRequested = false;
            _generationCts?.Dispose();
            _generationCts = null;
            _ = AutoSaveCurrentProjectAsync();
        }
    }

    private void PlayQueueRowButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: QueueRow row })
        {
            return;
        }

        if (row.IsRunning)
        {
            row.Status = "Stopping...";
            row.ChunkInfo = "Stopping...";
            row.Eta = "Stopping...";
            row.IsPaused = false;
            _pauseGenerationRequested = false;
            _generationCts?.Cancel();
            QueueGrid.Items.Refresh();
            return;
        }

        if (!row.CanPlay)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(row.OutputPath) || !File.Exists(row.OutputPath))
        {
            MessageBox.Show(this, "No generated audio found for this row yet.", "Play Output", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = row.OutputPath,
            UseShellExecute = true
        });
    }

    public sealed class QueueRow : INotifyPropertyChanged
    {
        private string _fileName = string.Empty;
        private string _sourcePath = string.Empty;
        private string _status = "Queued";
        private string _progressLabel = "0%";
        private double _progressValue;
        private string _chunkInfo = "Queued";
        private string _eta = "Queued";
        private string _outputPath = string.Empty;
        private bool _isRunning;
        private bool _isPaused;

        public string FileName
        {
            get => _fileName;
            set => SetField(ref _fileName, value);
        }

        public string SourcePath
        {
            get => _sourcePath;
            set => SetField(ref _sourcePath, value);
        }

        public string Status
        {
            get => _status;
            set
            {
                if (SetField(ref _status, value))
                {
                    NotifyActionStateChanged();
                }
            }
        }

        public string ProgressLabel
        {
            get => _progressLabel;
            set => SetField(ref _progressLabel, value);
        }

        public double ProgressValue
        {
            get => _progressValue;
            set => SetField(ref _progressValue, value);
        }

        public string ChunkInfo
        {
            get => _chunkInfo;
            set => SetField(ref _chunkInfo, value);
        }

        public string Eta
        {
            get => _eta;
            set => SetField(ref _eta, value);
        }

        public string OutputPath
        {
            get => _outputPath;
            set
            {
                if (SetField(ref _outputPath, value))
                {
                    NotifyActionStateChanged();
                }
            }
        }

        public bool IsRunning
        {
            get => _isRunning;
            set
            {
                if (SetField(ref _isRunning, value))
                {
                    NotifyActionStateChanged();
                }
            }
        }

        public bool IsPaused
        {
            get => _isPaused;
            set
            {
                if (SetField(ref _isPaused, value))
                {
                    NotifyActionStateChanged();
                }
            }
        }

        public bool CanRetry
        {
            get
            {
                if (IsRunning)
                {
                    return false;
                }

                return string.Equals(Status, "Done", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(Status, "Failed", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(Status, "Stopped", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(Status, "Missing file", StringComparison.OrdinalIgnoreCase);
            }
        }

        public bool CanRemoveQueued =>
            !IsRunning &&
            string.Equals(Status, "Queued", StringComparison.OrdinalIgnoreCase);

        public bool CanPlay => !IsRunning && !string.IsNullOrWhiteSpace(OutputPath) && File.Exists(OutputPath);

        public string PrimaryActionLabel => IsRunning
            ? (IsPaused ? "Resume" : "Pause")
            : (CanRemoveQueued ? "Remove" : (CanRetry ? "Retry" : string.Empty));

        public string SecondaryActionLabel => IsRunning ? "Stop" : (CanPlay ? "Play" : string.Empty);

        public string PrimaryActionGlyph => IsRunning
            ? (IsPaused ? "\uE768" : "\uE769")
            : (CanRemoveQueued ? "\uE74D" : (CanRetry ? "\uE72C" : string.Empty));

        public string SecondaryActionGlyph => IsRunning ? "\uE71A" : (CanPlay ? "\uE768" : string.Empty);

        public string PrimaryActionToolTip => IsRunning
            ? (IsPaused ? "Resume" : "Pause")
            : (CanRemoveQueued ? "Remove from queue" : (CanRetry ? "Retry" : string.Empty));

        public string SecondaryActionToolTip => IsRunning ? "Stop" : (CanPlay ? "Play output" : string.Empty);

        public event PropertyChangedEventHandler? PropertyChanged;

        private void NotifyActionStateChanged()
        {
            OnPropertyChanged(nameof(CanRetry));
            OnPropertyChanged(nameof(CanRemoveQueued));
            OnPropertyChanged(nameof(CanPlay));
            OnPropertyChanged(nameof(PrimaryActionLabel));
            OnPropertyChanged(nameof(SecondaryActionLabel));
            OnPropertyChanged(nameof(PrimaryActionGlyph));
            OnPropertyChanged(nameof(SecondaryActionGlyph));
            OnPropertyChanged(nameof(PrimaryActionToolTip));
            OnPropertyChanged(nameof(SecondaryActionToolTip));
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }

    public sealed class HistoryRow
    {
        public string FileName { get; set; } = string.Empty;
        public string SourcePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public string ModelLabel { get; set; } = string.Empty;
        public string DeviceLabel { get; set; } = string.Empty;
        public string VoiceLabel { get; set; } = string.Empty;
        public string CompletedAt { get; set; } = string.Empty;
        public string DurationLabel { get; set; } = string.Empty;
    }

    private void CreateVoiceClone_OnClick(object sender, RoutedEventArgs e)
    {
        var source = _lastRecordedPath ?? string.Empty;
        if (string.IsNullOrWhiteSpace(source) || !File.Exists(source))
        {
            MessageBox.Show(this, "Record first to create a clone. To add existing WAV/MP3, use Import Voice in Available Clone Voices.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!source.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "Voice cloning currently requires WAV input (PCM 16-bit mono/stereo).", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var rawName = CloneNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(rawName))
        {
            MessageBox.Show(this, "Enter a clone name.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var safeName = Regex.Replace(rawName, @"[^a-zA-Z0-9_\- ]", "_").Trim();
        if (string.IsNullOrWhiteSpace(safeName))
        {
            safeName = "voice_clone";
        }

        if (!TryInspectWav(source, out var wavInfo, out var wavError))
        {
            MessageBox.Show(this, wavError, "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (wavInfo.DurationSeconds < 10.0 || wavInfo.DurationSeconds > 60.0)
        {
            MessageBox.Show(this, $"Source duration is {wavInfo.DurationSeconds:0.0}s. Use 10s-60s for stable voice cloning.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (wavInfo.Rms < 0.010)
        {
            MessageBox.Show(this, "Source audio is too quiet. Please use a clearer/louder recording.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (wavInfo.Rms > 0.45)
        {
            MessageBox.Show(this, "Source audio appears clipped/distorted. Please use a cleaner recording.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
        Directory.CreateDirectory(voicesDir);

        var targetPath = Path.Combine(voicesDir, $"{safeName}.wav");
        if (File.Exists(targetPath))
        {
            var overwriteChoice = MessageBox.Show(
                this,
                $"A clone named '{safeName}' already exists.\n\nYes = overwrite\nNo = keep existing and create a numbered copy\nCancel = abort",
                "Voice Clone",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (overwriteChoice == MessageBoxResult.Cancel)
            {
                return;
            }

            if (overwriteChoice == MessageBoxResult.No)
            {
                var suffix = 2;
                string candidate;
                do
                {
                    candidate = Path.Combine(voicesDir, $"{safeName}_{suffix}.wav");
                    suffix++;
                } while (File.Exists(candidate));
                targetPath = candidate;
            }
        }

        File.Copy(source, targetPath, overwrite: true);

        CloneStatusTextBlock.Text = $"Clone created: {Path.GetFileName(targetPath)}";
        LoadVoiceList();
        RefreshCloneVoicesList();
        SelectVoiceByPath(targetPath);
    }

    private async void GenerateQwenPromptClone_OnClick(object sender, RoutedEventArgs e)
    {
        if (_voiceDesignCts is not null)
        {
            MessageBox.Show(this, "Voice design generation is already running.", "Qwen Voice Design", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var rawName = (QwenVoiceDesignNameTextBox?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(rawName))
        {
            rawName = CloneNameTextBox.Text.Trim();
        }
        if (string.IsNullOrWhiteSpace(rawName))
        {
            MessageBox.Show(this, "Enter a voice name.", "Qwen Voice Design", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var designText = (QwenVoiceDesignTextBox.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(designText))
        {
            MessageBox.Show(this, "Enter the text that should be spoken by the generated voice.", "Qwen Voice Design", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var prompt = (QwenVoicePromptTextBox.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(prompt))
        {
            MessageBox.Show(this, "Enter a Qwen style prompt first.", "Qwen Voice Design", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var language = ResolveVoiceDesignLanguageId();
        CaptureVoiceDesignDraftFromUi(_project);

        var safeName = Regex.Replace(rawName, @"[^a-zA-Z0-9_\- ]", "_").Trim();
        if (string.IsNullOrWhiteSpace(safeName))
        {
            safeName = "voice_design";
        }

        var button = sender as Button;
        var oldContent = button?.Content;
        if (button is not null)
        {
            button.IsEnabled = false;
            button.Content = "Generating...";
        }
        if (CancelVoiceDesignButton is not null)
        {
            CancelVoiceDesignButton.Visibility = Visibility.Visible;
            CancelVoiceDesignButton.IsEnabled = true;
        }
        StartVoiceDesignProgress("Preparing model...");
        _voiceDesignCts = new CancellationTokenSource();
        var ct = _voiceDesignCts.Token;

        QwenPythonBackend? backend = null;
        ApiBackend? apiBackend = null;
        try
        {
            var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
            Directory.CreateDirectory(voicesDir);
            var baseTarget = Path.Combine(voicesDir, $"{safeName}.wav");
            var targetPath = baseTarget;
            if (File.Exists(targetPath))
            {
                var suffix = 2;
                while (File.Exists(targetPath))
                {
                    targetPath = Path.Combine(voicesDir, $"{safeName}_{suffix}.wav");
                    suffix++;
                }
            }

            var localSettings = GetEffectiveLocalSettings();
            ApplyQwenStableAudiobookPreset(localSettings);
            localSettings.StylePresetKey = GetSelectedStylePresetKey();
            ApplySelectedStyleRuntimeOverrides(localSettings);

            var backendMode = (_config.BackendMode ?? "auto").Trim().ToLowerInvariant();
            var apiProvider = (_config.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
            var useAlibabaApi = backendMode == "api" && apiProvider == "alibaba";
            var effectiveDesignText = designText;

            if (useAlibabaApi)
            {
                effectiveDesignText = PrepareAlibabaVoiceDesignPreviewText(designText);
                apiBackend = CreateBackend() as ApiBackend;
                if (apiBackend is null)
                {
                    throw new InvalidOperationException("Failed to initialize Alibaba API backend.");
                }

                StartVoiceDesignProgress("Preparing model...");
                CloneStatusTextBlock.Text = "Generating Qwen voice design (Alibaba API)...";
                var workerProgress = new Progress<QwenWorkerProgress>(ApplyVoiceDesignWorkerProgress);
                await apiBackend.GenerateAlibabaVoiceDesignAsync(
                    voiceName: safeName,
                    previewText: effectiveDesignText,
                    stylePrompt: prompt,
                    languageType: language,
                    outputPath: targetPath,
                    progress: workerProgress,
                    ct: ct);
            }
            else
            {
                CloneStatusTextBlock.Text = "Preparing Qwen model...";
                var qwenCfg = BuildQwenVoiceDesignConfig();
                await EnsureQwenVoiceDesignModelAsync(qwenCfg, ct);
                backend = CreateQwenVoiceDesignBackend(qwenCfg);

                StartVoiceDesignProgress("Generating voice audio...");
                CloneStatusTextBlock.Text = "Generating Qwen voice design...";
                var workerProgress = new Progress<QwenWorkerProgress>(ApplyVoiceDesignWorkerProgress);
                await backend.GenerateVoiceDesignAsync(
                    text: designText,
                    instruct: prompt,
                    language: language,
                    outputPath: targetPath,
                    doSample: localSettings.QwenDoSample,
                    temperature: (float)localSettings.QwenTemperature,
                    topK: localSettings.QwenTopK,
                    topP: (float)localSettings.QwenTopP,
                    repetitionPenalty: (float)localSettings.QwenRepetitionPenalty,
                    progress: workerProgress,
                    ct: ct);
            }
            await GenerateSubtitlesIfEnabledAsync(effectiveDesignText, targetPath, null);

            CloneStatusTextBlock.Text = $"Qwen voice design created: {Path.GetFileName(targetPath)}";
            LoadVoiceList();
            RefreshCloneVoicesList();
            SelectVoiceByPath(targetPath);
            var generatedName = Path.GetFileNameWithoutExtension(targetPath);
            CloneNameTextBox.Text = generatedName;
            if (QwenVoiceDesignNameTextBox is not null)
            {
                QwenVoiceDesignNameTextBox.Text = generatedName;
            }
        }
        catch (OperationCanceledException)
        {
            CloneStatusTextBlock.Text = "Qwen voice design canceled.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                "Qwen voice design failed.\n\nThis feature requires the Qwen VoiceDesign model and the bundled Python worker runtime.\n\n" + ex.Message,
                "Qwen Voice Design",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            CloneStatusTextBlock.Text = $"Qwen voice design failed: {ex.Message}";
        }
        finally
        {
            _voiceDesignCts?.Dispose();
            _voiceDesignCts = null;
            if (backend is IDisposable disposable)
            {
                disposable.Dispose();
            }
            if (button is not null)
            {
                button.IsEnabled = true;
                button.Content = oldContent ?? "Generate Voice Design";
            }
            if (CancelVoiceDesignButton is not null)
            {
                CancelVoiceDesignButton.IsEnabled = false;
                CancelVoiceDesignButton.Visibility = Visibility.Collapsed;
            }
            StopVoiceDesignProgress();
            RefreshVoiceDesignReadiness();
        }
    }

    private void CancelVoiceDesign_OnClick(object sender, RoutedEventArgs e)
    {
        _voiceDesignCts?.Cancel();
    }

    private void StartVoiceDesignProgress(string stage)
    {
        if (VoiceDesignProgressPanel is null)
        {
            return;
        }

        VoiceDesignProgressPanel.Visibility = Visibility.Visible;
        VoiceDesignProgressText.Text = string.IsNullOrWhiteSpace(stage) ? "Generating..." : stage.Trim();
        VoiceDesignElapsedText.Text = "00:00";
        _voiceDesignCurrentStep = 0;
        _voiceDesignTotalSteps = 0;
        VoiceDesignProgressBar.IsIndeterminate = true;
        VoiceDesignProgressBar.Value = 0;
        _voiceDesignStartedAtUtc = DateTime.UtcNow;

        if (_voiceDesignTimer is null)
        {
            _voiceDesignTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _voiceDesignTimer.Tick += VoiceDesignTimerOnTick;
        }
        _voiceDesignTimer.Start();
    }

    private void StopVoiceDesignProgress()
    {
        if (_voiceDesignTimer is not null)
        {
            _voiceDesignTimer.Stop();
        }

        if (VoiceDesignProgressPanel is not null)
        {
            VoiceDesignProgressPanel.Visibility = Visibility.Collapsed;
        }
        _voiceDesignCurrentStep = 0;
        _voiceDesignTotalSteps = 0;
        if (VoiceDesignProgressBar is not null)
        {
            VoiceDesignProgressBar.IsIndeterminate = true;
            VoiceDesignProgressBar.Value = 0;
        }
        if (VoiceDesignElapsedText is not null)
        {
            VoiceDesignElapsedText.Text = "00:00";
        }
    }

    private void ApplyVoiceDesignWorkerProgress(QwenWorkerProgress update)
    {
        if (update is null)
        {
            return;
        }

        if (VoiceDesignProgressPanel is not null && VoiceDesignProgressPanel.Visibility != Visibility.Visible)
        {
            StartVoiceDesignProgress(update.Message);
        }
        else if (!string.IsNullOrWhiteSpace(update.Message))
        {
            VoiceDesignProgressText.Text = update.Message.Trim();
        }

        if (update.Step.HasValue && update.Step.Value > 0)
        {
            _voiceDesignCurrentStep = update.Step.Value;
        }
        if (update.TotalSteps.HasValue && update.TotalSteps.Value > 0)
        {
            _voiceDesignTotalSteps = update.TotalSteps.Value;
        }

        if (_voiceDesignCurrentStep > 0 && _voiceDesignTotalSteps > 0 && VoiceDesignProgressBar is not null)
        {
            var completed = Math.Max(0, _voiceDesignCurrentStep - 1);
            var percent = (completed * 100.0) / _voiceDesignTotalSteps;
            VoiceDesignProgressBar.IsIndeterminate = false;
            VoiceDesignProgressBar.Value = Math.Clamp(percent, 0, 99);
        }
    }

    private void VoiceDesignTimerOnTick(object? sender, EventArgs e)
    {
        if (VoiceDesignElapsedText is null || VoiceDesignProgressPanel is null || VoiceDesignProgressPanel.Visibility != Visibility.Visible)
        {
            return;
        }

        var elapsed = DateTime.UtcNow - _voiceDesignStartedAtUtc;
        var elapsedText = $"{Math.Min(99, (int)elapsed.TotalMinutes):00}:{elapsed.Seconds:00}";

        if (_voiceDesignCurrentStep > 1 && _voiceDesignTotalSteps > _voiceDesignCurrentStep - 1)
        {
            var completed = _voiceDesignCurrentStep - 1;
            var remaining = _voiceDesignTotalSteps - completed;
            if (completed > 0 && remaining > 0)
            {
                var etaSeconds = (elapsed.TotalSeconds / completed) * remaining;
                var eta = TimeSpan.FromSeconds(Math.Max(0, etaSeconds));
                VoiceDesignElapsedText.Text = $"{elapsedText}  ETA ~{Math.Min(99, (int)eta.TotalMinutes):00}:{eta.Seconds:00}";
                return;
            }
        }

        VoiceDesignElapsedText.Text = elapsedText;
    }

    private AppConfig BuildQwenVoiceDesignConfig()
    {
        var cfg = new AppConfig
        {
            DefaultOutputDir = _config.DefaultOutputDir,
            ModelCacheDir = _config.ModelCacheDir,
            PreferDevice = NormalizePreferDevice(_config.PreferDevice),
            OfflineMode = false,
            BackendMode = "local",
            LocalModelPreset = "qwen3_tts",
            AutoDownloadModel = true,
            ModelProfiles = _config.ModelProfiles ?? new Dictionary<string, SynthesisSettings>()
        };

        cfg.ModelRepoId = "Qwen/Qwen3-TTS-12Hz-1.7B-VoiceDesign";
        cfg.AdditionalModelRepoId = string.Empty;

        return cfg;
    }

    private async Task EnsureQwenVoiceDesignModelAsync(AppConfig cfg, CancellationToken ct)
    {
        var progress = new Progress<string>(msg =>
        {
            Dispatcher.Invoke(() => CloneStatusTextBlock.Text = msg);
        });
        await _modelDownloader.DownloadAsync(cfg, progress, ct: ct);
    }

    private QwenPythonBackend CreateQwenVoiceDesignBackend(AppConfig cfg)
    {
        var localOptions = new LocalInferenceOptions
        {
            ModelCacheDir = cfg.ModelCacheDir,
            ModelRepoId = cfg.ModelRepoId,
            PreferDevice = NormalizePreferDevice(cfg.PreferDevice),
            MaxNewTokens = ResolveLocalMaxNewTokens(),
            ValidateOnnxRuntimeSessions = false
        };

        return new QwenPythonBackend(localOptions);
    }

    private void RefreshVoiceDesignReadiness()
    {
        if (VoiceDesignReadinessTextBlock is null)
        {
            return;
        }

        var backendMode = (_config.BackendMode ?? "auto").Trim().ToLowerInvariant();
        var apiProvider = (_config.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
        if (backendMode == "api" && apiProvider == "alibaba")
        {
            var hasApiKey = !string.IsNullOrWhiteSpace(ResolveApiKeyForProvider("alibaba"));
            var hasTargetModel = !string.IsNullOrWhiteSpace(_config.ApiVoiceDesignTargetModel);
            if (!hasApiKey)
            {
                VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
                VoiceDesignReadinessTextBlock.Text = "VoiceDesign API: Not Ready (missing API key)";
            }
            else if (!hasTargetModel)
            {
                VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
                VoiceDesignReadinessTextBlock.Text = "VoiceDesign API: Not Ready (missing target model)";
            }
            else
            {
                VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(79, 125, 99));
                VoiceDesignReadinessTextBlock.Text = "VoiceDesign API: Ready (key + model configured)";
            }
            return;
        }

        var cfg = BuildQwenVoiceDesignConfig();
        var missing = ValidateQwenVoiceDesignModelFiles(cfg.ModelCacheDir, cfg.ModelRepoId);
        var runtimeOk = IsQwenPythonRuntimeAvailable();
        if (string.IsNullOrWhiteSpace(missing))
        {
            if (!runtimeOk)
            {
                VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
                VoiceDesignReadinessTextBlock.Text = "VoiceDesign model: Not Ready (bundled Qwen runtime missing)";
            }
            else
            {
                VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(79, 125, 99));
                VoiceDesignReadinessTextBlock.Text = "VoiceDesign model: Ready";
            }
            return;
        }

        VoiceDesignReadinessTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
        VoiceDesignReadinessTextBlock.Text = $"VoiceDesign model: Not Ready ({missing})";
    }

    private string ResolveVoiceDesignLanguageId()
    {
        var raw = "auto";
        if (QwenVoiceDesignLanguageCombo?.SelectedItem is ComboBoxItem selected && selected.Tag is not null)
        {
            raw = selected.Tag.ToString() ?? "auto";
        }
        else if (!string.IsNullOrWhiteSpace(QwenVoiceDesignLanguageCombo?.Text))
        {
            raw = QwenVoiceDesignLanguageCombo.Text;
        }

        var key = (raw ?? "auto").Trim().ToLowerInvariant();
        return key switch
        {
            "auto" => "auto",
            "en" or "english" => "english",
            "zh" or "chinese" or "mandarin" => "chinese",
            "ja" or "japanese" => "japanese",
            "ko" or "korean" => "korean",
            "de" or "german" => "german",
            "fr" or "french" => "french",
            "ru" or "russian" => "russian",
            "pt" or "portuguese" => "portuguese",
            "es" or "spanish" => "spanish",
            "it" or "italian" => "italian",
            _ => "auto",
        };
    }

    private static string? ValidateQwenVoiceDesignModelFiles(string modelCacheDir, string repoId)
    {
        var cacheRoot = Path.Combine(ModelCachePath.ResolveAbsolute(modelCacheDir, RuntimePaths.AppRoot), "hf-cache");
        if (!TryResolveRepoFolder(cacheRoot, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "model files not downloaded";
        }

        var configPath = Path.Combine(repoFolder, "config.json");
        if (!File.Exists(configPath) || new FileInfo(configPath).Length < 16)
        {
            return "config.json missing";
        }

        var tokenizerConfigPath = Path.Combine(repoFolder, "tokenizer_config.json");
        if (!File.Exists(tokenizerConfigPath) || new FileInfo(tokenizerConfigPath).Length < 16)
        {
            return "tokenizer_config.json missing";
        }

        var hasWeights = Directory.GetFiles(repoFolder, "*.safetensors", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.bin", SearchOption.AllDirectories).Length > 0;
        if (!hasWeights)
        {
            return "model weights missing";
        }

        var hasTokenizerAssets =
            Directory.GetFiles(repoFolder, "tokenizer.json", SearchOption.AllDirectories).Length > 0 ||
            (Directory.GetFiles(repoFolder, "vocab.json", SearchOption.AllDirectories).Length > 0 &&
             Directory.GetFiles(repoFolder, "merges.txt", SearchOption.AllDirectories).Length > 0);
        if (!hasTokenizerAssets)
        {
            return "tokenizer files missing";
        }

        return null;
    }

    private static bool IsQwenPythonRuntimeAvailable()
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var candidates = new[]
        {
            Path.Combine(appRoot, "python_qwen", "python.exe"),
            Path.Combine(appRoot, "python_qwen", "Scripts", "python.exe"),
            Path.Combine(appRoot, "tools", "python_qwen", "python.exe"),
            Path.Combine(appRoot, "tools", "python_qwen", "Scripts", "python.exe")
        };
        return candidates.Any(File.Exists);
    }

    private string PrepareAlibabaVoiceDesignPreviewText(string raw)
    {
        var clean = Regex.Replace((raw ?? string.Empty).Trim(), @"\s+", " ");
        if (clean.Length <= 600)
        {
            return clean;
        }

        var trimmed = clean[..600];
        var cut = trimmed.LastIndexOfAny(new[] { '.', '!', '?', ';', ',', ' ' });
        if (cut > 260)
        {
            trimmed = trimmed[..(cut + 1)];
        }
        trimmed = trimmed.Trim();
        CloneStatusTextBlock.Text = "VoiceDesign note: preview text was trimmed to 600 chars for Alibaba API limit.";
        return trimmed;
    }

    private static bool TryResolveRepoFolder(string cacheRoot, string repoId, out string folder)
    {
        folder = string.Empty;
        if (string.IsNullOrWhiteSpace(cacheRoot) || string.IsNullOrWhiteSpace(repoId) || !Directory.Exists(cacheRoot))
        {
            return false;
        }

        var repoKey = repoId.Trim().Replace('\\', '/').Trim('/');
        var normalizedLegacy = "models--" + repoKey.Replace('/', '-');
        var normalizedHf = "models--" + repoKey.Replace("/", "--", StringComparison.Ordinal);
        var repoFolder = Directory.GetDirectories(cacheRoot, "models--*", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(path =>
            {
                var name = Path.GetFileName(path);
                return string.Equals(name, normalizedHf, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(name, normalizedLegacy, StringComparison.OrdinalIgnoreCase);
            });
        if (string.IsNullOrWhiteSpace(repoFolder))
        {
            return false;
        }

        var snapshotsDir = Path.Combine(repoFolder, "snapshots");
        if (Directory.Exists(snapshotsDir))
        {
            var snapshot = Directory.GetDirectories(snapshotsDir, "*", SearchOption.TopDirectoryOnly)
                .OrderByDescending(Directory.GetLastWriteTimeUtc)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(snapshot))
            {
                folder = snapshot;
                return true;
            }
        }

        folder = repoFolder;
        return true;
    }

    private void RefreshCloneList_OnClick(object sender, RoutedEventArgs e)
    {
        LoadVoiceList();
        RefreshCloneVoicesList();
        CloneStatusTextBlock.Text = "Clone list refreshed.";
    }

    private void ImportCloneVoice_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Audio Files (*.wav;*.mp3)|*.wav;*.mp3",
            InitialDirectory = Path.Combine(RuntimePaths.AppRoot, "voices")
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var sourcePath = dialog.FileName;
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            MessageBox.Show(this, "Selected file no longer exists.", "Import Voice", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var ext = Path.GetExtension(sourcePath);
        string importedPath;
        if (string.Equals(ext, ".mp3", StringComparison.OrdinalIgnoreCase))
        {
            if (!TryNormalizeLocalVoiceFile(sourcePath, out importedPath, out var error))
            {
                MessageBox.Show(this, error, "Import Voice", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }
        else if (string.Equals(ext, ".wav", StringComparison.OrdinalIgnoreCase))
        {
            var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
            Directory.CreateDirectory(voicesDir);

            var safeName = Regex.Replace(Path.GetFileNameWithoutExtension(sourcePath) ?? "imported_voice", @"[^a-zA-Z0-9_\- ]", "_").Trim();
            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = "imported_voice";
            }

            importedPath = Path.Combine(voicesDir, $"{safeName}.wav");
            if (!string.Equals(Path.GetFullPath(importedPath), Path.GetFullPath(sourcePath), StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(importedPath))
                {
                    var suffix = 2;
                    while (File.Exists(importedPath))
                    {
                        importedPath = Path.Combine(voicesDir, $"{safeName}_{suffix}.wav");
                        suffix++;
                    }
                }

                File.Copy(sourcePath, importedPath, overwrite: false);
            }
        }
        else
        {
            MessageBox.Show(this, "Only WAV and MP3 are supported for import.", "Import Voice", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        LoadVoiceList();
        RefreshCloneVoicesList();
        SelectCloneVoiceByPath(importedPath);
        SelectVoiceByPath(importedPath);
        CloneStatusTextBlock.Text = $"Imported voice: {Path.GetFileName(importedPath)}";
    }

    private void DeleteSelectedClone_OnClick(object sender, RoutedEventArgs e)
    {
        var voicePath = ExtractPathFromListItem(CloneVoicesList.SelectedItem);
        if (string.IsNullOrWhiteSpace(voicePath))
        {
            MessageBox.Show(this, "Select a clone voice from the list.", "Delete Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!File.Exists(voicePath))
        {
            MessageBox.Show(this, "Selected clone file no longer exists.", "Delete Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            RefreshCloneVoicesList();
            return;
        }

        var voicesDir = Path.GetFullPath(Path.Combine(RuntimePaths.AppRoot, "voices"));
        var fullVoicePath = Path.GetFullPath(voicePath);
        if (!fullVoicePath.StartsWith(voicesDir, StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "Delete is allowed only for files inside the voices folder.", "Delete Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            this,
            $"Delete clone voice '{Path.GetFileName(fullVoicePath)}'?\n\nThis also deletes matching .ref.txt if it exists.",
            "Delete Clone",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            File.Delete(fullVoicePath);
            var refTextPath = Path.ChangeExtension(fullVoicePath, ".ref.txt");
            if (!string.IsNullOrWhiteSpace(refTextPath) && File.Exists(refTextPath))
            {
                File.Delete(refTextPath);
            }

            if (string.Equals(_project.VoicePath, fullVoicePath, StringComparison.OrdinalIgnoreCase))
            {
                _project.VoicePath = string.Empty;
            }

            LoadVoiceList();
            RefreshCloneVoicesList();
            CloneStatusTextBlock.Text = $"Deleted clone: {Path.GetFileName(fullVoicePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Delete clone", ex.Message),
                "Delete Clone",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void PreparedScriptCombo_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (PreparedScriptTextBox is null)
        {
            return;
        }
        if (PreparedScriptCombo.SelectedItem is not PreparedScriptOption option)
        {
            return;
        }
        PreparedScriptTextBox.Text = option.Text;
    }

    private void StartRecording_OnClick(object sender, RoutedEventArgs e)
    {
        if (_waveIn is not null)
        {
            RecordingStatusText.Text = "Recording already in progress.";
            return;
        }
        if (_micTestIn is not null)
        {
            CleanupMicTestResources();
        }
        RefreshMicrophoneDetectionInfo();
        if (WaveInEvent.DeviceCount <= 0)
        {
            MessageBox.Show(this, "No microphone device found.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
            Directory.CreateDirectory(voicesDir);
            _lastRecordedPath = Path.Combine(voicesDir, $"recorded_{DateTime.Now:yyyyMMdd_HHmmss}.wav");

            _waveIn = new WaveInEvent
            {
                DeviceNumber = 0,
                WaveFormat = new WaveFormat(22050, 16, 1),
                BufferMilliseconds = 120
            };
            _waveWriter = new WaveFileWriter(_lastRecordedPath, _waveIn.WaveFormat);
            _recordingPeak = 0;
            _recordingStartedAtUtc = DateTime.UtcNow;

            _waveIn.DataAvailable += WaveInOnDataAvailable;
            _waveIn.RecordingStopped += WaveInOnRecordingStopped;

            _recordingTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) };
            _recordingTimer.Tick += RecordingTimerOnTick;
            _recordingTimer.Start();

            _waveIn.StartRecording();
            RecordingStatusText.Text = "Recording...";
            CloneStatusTextBlock.Text = "Recording in progress. Read the prepared script.";
        }
        catch (Exception ex)
        {
            CleanupRecordingResources();
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Start recording", ex.Message),
                "Voice Clone",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void StopRecording_OnClick(object sender, RoutedEventArgs e)
    {
        if (_waveIn is null)
        {
            RecordingStatusText.Text = "Recorder idle.";
            return;
        }
        try
        {
            _waveIn.StopRecording();
        }
        catch (Exception ex)
        {
            CleanupRecordingResources();
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Stop recording", ex.Message),
                "Voice Clone",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void StartMicTest_OnClick(object sender, RoutedEventArgs e)
    {
        if (_waveIn is not null)
        {
            RecordingStatusText.Text = "Stop recording before mic test.";
            return;
        }
        if (_micTestIn is not null)
        {
            RecordingStatusText.Text = "Mic test already running.";
            return;
        }

        RefreshMicrophoneDetectionInfo();
        if (WaveInEvent.DeviceCount <= 0)
        {
            MessageBox.Show(this, "No microphone device found.", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            _micTestPeak = 0f;
            _micTestIn = new WaveInEvent
            {
                DeviceNumber = 0,
                WaveFormat = new WaveFormat(22050, 16, 1),
                BufferMilliseconds = 120
            };
            _micTestIn.DataAvailable += MicTestOnDataAvailable;
            _micTestIn.RecordingStopped += MicTestOnRecordingStopped;

            _micTestTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) };
            _micTestTimer.Tick += MicTestTimerOnTick;
            _micTestTimer.Start();

            _micTestIn.StartRecording();
            RecordingStatusText.Text = "Mic test listening...";
        }
        catch (Exception ex)
        {
            CleanupMicTestResources();
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Start mic test", ex.Message),
                "Voice Clone",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void StopMicTest_OnClick(object sender, RoutedEventArgs e)
    {
        if (_micTestIn is null)
        {
            RecordingStatusText.Text = _waveIn is null ? "Recorder idle." : "Recording...";
            return;
        }

        try
        {
            _micTestIn.StopRecording();
        }
        catch (Exception ex)
        {
            CleanupMicTestResources();
            RecordingLevelBar.Value = 0;
            RecordingStatusText.Text = "Recorder idle.";
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Stop mic test", ex.Message),
                "Voice Clone",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void WaveInOnDataAvailable(object? sender, WaveInEventArgs e)
    {
        if (_waveWriter is null)
        {
            return;
        }

        _waveWriter.Write(e.Buffer, 0, e.BytesRecorded);
        _waveWriter.Flush();

        var peak = 0f;
        for (var i = 0; i < e.BytesRecorded - 1; i += 2)
        {
            short sample = (short)(e.Buffer[i] | (e.Buffer[i + 1] << 8));
            var abs = Math.Abs(sample) / 32768f;
            if (abs > peak)
            {
                peak = abs;
            }
        }
        _recordingPeak = Math.Max(_recordingPeak * 0.82f, peak);
    }

    private void WaveInOnRecordingStopped(object? sender, StoppedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            CleanupRecordingResources();
            RecordingStatusText.Text = "Recorder idle.";
            RecordingLevelBar.Value = 0;
            RecordingTimerText.Text = "00:00";

            if (!string.IsNullOrWhiteSpace(_lastRecordedPath) && File.Exists(_lastRecordedPath))
            {
                if (string.IsNullOrWhiteSpace(CloneNameTextBox.Text))
                {
                    CloneNameTextBox.Text = Path.GetFileNameWithoutExtension(_lastRecordedPath);
                }
                CloneStatusTextBlock.Text = $"Recording saved: {Path.GetFileName(_lastRecordedPath)}";
            }

            if (e.Exception is not null)
            {
                MessageBox.Show(this, $"Recording error: {e.Exception.Message}", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        });
    }

    private void MicTestOnDataAvailable(object? sender, WaveInEventArgs e)
    {
        var peak = 0f;
        for (var i = 0; i < e.BytesRecorded - 1; i += 2)
        {
            short sample = (short)(e.Buffer[i] | (e.Buffer[i + 1] << 8));
            var abs = Math.Abs(sample) / 32768f;
            if (abs > peak)
            {
                peak = abs;
            }
        }
        _micTestPeak = Math.Max(_micTestPeak * 0.82f, peak);
    }

    private void MicTestOnRecordingStopped(object? sender, StoppedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            CleanupMicTestResources();
            if (_waveIn is null)
            {
                RecordingLevelBar.Value = 0;
                RecordingStatusText.Text = "Recorder idle.";
            }

            if (e.Exception is not null)
            {
                MessageBox.Show(this, $"Mic test error: {e.Exception.Message}", "Voice Clone", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        });
    }

    private void MicTestTimerOnTick(object? sender, EventArgs e)
    {
        if (_micTestIn is null || _waveIn is not null)
        {
            return;
        }

        RecordingLevelBar.Value = Math.Clamp(_micTestPeak * 100.0, 0.0, 100.0);
    }

    private void RecordingTimerOnTick(object? sender, EventArgs e)
    {
        if (_waveIn is null)
        {
            return;
        }

        var elapsed = DateTime.UtcNow - _recordingStartedAtUtc;
        RecordingTimerText.Text = $"{Math.Min(99, (int)elapsed.TotalMinutes):00}:{elapsed.Seconds:00}";
        RecordingLevelBar.Value = Math.Clamp(_recordingPeak * 100.0, 0.0, 100.0);

        if (elapsed.TotalSeconds >= 60)
        {
            StopRecording_OnClick(this, new RoutedEventArgs());
            CloneStatusTextBlock.Text = "Recording reached 60s and stopped automatically.";
        }
    }

    private void CleanupRecordingResources()
    {
        if (_recordingTimer is not null)
        {
            _recordingTimer.Stop();
            _recordingTimer.Tick -= RecordingTimerOnTick;
            _recordingTimer = null;
        }

        if (_waveIn is not null)
        {
            _waveIn.DataAvailable -= WaveInOnDataAvailable;
            _waveIn.RecordingStopped -= WaveInOnRecordingStopped;
            _waveIn.Dispose();
            _waveIn = null;
        }

        _waveWriter?.Dispose();
        _waveWriter = null;
        _recordingPeak = 0;
    }

    private void CleanupMicTestResources()
    {
        if (_micTestTimer is not null)
        {
            _micTestTimer.Stop();
            _micTestTimer.Tick -= MicTestTimerOnTick;
            _micTestTimer = null;
        }

        if (_micTestIn is not null)
        {
            _micTestIn.DataAvailable -= MicTestOnDataAvailable;
            _micTestIn.RecordingStopped -= MicTestOnRecordingStopped;
            _micTestIn.Dispose();
            _micTestIn = null;
        }

        _micTestPeak = 0;
    }

    private void RefreshMicrophoneDetectionInfo()
    {
        if (MicrophoneDetectionText is null)
        {
            return;
        }

        try
        {
            var count = WaveInEvent.DeviceCount;
            if (count <= 0)
            {
                MicrophoneDetectionText.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
                MicrophoneDetectionText.Text = "Microphone: not detected";
                return;
            }

            var cap = WaveIn.GetCapabilities(0);
            var name = string.IsNullOrWhiteSpace(cap.ProductName) ? "Default input" : cap.ProductName.Trim();
            MicrophoneDetectionText.Foreground = new SolidColorBrush(Color.FromRgb(74, 85, 80));
            MicrophoneDetectionText.Text = count == 1
                ? $"Microphone detected: {name}"
                : $"Microphone detected: {name} (+{count - 1} more)";
        }
        catch (Exception ex)
        {
            MicrophoneDetectionText.Foreground = new SolidColorBrush(Color.FromRgb(170, 47, 47));
            MicrophoneDetectionText.Text = $"Microphone check failed: {ex.Message}";
        }
    }

    private void MainWindow_OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _enhanceAudioCts?.Cancel();
        CleanupRecordingResources();
        CleanupMicTestResources();
        if (_voiceDesignTimer is not null)
        {
            _voiceDesignTimer.Stop();
            _voiceDesignTimer.Tick -= VoiceDesignTimerOnTick;
            _voiceDesignTimer = null;
        }
        if (_voiceDesignDraftAutosaveTimer is not null)
        {
            _voiceDesignDraftAutosaveTimer.Stop();
            _voiceDesignDraftAutosaveTimer.Tick -= VoiceDesignDraftAutosaveTimer_OnTick;
            _voiceDesignDraftAutosaveTimer = null;
        }
        ScriptPrepLlmClient.ForceStopAllLocalServers();
        try
        {
            CollectProjectFromUi();
            if (string.IsNullOrWhiteSpace(_currentProjectPath))
            {
                var baseName = string.IsNullOrWhiteSpace(_project.Name) ? "New_Audiobook_Project" : SanitizeFileName(_project.Name);
                if (string.IsNullOrWhiteSpace(baseName))
                {
                    baseName = "New_Audiobook_Project";
                }
                _currentProjectPath = MakeUniqueProjectPath(baseName);
            }

            Task.Run(() => _projectStore.SaveAsync(_project, _currentProjectPath!)).GetAwaiter().GetResult();
            if (!_config.RecentProjects.Any(path => string.Equals(path, _currentProjectPath, StringComparison.OrdinalIgnoreCase)))
            {
                _config.RecentProjects.Insert(0, _currentProjectPath!);
                _config.RecentProjects = _config.RecentProjects.Take(20).ToList();
            }
            Task.Run(() => _configStore.SaveAsync(_config)).GetAwaiter().GetResult();
        }
        catch
        {
            // Keep closing even if final autosave fails.
        }
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        if (Application.Current is not null)
        {
            Application.Current.Shutdown();
        }
    }

    private static bool ValidateLocalVoiceForGeneration(string voicePath, out string error)
    {
        if (!voicePath.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
        {
            error = "Local model requires WAV voice files. Please select/import a .wav voice.";
            return false;
        }

        if (!TryInspectWav(voicePath, out var wav, out error))
        {
            return false;
        }

        if (wav.DurationSeconds < 2.0)
        {
            error = "Voice sample is too short. Use at least 2 seconds.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    private void NormalizeLocalVoiceLibrary()
    {
        var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
        if (!Directory.Exists(voicesDir))
        {
            return;
        }

        var legacySources = Directory.GetFiles(voicesDir, "*.*", SearchOption.TopDirectoryOnly)
            .Where(IsConvertibleVoiceFile)
            .ToList();

        foreach (var sourcePath in legacySources)
        {
            if (!TryNormalizeLocalVoiceFile(sourcePath, out _, out var error))
            {
                Debug.WriteLine($"Failed to normalize local voice '{sourcePath}': {error}");
            }
        }
    }

    private bool TryEnsureProjectVoiceIsWav(out string error)
    {
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(_project.VoicePath) ||
            !File.Exists(_project.VoicePath) ||
            _project.VoicePath.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!TryNormalizeLocalVoiceFile(_project.VoicePath, out var normalizedPath, out error))
        {
            return false;
        }

        _project.VoicePath = normalizedPath;
        LoadVoiceList();
        SelectVoiceByPath(normalizedPath);
        _ = AutoSaveCurrentProjectAsync();
        return true;
    }

    private static bool TryNormalizeLocalVoiceFile(string sourcePath, out string outputPath, out string error)
    {
        outputPath = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            error = "Selected voice file no longer exists.";
            return false;
        }

        if (sourcePath.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
        {
            outputPath = sourcePath;
            return true;
        }

        if (!IsConvertibleVoiceFile(sourcePath))
        {
            error = "Only WAV, MP3, FLAC, and M4A voices are supported.";
            return false;
        }

        try
        {
            var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
            var backupDir = Path.Combine(voicesDir, "backupvoice");
            Directory.CreateDirectory(voicesDir);
            Directory.CreateDirectory(backupDir);

            var sourceFullPath = Path.GetFullPath(sourcePath);
            var backupPath = EnsureManagedBackupVoiceCopy(sourceFullPath, backupDir);
            var targetPath = Path.Combine(
                voicesDir,
                $"{Path.GetFileNameWithoutExtension(backupPath) ?? "voice_source"}.wav");

            if (!File.Exists(targetPath) &&
                !TryConvertAudioToWav(backupPath, targetPath, out outputPath, out error))
            {
                return false;
            }

            outputPath = targetPath;
            return true;
        }
        catch (Exception ex)
        {
            error = $"Voice conversion failed: {ex.Message}";
            return false;
        }
    }

    private static bool TryInspectWav(string path, out WavInfo info, out string error)
    {
        info = new WavInfo(0, 0, 0, 0);
        error = string.Empty;
        try
        {
            using var fs = File.OpenRead(path);
            using var br = new BinaryReader(fs);

            var riff = new string(br.ReadChars(4));
            _ = br.ReadUInt32();
            var wave = new string(br.ReadChars(4));
            if (!string.Equals(riff, "RIFF", StringComparison.Ordinal) || !string.Equals(wave, "WAVE", StringComparison.Ordinal))
            {
                error = "Voice file is not a valid WAV (RIFF/WAVE).";
                return false;
            }

            ushort channels = 0;
            uint sampleRate = 0;
            ushort bitsPerSample = 0;
            byte[]? pcmBytes = null;

            while (fs.Position + 8 <= fs.Length)
            {
                var chunkId = new string(br.ReadChars(4));
                var chunkSize = br.ReadUInt32();
                if (chunkSize > int.MaxValue || fs.Position + chunkSize > fs.Length)
                {
                    error = "Invalid WAV chunk structure.";
                    return false;
                }

                if (chunkId == "fmt ")
                {
                    var audioFormat = br.ReadUInt16();
                    channels = br.ReadUInt16();
                    sampleRate = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt16();
                    bitsPerSample = br.ReadUInt16();
                    var remaining = (int)chunkSize - 16;
                    if (remaining > 0)
                    {
                        br.ReadBytes(remaining);
                    }

                    if (audioFormat != 1)
                    {
                        error = "Only PCM WAV is supported for local voice cloning.";
                        return false;
                    }
                }
                else if (chunkId == "data")
                {
                    pcmBytes = br.ReadBytes((int)chunkSize);
                }
                else
                {
                    br.ReadBytes((int)chunkSize);
                }

                if ((chunkSize & 1) == 1 && fs.Position < fs.Length)
                {
                    fs.Position++;
                }
            }

            if (channels is not (1 or 2))
            {
                error = "WAV must be mono or stereo.";
                return false;
            }
            if (bitsPerSample != 16)
            {
                error = "WAV must be 16-bit PCM.";
                return false;
            }
            if (sampleRate < 8000)
            {
                error = "WAV sample rate is too low.";
                return false;
            }
            if (pcmBytes is null || pcmBytes.Length < 2048)
            {
                error = "WAV data is too short.";
                return false;
            }

            var bytesPerSample = bitsPerSample / 8;
            var totalSamples = pcmBytes.Length / bytesPerSample / channels;
            if (totalSamples <= 0)
            {
                error = "WAV contains no valid samples.";
                return false;
            }

            var durationSeconds = (double)totalSamples / sampleRate;
            var sumSq = 0.0;
            var sampleCount = 0;

            for (var i = 0; i + 1 < pcmBytes.Length; i += 2 * channels)
            {
                short s = BitConverter.ToInt16(pcmBytes, i);
                if (channels == 2 && i + 3 < pcmBytes.Length)
                {
                    short r = BitConverter.ToInt16(pcmBytes, i + 2);
                    s = (short)((s + r) / 2);
                }
                var v = s / 32768.0;
                sumSq += v * v;
                sampleCount++;
            }

            var rms = sampleCount > 0 ? Math.Sqrt(sumSq / sampleCount) : 0.0;
            info = new WavInfo(channels, sampleRate, bitsPerSample, durationSeconds, rms);
            return true;
        }
        catch (Exception ex)
        {
            error = $"Unable to parse WAV: {ex.Message}";
            return false;
        }
    }

    private static bool TryConvertAudioToWav(string sourcePath, out string outputPath, out string error)
    {
        outputPath = string.Empty;
        error = string.Empty;
        try
        {
            var voicesDir = Path.Combine(RuntimePaths.AppRoot, "voices");
            Directory.CreateDirectory(voicesDir);

            var baseName = Path.GetFileNameWithoutExtension(sourcePath);
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "voice_source";
            }

            var safeName = Regex.Replace(baseName, @"[^a-zA-Z0-9_\- ]", "_").Trim();
            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = "voice_source";
            }

            var target = Path.Combine(voicesDir, $"{safeName}_converted.wav");
            if (File.Exists(target))
            {
                var suffix = 2;
                string candidate;
                do
                {
                    candidate = Path.Combine(voicesDir, $"{safeName}_converted_{suffix}.wav");
                    suffix++;
                } while (File.Exists(candidate));
                target = candidate;
            }

            using var reader = new MediaFoundationReader(sourcePath);
            var outFormat = new WaveFormat(22050, 16, 1);
            using var resampler = new MediaFoundationResampler(reader, outFormat)
            {
                ResamplerQuality = 60
            };
            WaveFileWriter.CreateWaveFile(target, resampler);

            outputPath = target;
            return true;
        }
        catch (Exception ex)
        {
            error = $"Audio conversion failed: {ex.Message}";
            return false;
        }
    }

    private static bool TryConvertAudioToWav(string sourcePath, string targetPath, out string outputPath, out string error)
    {
        outputPath = string.Empty;
        error = string.Empty;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath) ?? Path.Combine(RuntimePaths.AppRoot, "voices"));

            using var reader = new MediaFoundationReader(sourcePath);
            var outFormat = new WaveFormat(22050, 16, 1);
            using var resampler = new MediaFoundationResampler(reader, outFormat)
            {
                ResamplerQuality = 60
            };
            WaveFileWriter.CreateWaveFile(targetPath, resampler);

            outputPath = targetPath;
            return true;
        }
        catch (Exception ex)
        {
            error = $"Audio conversion failed: {ex.Message}";
            return false;
        }
    }

    private static string EnsureManagedBackupVoiceCopy(string sourceFullPath, string backupDir)
    {
        var sourceDirectory = Path.GetDirectoryName(sourceFullPath) ?? string.Empty;
        if (string.Equals(Path.GetFullPath(sourceDirectory), Path.GetFullPath(backupDir), StringComparison.OrdinalIgnoreCase))
        {
            return sourceFullPath;
        }

        var safeName = SanitizeVoiceFileStem(Path.GetFileNameWithoutExtension(sourceFullPath));
        var ext = Path.GetExtension(sourceFullPath);
        var backupPath = GetUniqueVoiceFilePath(backupDir, safeName, ext);

        var voicesDir = Path.GetDirectoryName(backupDir) ?? string.Empty;
        if (string.Equals(Path.GetFullPath(sourceDirectory), Path.GetFullPath(voicesDir), StringComparison.OrdinalIgnoreCase))
        {
            File.Move(sourceFullPath, backupPath);
        }
        else
        {
            File.Copy(sourceFullPath, backupPath, overwrite: false);
        }

        return backupPath;
    }

    private static string GetUniqueVoiceFilePath(string directory, string baseName, string extension)
    {
        var candidate = Path.Combine(directory, $"{baseName}{extension}");
        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var suffix = 2;
        do
        {
            candidate = Path.Combine(directory, $"{baseName}_{suffix}{extension}");
            suffix++;
        } while (File.Exists(candidate));

        return candidate;
    }

    private static string SanitizeVoiceFileStem(string? baseName)
    {
        var safeName = Regex.Replace(baseName ?? "voice_source", @"[^a-zA-Z0-9_\- ]", "_").Trim();
        return string.IsNullOrWhiteSpace(safeName) ? "voice_source" : safeName;
    }

    private static bool IsConvertibleVoiceFile(string path)
    {
        return path.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".flac", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".m4a", StringComparison.OrdinalIgnoreCase);
    }

    private string? PromptForProjectName(string? currentName)
    {
        var dlg = new Window
        {
            Title = "Create New Project",
            Width = 460,
            Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            Background = System.Windows.Media.Brushes.White
        };

        var panel = new System.Windows.Controls.Grid { Margin = new Thickness(16) };
        panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var label = new System.Windows.Controls.TextBlock
        {
            Text = "Project Name",
            FontSize = 13
        };
        System.Windows.Controls.Grid.SetRow(label, 0);
        panel.Children.Add(label);

        var input = new System.Windows.Controls.TextBox
        {
            Margin = new Thickness(0, 8, 0, 0),
            Height = 34,
            Text = string.IsNullOrWhiteSpace(currentName) ? "New Audiobook Project" : currentName.Trim()
        };
        System.Windows.Controls.Grid.SetRow(input, 1);
        panel.Children.Add(input);

        var buttons = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };

        var cancelBtn = new System.Windows.Controls.Button
        {
            Content = "Cancel",
            Width = 90,
            Height = 32,
            Margin = new Thickness(0, 0, 8, 0)
        };
        cancelBtn.Click += (_, _) => { dlg.DialogResult = false; dlg.Close(); };

        var createBtn = new System.Windows.Controls.Button
        {
            Content = "Create",
            Width = 100,
            Height = 34
        };
        createBtn.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(input.Text))
            {
                MessageBox.Show(dlg, "Project name is required.", "Create Project", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            dlg.DialogResult = true;
            dlg.Close();
        };

        buttons.Children.Add(cancelBtn);
        buttons.Children.Add(createBtn);
        System.Windows.Controls.Grid.SetRow(buttons, 2);
        panel.Children.Add(buttons);

        dlg.Content = panel;
        input.SelectAll();
        input.Focus();

        var result = dlg.ShowDialog();
        if (result != true)
        {
            return null;
        }
        return input.Text.Trim();
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
        return new string(chars).Trim();
    }

    private static string MakeUniqueProjectPath(string baseFileName)
    {
        var projectsDir = RuntimePaths.ProjectsDir;
        Directory.CreateDirectory(projectsDir);
        var initial = Path.Combine(projectsDir, $"{baseFileName}{JsonProjectStore.Extension}");
        if (!File.Exists(initial))
        {
            return initial;
        }

        var index = 2;
        while (true)
        {
            var candidate = Path.Combine(projectsDir, $"{baseFileName}_{index}{JsonProjectStore.Extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
            index++;
        }
    }

    private static string ResolveOutputDir(string? configured)
    {
        var value = string.IsNullOrWhiteSpace(configured) ? "output" : configured.Trim();
        if (Path.IsPathRooted(value))
        {
            return value;
        }

        return Path.Combine(RuntimePaths.AppRoot, value);
    }

    private static string ResolveProjectOutputDir(string? configured, string? projectName)
    {
        var baseDir = ResolveOutputDir(configured);
        var name = string.IsNullOrWhiteSpace(projectName) ? "Project" : SanitizeFileName(projectName);
        if (string.IsNullOrWhiteSpace(name))
        {
            name = "Project";
        }
        return Path.Combine(baseDir, name);
    }

    private string ResolveEnhanceOutputDir(string? preferredBaseOutput = null, bool forceFromBaseOutput = false)
    {
        var raw = forceFromBaseOutput ? string.Empty : (EnhanceOutputFolderTextBox?.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(raw))
        {
            return ResolveOutputDir(raw);
        }

        var baseDir = string.IsNullOrWhiteSpace(preferredBaseOutput)
            ? ResolveOutputDir(_project.OutputDir)
            : ResolveOutputDir(preferredBaseOutput);
        return Path.Combine(baseDir, "enhanced_audio");
    }

    private static string NormalizePreferDevice(string? value)
    {
        return (value ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "gpu" => "gpu",
            "cuda" => "gpu",
            "cpu" => "cpu",
            _ => "auto"
        };
    }

    private async Task AutoSaveCurrentProjectAsync()
    {
        try
        {
            var hadPath = !string.IsNullOrWhiteSpace(_currentProjectPath);
            if (string.IsNullOrWhiteSpace(_currentProjectPath))
            {
                var baseName = string.IsNullOrWhiteSpace(_project.Name) ? "New_Audiobook_Project" : SanitizeFileName(_project.Name);
                if (string.IsNullOrWhiteSpace(baseName))
                {
                    baseName = "New_Audiobook_Project";
                }
                _currentProjectPath = MakeUniqueProjectPath(baseName);
            }

            CollectProjectFromUi();
            await _projectStore.SaveAsync(_project, _currentProjectPath!);
            await RegisterRecentProjectAsync(_currentProjectPath!);
            if (!hadPath)
            {
                RefreshProjectSelector();
            }
        }
        catch
        {
            // Non-fatal: explicit Save still available.
        }
    }

    private async void ProjectSelectorCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isRefreshingProjectSelector)
        {
            return;
        }
        var index = ProjectSelectorCombo.SelectedIndex;
        if (index < 0 || index >= _projectSelectorPaths.Count)
        {
            return;
        }
        var path = _projectSelectorPaths[index];
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return;
        }
        if (string.Equals(_currentProjectPath, path, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            _project = await _projectStore.LoadAsync(path);
            _currentProjectPath = path;
            _config.LastOpenDir = Path.GetDirectoryName(path) ?? RuntimePaths.ProjectsDir;
            await RegisterRecentProjectAsync(path);
            LoadProjectToUi(_project);
            RefreshProjectSelector();
            ProjectSelectorCombo.Text = _project.Name;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Project load", ex.Message),
                "Project Load",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static BitmapSource TrimTransparentPadding(BitmapSource source)
    {
        if (source.Format != PixelFormats.Bgra32 && source.Format != PixelFormats.Pbgra32)
        {
            source = new FormatConvertedBitmap(source, PixelFormats.Bgra32, null, 0);
        }

        var width = source.PixelWidth;
        var height = source.PixelHeight;
        if (width <= 0 || height <= 0)
        {
            return source;
        }

        var stride = width * 4;
        var pixels = new byte[height * stride];
        source.CopyPixels(pixels, stride, 0);

        var minX = width;
        var minY = height;
        var maxX = -1;
        var maxY = -1;
        const byte alphaThreshold = 8;

        for (var y = 0; y < height; y++)
        {
            var row = y * stride;
            for (var x = 0; x < width; x++)
            {
                var a = pixels[row + x * 4 + 3];
                if (a <= alphaThreshold)
                {
                    continue;
                }

                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }
        }

        if (maxX < minX || maxY < minY)
        {
            return source;
        }

        var cropWidth = Math.Max(1, maxX - minX + 1);
        var cropHeight = Math.Max(1, maxY - minY + 1);
        return new CroppedBitmap(source, new Int32Rect(minX, minY, cropWidth, cropHeight));
    }

    private readonly record struct PreparedScriptOption(string Title, string Text)
    {
        public override string ToString() => Title;
    }

    private readonly record struct StylePresetOption(string Key, string DisplayName);

    private enum ChunkProfile
    {
        Chatterbox,
        Qwen3Tts,
        KittenTts
    }

    private readonly record struct SubtitleChunkTiming(string Text, double StartSeconds, double EndSeconds);
    private readonly record struct SubtitleCue(double StartSeconds, double EndSeconds, string Text);
    private readonly record struct TextChunk(string Text, bool ParagraphEnd, bool IsDialogue = false);
    private readonly record struct TimedUnit(string Text, bool IsDialogue, bool ForceBoundaryAfter = false);

    private sealed class AudioEnhanceInputFileEntry : INotifyPropertyChanged
    {
        private bool _isSelected;

        public AudioEnhanceInputFileEntry(string fullPath, bool isSelected)
        {
            FullPath = fullPath;
            _isSelected = isSelected;
        }

        public string FullPath { get; }
        public string FileName => Path.GetFileName(FullPath);

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    private sealed class InputFileEntry : INotifyPropertyChanged
    {
        private bool _isSelected;
        private bool _isPrepared;
        private int _preparedPartCount;
        private bool _isPreparedStale;
        private bool _canPrepareScript;

        public InputFileEntry(string fullPath, bool isSelected)
        {
            FullPath = fullPath;
            _isSelected = isSelected;
        }

        public string FullPath { get; }

        public string FileName => Path.GetFileName(FullPath);

        public string DisplayName
        {
            get
            {
                if (!_canPrepareScript)
                {
                    return FileName;
                }

                if (!_isPrepared)
                {
                    return $"{FileName}  [Not prepared]";
                }

                var stale = _isPreparedStale ? " (stale)" : string.Empty;
                return $"{FileName}  [Prepared: {_preparedPartCount} part(s){stale}]";
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public bool IsPrepared
        {
            get => _isPrepared;
            set
            {
                if (_isPrepared == value)
                {
                    return;
                }

                _isPrepared = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsPrepared)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
            }
        }

        public int PreparedPartCount
        {
            get => _preparedPartCount;
            set
            {
                if (_preparedPartCount == value)
                {
                    return;
                }

                _preparedPartCount = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PreparedPartCount)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
            }
        }

        public bool IsPreparedStale
        {
            get => _isPreparedStale;
            set
            {
                if (_isPreparedStale == value)
                {
                    return;
                }

                _isPreparedStale = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsPreparedStale)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
            }
        }

        public bool CanPrepareScript
        {
            get => _canPrepareScript;
            set
            {
                if (_canPrepareScript == value)
                {
                    return;
                }

                _canPrepareScript = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanPrepareScript)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PrepareButtonVisibility)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
            }
        }

        public Visibility PrepareButtonVisibility => _canPrepareScript ? Visibility.Visible : Visibility.Collapsed;

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    private sealed record FileChoice(string DisplayName, string FullPath)
    {
        public override string ToString() => DisplayName;
    }

    /// <summary>Voice list item. ShowRefCheckmark is true only when Qwen3 is selected and this voice has a .ref.txt transcript.</summary>
    private sealed record VoiceItem(string DisplayName, string FullPath, bool HasRefText, bool ShowRefCheckmark)
    {
        public override string ToString() => DisplayName;
    }

    private readonly record struct WavInfo(ushort Channels, uint SampleRate, ushort BitsPerSample, double DurationSeconds, double Rms = 0.0);

    private void TraceGenerate(string message)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";

        try
        {
            var logPath = Path.Combine(RuntimePaths.AppRoot, "generate_runtime.log");
            lock (GenerateLogLock)
            {
                File.AppendAllText(logPath, line + Environment.NewLine);
            }
        }
        catch
        {
            // Non-fatal diagnostics only.
        }
    }

    private static string NormalizeGenerationError(string message)
    {
        return UserMessageFormatter.FormatOperationError("Generation", message);
    }
}

public sealed class FileNameOnlyConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var raw = value?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        return Path.GetFileName(raw);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value?.ToString() ?? string.Empty;
    }
}
