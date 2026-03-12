using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO.Compression;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using App.Core.Models;
using App.Core.Runtime;
using App.Diagnostics;
using App.Inference;
using App.Storage;
using Microsoft.ML.OnnxRuntime;

namespace AudiobookCreator.UI;

public partial class SettingsWindow : Window
{
    private const string EspeakNgMsiUrl = "https://github.com/espeak-ng/espeak-ng/releases/download/1.52.0/espeak-ng.msi";
    private const string KittenTtsWheelUrl = "https://github.com/KittenML/KittenTTS/releases/download/0.8.1/kittentts-0.8.1-py3-none-any.whl";
    private readonly AppConfig _original;
    private readonly HuggingFaceModelDownloader _downloader = new();
    private readonly Dictionary<string, SynthesisSettings> _modelProfiles;
    private readonly SystemProfile _systemProfile;
    private bool _isApplyingQwenVariant;
    private bool _llmEventsHooked;
    private bool _audioEnhanceEventsHooked;
    private bool _isAutoFetchingApiTtsModels;
    private CancellationTokenSource? _llmInstallCts;
    private CancellationTokenSource? _audioEnhanceInstallCts;
    private IReadOnlyList<string>? _openAiFetchedLlmModels;
    private IReadOnlyList<string>? _alibabaFetchedLlmModels;
    private IReadOnlyList<string>? _openAiFetchedTtsModels;
    private IReadOnlyList<string>? _alibabaFetchedTtsModels;
    private bool _removeAlibabaApiKeyRequested;
    private bool _removeOpenAiApiKeyRequested;

    private static readonly IReadOnlyList<QwenVariantOption> QwenVariants = new[]
    {
        new QwenVariantOption("auto", "Auto (Recommended)", "", "", 0, 0, 0, 0),
        new QwenVariantOption("qwen_py_17b", "Qwen3 1.7B Base (Python Worker)", "Qwen/Qwen3-TTS-12Hz-1.7B-Base", "", 13.0, 24.0, 24, 10),
        new QwenVariantOption("qwen_py_17b_customvoice", "Qwen3 1.7B CustomVoice (Python Worker)", "Qwen/Qwen3-TTS-12Hz-1.7B-CustomVoice", "", 13.0, 24.0, 24, 10),
        new QwenVariantOption("qwen_py_06b", "Qwen3 0.6B Base (Python Worker)", "Qwen/Qwen3-TTS-12Hz-0.6B-Base", "", 4.5, 9.0, 16, 8),
        new QwenVariantOption("qwen_py_17b_voicedesign", "Qwen3 1.7B VoiceDesign (Python Worker)", "Qwen/Qwen3-TTS-12Hz-1.7B-VoiceDesign", "", 13.0, 24.0, 24, 10)
    };

    private static readonly (string Key, string DisplayName, string ApiPreset)[] AlibabaApiVariants = new[]
    {
        ("alibaba_qwen3_tts", "Qwen3 TTS Flash", "alibaba_qwen3_tts"),
        ("alibaba_qwen3_tts_instruct", "Qwen3 TTS Instruct Flash", "alibaba_qwen3_tts_instruct"),
        ("alibaba_qwen3_tts_vc", "Qwen3 TTS VC", "alibaba_qwen3_tts_vc"),
        ("alibaba_voice_design", "Qwen VoiceDesign (Create Voice)", "alibaba_voice_design"),
        ("alibaba_qwen_tts_latest", "Qwen TTS Latest (Legacy)", "alibaba_qwen_tts_latest")
    };

    private static readonly string[] LocalLlmModelOptions =
    {
        "qwen2.5-7b-instruct",
        "qwen2.5-14b-instruct",
        "qwen3-8b-instruct",
        "qwen3-14b-instruct"
    };

    private static readonly string[] OpenAiLlmModelOptions =
    {
        "gpt-5.2",
        "gpt-5.1",
        "gpt-5",
        "gpt-5-mini",
        "gpt-5-nano",
        "gpt-4.1",
        "gpt-4.1-mini",
        "gpt-4o-mini"
    };

    private static readonly string[] AlibabaLlmModelOptions =
    {
        "qwen-plus",
        "qwen-max",
        "qwen-turbo"
    };

    private static readonly string[] OpenAiTtsModelOptions =
    {
        "gpt-4o-mini-tts",
        "tts-1",
        "tts-1-hd"
    };

    private static readonly string[] AlibabaTtsModelOptions =
    {
        "qwen3-tts-flash",
        "qwen3-tts-instruct-flash",
        "qwen3-tts-vc-2026-01-22",
        "qwen3-tts-vd-2026-01-26",
        "qwen-tts-latest"
    };

    private static readonly string[] AudioEnhanceVariantOptions =
    {
        "audiox_base",
        "audiox_maf",
        "audiox_maf_mmdit"
    };

    private static readonly Dictionary<string, string[]> AudioEnhanceVariantRepoCandidates = new(StringComparer.OrdinalIgnoreCase)
    {
        ["audiox_base"] = new[]
        {
            "HKUSTAudio/AudioX",
            "Zeyue7/AudioX"
        },
        ["audiox_maf"] = new[]
        {
            "HKUSTAudio/AudioX-MAF"
        },
        ["audiox_maf_mmdit"] = new[]
        {
            "HKUSTAudio/AudioX-MAF-MMDiT"
        }
    };

    public AppConfig? UpdatedConfig { get; private set; }

    private readonly string? _currentVoicePath;
    private string? _selectedTranscriptVoicePath;

    public SettingsWindow(AppConfig config, string? currentVoicePath = null)
    {
        _original = Clone(config);
        _currentVoicePath = currentVoicePath;
        _modelProfiles = CloneProfiles(_original.ModelProfiles);
        _systemProfile = SystemProbe.Detect();
        InitializeComponent();
        RemoveNoAudioXSettingsSections();
        InitializeQwenVariantOptions();
        LoadUiFromConfig(_original);
        OpenAiTtsModelCombo.SelectionChanged += ApiTtsModelCombo_OnSelectionChanged;
        AlibabaTtsModelCombo.SelectionChanged += ApiTtsModelCombo_OnSelectionChanged;
        Loaded += SettingsWindow_OnLoaded;
        HookLlmUiEvents();
        HookAudioEnhanceUiEvents();
        LoadVoiceTranscriptSection();
    }

    private static bool SupportsPythonAlternativePreset(string preset)
    {
        var key = (preset ?? string.Empty).Trim().ToLowerInvariant();
        return key is "chatterbox_onnx" or "kitten_tts";
    }

    private static string NormalizeLocalModelBackendChoice(string? value)
    {
        return string.Equals((value ?? string.Empty).Trim(), "python", StringComparison.OrdinalIgnoreCase)
            ? "python"
            : "onnx";
    }

    private string ResolveRecommendedLocalBackend(string preset)
    {
        if (!SupportsPythonAlternativePreset(preset))
        {
            return "onnx";
        }

        return _systemProfile.GpuVendor switch
        {
            "amd" => "python",
            "intel" => "python",
            _ => "onnx"
        };
    }

    private string ResolveEffectiveLocalBackend(AppConfig cfg)
    {
        var preset = (cfg.LocalModelPreset ?? SelectedPreset()).Trim();
        if (!SupportsPythonAlternativePreset(preset))
        {
            return "onnx";
        }

        var configured = NormalizeLocalModelBackendChoice(cfg.LocalModelBackend);
        if (!string.IsNullOrWhiteSpace(cfg.LocalModelBackend))
        {
            return configured;
        }

        return ResolveRecommendedLocalBackend(preset);
    }

    private static string ResolveRepoForPresetAndBackend(string preset, string backend)
    {
        var normalizedPreset = (preset ?? string.Empty).Trim().ToLowerInvariant();
        var normalizedBackend = NormalizeLocalModelBackendChoice(backend);
        return normalizedPreset switch
        {
            "chatterbox_onnx" when normalizedBackend == "python" => "ResembleAI/chatterbox",
            "chatterbox_onnx" => "onnx-community/chatterbox-ONNX",
            "kitten_tts" => "KittenML/kitten-tts-mini-0.8",
            _ => string.Empty
        };
    }

    private string SelectedLocalBackendChoice()
    {
        if (LocalModelBackendCombo?.SelectedItem is ComboBoxItem { Tag: string tag } && !string.IsNullOrWhiteSpace(tag))
        {
            return NormalizeLocalModelBackendChoice(tag);
        }

        return ResolveRecommendedLocalBackend(SelectedPreset());
    }

    private void SetSelectedLocalBackendChoice(string backend)
    {
        var normalized = NormalizeLocalModelBackendChoice(backend);
        if (LocalModelBackendCombo is null)
        {
            return;
        }

        foreach (var item in LocalModelBackendCombo.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), normalized, StringComparison.OrdinalIgnoreCase))
            {
                LocalModelBackendCombo.SelectedItem = item;
                return;
            }
        }

        LocalModelBackendCombo.SelectedIndex = 0;
    }

    private void ApplyLocalBackendSelectionToRepo()
    {
        var preset = SelectedPreset();
        if (!SupportsPythonAlternativePreset(preset))
        {
            return;
        }

        ModelRepoTextBox.Text = ResolveRepoForPresetAndBackend(preset, SelectedLocalBackendChoice());
        AdditionalRepoTextBox.Text = string.Empty;
    }

    private void RefreshLocalBackendUi(AppConfig cfg)
    {
        var preset = SelectedPreset();
        var supportsPythonAlt = SupportsPythonAlternativePreset(preset);
        var allowManualChoice = supportsPythonAlt && string.Equals(_systemProfile.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase);
        var panelVisibility = supportsPythonAlt ? Visibility.Visible : Visibility.Collapsed;

        if (LocalBackendLabel is not null)
        {
            LocalBackendLabel.Visibility = panelVisibility;
        }

        if (LocalBackendPanel is not null)
        {
            LocalBackendPanel.Visibility = panelVisibility;
        }

        if (!supportsPythonAlt)
        {
            return;
        }

        var effectiveBackend = ResolveEffectiveLocalBackend(cfg);
        SetSelectedLocalBackendChoice(effectiveBackend);
        LocalModelBackendCombo.IsEnabled = allowManualChoice;

        var vendorLabel = (_systemProfile.GpuVendor ?? "none").ToUpperInvariant();
        if (allowManualChoice)
        {
            LocalBackendHintTextBlock.Text = "NVIDIA detected. ONNX is the default path; you can switch this model to Python before downloading.";
        }
        else if (string.Equals(effectiveBackend, "python", StringComparison.OrdinalIgnoreCase))
        {
            LocalBackendHintTextBlock.Text = $"{vendorLabel} GPU detected. This model will use the Python backend when you download it.";
        }
        else
        {
            LocalBackendHintTextBlock.Text = "This model will use the ONNX backend when you download it.";
        }
    }

    private static string ResolvePythonEnvNameForPreset(string preset)
    {
        return (preset ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "chatterbox_onnx" => "python_chatterbox",
            "kitten_tts" => "python_kitten",
            _ => "python_qwen"
        };
    }

    private static string ResolvePythonBootstrapBaseName(string envName)
    {
        return (envName ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "python_chatterbox" => "python_chatterbox_base",
            "python_kitten" => "python_kitten_base",
            _ => "python_qwen"
        };
    }

    private static string GetRefTextPathForVoice(string? voicePath)
    {
        if (string.IsNullOrWhiteSpace(voicePath) || !File.Exists(voicePath))
            return string.Empty;
        var dir = Path.GetDirectoryName(voicePath);
        var stem = Path.GetFileNameWithoutExtension(voicePath);
        return string.IsNullOrEmpty(dir) ? string.Empty : Path.Combine(dir, stem + ".ref.txt");
    }

    private void LoadVoiceTranscriptSection()
    {
        PopulateVoiceTranscriptVoiceList();
        var targetVoicePath = _selectedTranscriptVoicePath ?? _currentVoicePath;
        var refPath = GetRefTextPathForVoice(targetVoicePath);
        var canEdit = !string.IsNullOrEmpty(refPath);
        if (VoiceTranscriptSection is not null)
            VoiceTranscriptSection.Visibility = canEdit ? Visibility.Visible : Visibility.Collapsed;
        if (VoiceTranscriptTargetLabel is not null)
            VoiceTranscriptTargetLabel.Text = canEdit && !string.IsNullOrEmpty(targetVoicePath)
                ? "Save to voice: " + Path.GetFileName(targetVoicePath)
                : string.Empty;
        if (canEdit && VoiceTranscriptTextBox is not null)
        {
            try
            {
                VoiceTranscriptTextBox.Text = File.Exists(refPath) ? File.ReadAllText(refPath) : string.Empty;
            }
            catch
            {
                VoiceTranscriptTextBox.Text = string.Empty;
            }
        }
    }

    private void RemoveNoAudioXSettingsSections()
    {
        if (AppFeatureFlags.AudioEnhancementEnabled || AudioEnhancementSection is null)
        {
            return;
        }

        if (AudioEnhancementSection.Parent is Panel panel)
        {
            panel.Children.Remove(AudioEnhancementSection);
        }
    }

    private void PopulateVoiceTranscriptVoiceList()
    {
        if (VoiceTranscriptVoiceCombo is null)
            return;

        var current = _selectedTranscriptVoicePath ?? _currentVoicePath;
        var candidates = new List<string>();

        void AddVoiceDir(string? dir)
        {
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
                return;
            foreach (var file in Directory.EnumerateFiles(dir, "*.wav", SearchOption.TopDirectoryOnly))
            {
                if (!candidates.Any(x => string.Equals(x, file, StringComparison.OrdinalIgnoreCase)))
                    candidates.Add(file);
            }
        }

        AddVoiceDir(!string.IsNullOrWhiteSpace(_currentVoicePath) ? Path.GetDirectoryName(_currentVoicePath) : null);
        AddVoiceDir(Path.Combine(RuntimePaths.AppRoot, "voices"));

        VoiceTranscriptVoiceCombo.SelectionChanged -= VoiceTranscriptVoiceCombo_OnSelectionChanged;
        VoiceTranscriptVoiceCombo.Items.Clear();

        foreach (var path in candidates.OrderBy(Path.GetFileName))
        {
            VoiceTranscriptVoiceCombo.Items.Add(new ComboBoxItem
            {
                Content = Path.GetFileName(path),
                Tag = path
            });
        }

        if (VoiceTranscriptVoiceCombo.Items.Count > 0)
        {
            var selectedItem = VoiceTranscriptVoiceCombo.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Tag as string, current, StringComparison.OrdinalIgnoreCase))
                ?? VoiceTranscriptVoiceCombo.Items.OfType<ComboBoxItem>().First();

            VoiceTranscriptVoiceCombo.SelectedItem = selectedItem;
            _selectedTranscriptVoicePath = selectedItem.Tag as string;
            VoiceTranscriptVoiceCombo.IsEnabled = true;
        }
        else
        {
            _selectedTranscriptVoicePath = null;
            VoiceTranscriptVoiceCombo.IsEnabled = false;
        }

        VoiceTranscriptVoiceCombo.SelectionChanged += VoiceTranscriptVoiceCombo_OnSelectionChanged;
    }

    private void VoiceTranscriptVoiceCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (VoiceTranscriptVoiceCombo?.SelectedItem is ComboBoxItem item)
        {
            _selectedTranscriptVoicePath = item.Tag as string;
            LoadVoiceTranscriptSection();
        }
    }

    private void SaveVoiceTranscriptButton_OnClick(object sender, RoutedEventArgs e)
    {
        var refPath = GetRefTextPathForVoice(_selectedTranscriptVoicePath ?? _currentVoicePath);
        if (string.IsNullOrEmpty(refPath))
        {
            MessageBox.Show(this, "No voice file selected or not a file-based voice.", "Voice transcript", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        try
        {
            var text = VoiceTranscriptTextBox?.Text?.Trim() ?? string.Empty;
            File.WriteAllText(refPath, text);
            MessageBox.Show(this, "Transcript saved next to the voice file.", "Voice transcript", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Voice transcript save", ex.Message),
                "Voice Transcript",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void LoadUiFromConfig(AppConfig cfg)
    {
        NormalizeModelSettings(cfg);
        _openAiFetchedLlmModels = cfg.CachedOpenAiLlmModels is { Count: > 0 }
            ? cfg.CachedOpenAiLlmModels.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList()
            : null;
        _alibabaFetchedLlmModels = cfg.CachedAlibabaLlmModels is { Count: > 0 }
            ? cfg.CachedAlibabaLlmModels.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList()
            : null;
        _openAiFetchedTtsModels = cfg.CachedOpenAiTtsModels is { Count: > 0 }
            ? cfg.CachedOpenAiTtsModels.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList()
            : null;
        _alibabaFetchedTtsModels = cfg.CachedAlibabaTtsModels is { Count: > 0 }
            ? cfg.CachedAlibabaTtsModels.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList()
            : null;
        LocalDeviceCombo.Text = NormalizePreferDevice(cfg.PreferDevice);
        DefaultOutputDirTextBox.Text = cfg.DefaultOutputDir;
        SelectRuntimePresetFromConfig(cfg);
        BackendModeCombo.Text = SelectedPreset().StartsWith("api_", StringComparison.OrdinalIgnoreCase) ? "api" : "local";
        RefreshLocalBackendUi(cfg);
        ModelRepoTextBox.Text = cfg.ModelRepoId;
        AdditionalRepoTextBox.Text = cfg.AdditionalModelRepoId;
        ModelCacheTextBox.Text = cfg.ModelCacheDir;
        AutoDownloadModelCheckBox.IsChecked = cfg.AutoDownloadModel;
        AutoRemoveInputFilesCheckBox.IsChecked = cfg.AutoRemoveCompletedInputFiles;
        GenerateSrtSubtitlesCheckBox.IsChecked = cfg.GenerateSrtSubtitles;
        GenerateAssSubtitlesCheckBox.IsChecked = cfg.GenerateAssSubtitles;
        ApiKeyAlibabaPasswordBox.Password = string.Empty;
        ApiKeyOpenAiPasswordBox.Password = string.Empty;
        _removeAlibabaApiKeyRequested = false;
        _removeOpenAiApiKeyRequested = false;
        UpdateApiKeyStatusUi(cfg.ApiKeyAlibaba, cfg.ApiKeyOpenAi);
        InitializeApiTtsModelSelectors(cfg);
        SetComboSelectionByText(LlmProviderCombo, string.IsNullOrWhiteSpace(cfg.LlmPrepProvider) ? "local" : cfg.LlmPrepProvider, "local");
        LlmSeparateModelsCheckBox.IsChecked = cfg.LlmPrepUseSeparateModels;
        ApplyLlmProviderModelOptions(GetSelectedComboText(LlmProviderCombo, "local"), cfg.LlmPrepSplitModel, cfg.LlmPrepInstructionModel);
        LlmRuntimePathTextBox.Text = cfg.LlmLocalRuntimePath;
        LlmModelDirTextBox.Text = cfg.LlmLocalModelDir;
        LlmSplitTempTextBox.Text = cfg.LlmPrepTemperatureSplit.ToString("0.##", CultureInfo.InvariantCulture);
        LlmInstructionTempTextBox.Text = cfg.LlmPrepTemperatureInstruction.ToString("0.##", CultureInfo.InvariantCulture);
        LlmSplitMaxTokensTextBox.Text = cfg.LlmPrepMaxTokensSplit.ToString(CultureInfo.InvariantCulture);
        LlmInstructionMaxTokensTextBox.Text = cfg.LlmPrepMaxTokensInstruction.ToString(CultureInfo.InvariantCulture);
        LlmAdvancedToggleCheckBox.IsChecked = true;
        UpdateLlmAdvancedVisibility();
        SetComboSelectionByTag(AudioEnhanceProviderCombo, string.IsNullOrWhiteSpace(cfg.AudioEnhanceProvider) ? "local_audiox" : cfg.AudioEnhanceProvider, "local_audiox");
        SetComboSelectionByTag(AudioEnhanceVariantCombo, string.IsNullOrWhiteSpace(cfg.AudioEnhanceVariant) ? "audiox_base" : cfg.AudioEnhanceVariant, "audiox_base");
        AudioEnhanceAutoRunCheckBox.IsChecked = cfg.AudioEnhanceAutoRun;
        AudioEnhanceLlmDetectCheckBox.IsChecked = cfg.AudioEnhanceUseLlmCueDetection;
        AudioEnhanceLlmRefineCheckBox.IsChecked = cfg.AudioEnhanceUseLlmCueRefine;
        AudioEnhanceExportNarrationOnlyCheckBox.IsChecked = cfg.AudioEnhanceExportNarrationOnly;
        AudioEnhanceExportStemsCheckBox.IsChecked = cfg.AudioEnhanceExportStems;
        AudioEnhanceAmbienceDbTextBox.Text = cfg.AudioEnhanceAmbienceDb.ToString("0.##", CultureInfo.InvariantCulture);
        AudioEnhanceOneShotDbTextBox.Text = cfg.AudioEnhanceOneShotDb.ToString("0.##", CultureInfo.InvariantCulture);
        AudioEnhanceDuckingDbTextBox.Text = cfg.AudioEnhanceDuckingDb.ToString("0.##", CultureInfo.InvariantCulture);
        AudioEnhanceCueMaxPerMinuteTextBox.Text = cfg.AudioEnhanceCueMaxPerMinute.ToString(CultureInfo.InvariantCulture);
        AudioEnhanceRuntimePathTextBox.Text = cfg.AudioEnhanceRuntimePath ?? string.Empty;
        AudioEnhanceModelDirTextBox.Text = string.IsNullOrWhiteSpace(cfg.AudioEnhanceModelDir) ? "models/audiox" : cfg.AudioEnhanceModelDir;
        AudioEnhanceAdvancedToggleCheckBox.IsChecked = true;
        UpdateAudioEnhanceAdvancedVisibility();
        SelectQwenVariantFromConfig(cfg);
        ApplyPresetFieldVisibility(SelectedPreset());
        UpdateLlmRuntimeStatus(cfg);
        UpdateAudioEnhanceStatus(cfg);
        UpdateLlmFetchButtonState();
        RefreshModelInstallChecklist();
        SetRuntimeStatus("Ready.", isError: false);
        LlmInstallProgressBar.Value = 0;
        LlmInstallMetricsTextBlock.Text = "Idle";
        LlmInstallCurrentFileTextBlock.Text = "No active install";
        AudioEnhanceInstallProgressBar.Value = 0;
        AudioEnhanceInstallMetricsTextBlock.Text = "Idle";
        AudioEnhanceInstallCurrentFileTextBlock.Text = "No active install";
        DownloadMetricsTextBlock.Text = "Idle";
        DownloadCurrentFileTextBlock.Text = "No active download";
    }

    private async void SettingsWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        await AutoFetchApiTtsModelsForActivePresetAsync();
    }

    private async Task AutoFetchApiTtsModelsForActivePresetAsync(bool showErrors = false)
    {
        if (_isAutoFetchingApiTtsModels)
        {
            return;
        }

        var preset = SelectedPreset();
        string provider;
        string apiKey;
        if (string.Equals(preset, "api_openai_default", StringComparison.OrdinalIgnoreCase))
        {
            provider = "openai";
            apiKey = ApiKeyOpenAiPasswordBox.Password.Trim();
        }
        else if (string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase))
        {
            provider = "alibaba";
            apiKey = ApiKeyAlibabaPasswordBox.Password.Trim();
        }
        else
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return;
        }

        _isAutoFetchingApiTtsModels = true;
        try
        {
            var models = await FetchTtsModelListAsync(provider, apiKey);
            if (models.Count == 0)
            {
                return;
            }

            if (string.Equals(provider, "openai", StringComparison.OrdinalIgnoreCase))
            {
                _openAiFetchedTtsModels = models;
                await PersistFetchedModelCachesAsync();
                var current = GetSelectedComboText(OpenAiTtsModelCombo, models[0]);
                SetComboOptions(OpenAiTtsModelCombo, models, current, models[0]);
                if (string.Equals(preset, "api_openai_default", StringComparison.OrdinalIgnoreCase))
                {
                    ConfigureVariantSelector("api_openai_default");
                }
            }
            else
            {
                _alibabaFetchedTtsModels = models;
                await PersistFetchedModelCachesAsync();
                var current = GetSelectedComboText(AlibabaTtsModelCombo, models[0]);
                SetComboOptions(AlibabaTtsModelCombo, models, current, models[0]);
                if (string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase))
                {
                    ConfigureVariantSelector("api_alibaba");
                }
            }
        }
        catch (Exception ex)
        {
            if (showErrors)
            {
                MessageBox.Show(this,
                    UserMessageFormatter.FormatOperationError("API model auto-fetch", ex.Message),
                    "API TTS Models",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        finally
        {
            _isAutoFetchingApiTtsModels = false;
        }
    }

    private AppConfig ReadConfigFromUi()
    {
        var modelPreset = SelectedPreset();
        var modelRepo = ModelRepoTextBox.Text.Trim();
        var tokenizerRepo = AdditionalRepoTextBox.Text.Trim();
        var localModelBackend = SupportsPythonAlternativePreset(modelPreset)
            ? SelectedLocalBackendChoice()
            : "onnx";
        if (string.Equals(modelPreset, "qwen3_tts", StringComparison.OrdinalIgnoreCase))
        {
            var variant = ResolveEffectiveQwenVariant(SelectedQwenVariantKey());
            modelRepo = variant.RepoId;
            tokenizerRepo = variant.TokenizerRepo ?? string.Empty;
        }
        else if (SupportsPythonAlternativePreset(modelPreset))
        {
            modelRepo = ResolveRepoForPresetAndBackend(modelPreset, localModelBackend);
            tokenizerRepo = string.Empty;
        }

        var isApiPreset = modelPreset.StartsWith("api_", StringComparison.OrdinalIgnoreCase);
        var apiPreset = isApiPreset
            ? (string.Equals(modelPreset, "api_alibaba", StringComparison.OrdinalIgnoreCase)
                ? SelectedAlibabaApiPreset()
                : ApiPresetFromTag(modelPreset))
            : _original.ApiPreset;

        var apiProvider = isApiPreset ? string.Empty : _original.ApiProvider;
        var apiModelId = isApiPreset ? ResolveApiModelIdForPreset(modelPreset) : _original.ApiModelId;
        var apiBaseUrl = isApiPreset ? string.Empty : _original.ApiBaseUrl;
        var apiLanguageType = isApiPreset ? string.Empty : _original.ApiLanguageType;
        var apiVoiceDesignTargetModel = isApiPreset ? string.Empty : _original.ApiVoiceDesignTargetModel;

        var llmSplitModel = GetSelectedComboText(LlmSplitModelCombo, "qwen2.5-7b-instruct");
        var llmInstructionModel = GetSelectedComboText(LlmInstructionModelCombo, "qwen2.5-14b-instruct");
        var llmSeparate = LlmSeparateModelsCheckBox.IsChecked == true;

        var existingAlibabaKey = _removeAlibabaApiKeyRequested ? string.Empty : (_original.ApiKeyAlibaba ?? string.Empty).Trim();
        var existingOpenAiKey = _removeOpenAiApiKeyRequested ? string.Empty : (_original.ApiKeyOpenAi ?? string.Empty).Trim();
        var updatedAlibabaKey = ResolveApiKeyFromUi(existingAlibabaKey, ApiKeyAlibabaPasswordBox.Password);
        var updatedOpenAiKey = ResolveApiKeyFromUi(existingOpenAiKey, ApiKeyOpenAiPasswordBox.Password);

        return new AppConfig
        {
            DefaultOutputDir = DefaultOutputDirTextBox.Text.Trim(),
            ModelCacheDir = ModelCacheTextBox.Text.Trim(),
            PreferDevice = NormalizePreferDevice(LocalDeviceCombo.Text),
            OfflineMode = false,
            BackendMode = isApiPreset ? "api" : "local",
            LocalModelPreset = modelPreset,
            LocalModelBackend = localModelBackend,
            ModelRepoId = modelRepo,
            AdditionalModelRepoId = tokenizerRepo,
            AutoDownloadModel = AutoDownloadModelCheckBox.IsChecked == true,
            AutoRemoveCompletedInputFiles = AutoRemoveInputFilesCheckBox.IsChecked == true,
            GenerateSrtSubtitles = GenerateSrtSubtitlesCheckBox.IsChecked == true,
            GenerateAssSubtitles = GenerateAssSubtitlesCheckBox.IsChecked == true,
            ApiPreset = apiPreset,
            ApiProvider = apiProvider,
            ApiKey = _original.ApiKey,
            ApiKeyAlibaba = updatedAlibabaKey,
            ApiKeyOpenAi = updatedOpenAiKey,
            ApiModelId = apiModelId,
            ApiVoice = _original.ApiVoice,
            ApiBaseUrl = apiBaseUrl,
            ApiLanguageType = apiLanguageType,
            ApiVoiceDesignTargetModel = apiVoiceDesignTargetModel,
            CachedOpenAiLlmModels = (_openAiFetchedLlmModels ?? _original.CachedOpenAiLlmModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            CachedAlibabaLlmModels = (_alibabaFetchedLlmModels ?? _original.CachedAlibabaLlmModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            CachedOpenAiTtsModels = (_openAiFetchedTtsModels ?? _original.CachedOpenAiTtsModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            CachedAlibabaTtsModels = (_alibabaFetchedTtsModels ?? _original.CachedAlibabaTtsModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            AudioEnhanceEnabledByDefault = false,
            AudioEnhanceAutoRun = AudioEnhanceAutoRunCheckBox.IsChecked != false,
            AudioEnhanceMode = "sfx_ambience",
            AudioEnhanceProvider = GetSelectedComboTag(AudioEnhanceProviderCombo, "local_audiox"),
            AudioEnhanceVariant = GetSelectedComboTag(AudioEnhanceVariantCombo, "audiox_base"),
            AudioEnhanceRuntimePath = (AudioEnhanceRuntimePathTextBox.Text ?? string.Empty).Trim(),
            AudioEnhanceModelDir = (AudioEnhanceModelDirTextBox.Text ?? "models/audiox").Trim(),
            AudioEnhanceModelRepoId = string.IsNullOrWhiteSpace(_original.AudioEnhanceModelRepoId) ? "HKUSTAudio/AudioX" : _original.AudioEnhanceModelRepoId,
            AudioEnhanceUseLlmCueDetection = AudioEnhanceLlmDetectCheckBox.IsChecked == true,
            AudioEnhanceUseLlmCueRefine = AudioEnhanceLlmRefineCheckBox.IsChecked == true,
            AudioEnhanceAmbienceDb = Math.Clamp(ParseDoubleOrFallback(AudioEnhanceAmbienceDbTextBox.Text, -24.0), -40.0, 0.0),
            AudioEnhanceOneShotDb = Math.Clamp(ParseDoubleOrFallback(AudioEnhanceOneShotDbTextBox.Text, -15.0), -30.0, 0.0),
            AudioEnhanceDuckingDb = Math.Clamp(ParseDoubleOrFallback(AudioEnhanceDuckingDbTextBox.Text, -8.0), -24.0, 0.0),
            AudioEnhanceCueMaxPerMinute = Math.Clamp(ParseIntOrFallback(AudioEnhanceCueMaxPerMinuteTextBox.Text, 10), 1, 60),
            AudioEnhanceExportNarrationOnly = AudioEnhanceExportNarrationOnlyCheckBox.IsChecked != false,
            AudioEnhanceExportStems = AudioEnhanceExportStemsCheckBox.IsChecked == true,
            LlmPrepProvider = GetSelectedComboText(LlmProviderCombo, "local").ToLowerInvariant(),
            LlmPrepSplitModel = llmSplitModel,
            LlmPrepInstructionModel = llmInstructionModel,
            LlmPrepUseSeparateModels = llmSeparate,
            LlmLocalRuntimePath = (LlmRuntimePathTextBox.Text ?? string.Empty).Trim(),
            LlmLocalModelDir = (LlmModelDirTextBox.Text ?? "models/llm").Trim(),
            LlmLocalSplitModelFile = DeriveLlmModelFileName(llmSplitModel),
            LlmLocalInstructionModelFile = llmSeparate ? DeriveLlmModelFileName(llmInstructionModel) : DeriveLlmModelFileName(llmSplitModel),
            LlmPrepTemperatureSplit = Math.Clamp(ParseDoubleOrFallback(LlmSplitTempTextBox.Text, 0.2), 0.0, 1.5),
            LlmPrepTemperatureInstruction = Math.Clamp(ParseDoubleOrFallback(LlmInstructionTempTextBox.Text, 0.6), 0.0, 1.5),
            LlmPrepMaxTokensSplit = Math.Clamp(ParseIntOrFallback(LlmSplitMaxTokensTextBox.Text, 1024), 128, 8192),
            LlmPrepMaxTokensInstruction = Math.Clamp(ParseIntOrFallback(LlmInstructionMaxTokensTextBox.Text, 512), 64, 4096),
            ModelProfiles = CloneProfiles(_modelProfiles),
            LastOpenDir = _original.LastOpenDir,
            RecentProjects = new List<string>(_original.RecentProjects)
        };
    }

    private void UpdateLlmRuntimeStatus(AppConfig cfg)
    {
        var provider = (cfg.LlmPrepProvider ?? "local").Trim().ToLowerInvariant();
        var providerReady = provider switch
        {
            "openai" => !string.IsNullOrWhiteSpace(cfg.ApiKeyOpenAi),
            "alibaba" => !string.IsNullOrWhiteSpace(cfg.ApiKeyAlibaba),
            _ => true
        };

        LlmProviderReadyCheckBox.IsChecked = providerReady;
        LlmProviderReadyCheckBox.Content = provider switch
        {
            "openai" => "Provider/API key ready (OpenAI key)",
            "alibaba" => "Provider/API key ready (Alibaba key)",
            _ => "Provider/API key ready (local provider)"
        };

        if (!string.Equals(provider, "local", StringComparison.OrdinalIgnoreCase))
        {
            LlmRamReadyCheckBox.IsChecked = true;
            LlmRamReadyCheckBox.Content = "RAM requirement (N/A for cloud)";
            LlmVramReadyCheckBox.IsChecked = true;
            LlmVramReadyCheckBox.Content = "VRAM requirement (N/A for cloud)";
            LlmRuntimeReadyCheckBox.IsChecked = true;
            LlmRuntimeReadyCheckBox.Content = "Local runtime executable found (N/A for cloud)";
            LlmSplitModelReadyCheckBox.IsChecked = true;
            LlmSplitModelReadyCheckBox.Content = "Split model file found (cloud model)";
            LlmInstructionModelReadyCheckBox.IsChecked = true;
            LlmInstructionModelReadyCheckBox.Content = "Instruction model file found (cloud model)";
            LlmDiskReadyCheckBox.IsChecked = true;
            LlmDiskReadyCheckBox.Content = "Disk space looks sufficient for local GGUF files (N/A for cloud)";
            LlmRuntimeStatusTextBlock.Text = providerReady
                ? $"LLM runtime status: cloud provider '{provider}' selected and API key is ready."
                : $"LLM runtime status: cloud provider '{provider}' selected but API key is missing.";
            LlmRuntimeStatusTextBlock.Foreground = new SolidColorBrush(providerReady ? Color.FromRgb(79, 125, 99) : Color.FromRgb(170, 47, 47));
            return;
        }

        var root = RuntimePathResolver.AppRoot;
        var runtimePath = (cfg.LlmLocalRuntimePath ?? string.Empty).Trim();
        var runtimeExists = !string.IsNullOrWhiteSpace(runtimePath) && File.Exists(runtimePath);
        var modelDirInput = string.IsNullOrWhiteSpace(cfg.LlmLocalModelDir) ? "models/llm" : cfg.LlmLocalModelDir.Trim();
        var modelDir = Path.IsPathRooted(modelDirInput) ? modelDirInput : Path.Combine(root, modelDirInput);
        var splitModel = Path.Combine(modelDir, cfg.LlmLocalSplitModelFile ?? string.Empty);
        var instructionModel = Path.Combine(modelDir, cfg.LlmLocalInstructionModelFile ?? string.Empty);
        var splitExists = File.Exists(splitModel);
        var instructionExists = File.Exists(instructionModel);
        var requiredRamGb = EstimateLlmRequiredRamGb(cfg);
        var requiredVramGb = EstimateLlmRequiredVramGb(cfg);
        var ramReady = _systemProfile.RamGb >= requiredRamGb;
        var vramReady = _systemProfile.GpuVramGb >= requiredVramGb;
        var diskFreeGb = TryGetDiskFreeGb(modelDir);
        var requiredGb = EstimateLlmRequiredDiskGb(cfg);
        var diskReady = diskFreeGb <= 0 || diskFreeGb >= requiredGb;

        LlmRamReadyCheckBox.IsChecked = ramReady;
        LlmRamReadyCheckBox.Content = $"RAM requirement: {requiredRamGb:0.#}+ GB (Detected: {_systemProfile.RamGb:0.#} GB)";
        LlmVramReadyCheckBox.IsChecked = vramReady;
        LlmVramReadyCheckBox.Content = $"VRAM requirement: {requiredVramGb:0.#}+ GB (Detected: {_systemProfile.GpuVramGb:0.#} GB)";
        LlmRuntimeReadyCheckBox.IsChecked = runtimeExists;
        LlmRuntimeReadyCheckBox.Content = $"Local runtime executable found ({Path.GetFileName(runtimePath)})";
        LlmSplitModelReadyCheckBox.IsChecked = splitExists;
        LlmSplitModelReadyCheckBox.Content = $"Split model file found ({cfg.LlmLocalSplitModelFile})";
        LlmInstructionModelReadyCheckBox.IsChecked = instructionExists;
        LlmInstructionModelReadyCheckBox.Content = $"Instruction model file found ({cfg.LlmLocalInstructionModelFile})";
        LlmDiskReadyCheckBox.IsChecked = diskReady;
        LlmDiskReadyCheckBox.Content = diskFreeGb > 0
            ? $"Disk space looks sufficient for local GGUF files ({diskFreeGb:0.#} GB free, ~{requiredGb:0.#} GB needed)"
            : "Disk space looks sufficient for local GGUF files (unable to detect free space)";

        var allReady = ramReady && vramReady && runtimeExists && splitExists && instructionExists && diskReady;
        LlmRuntimeStatusTextBlock.Foreground = new SolidColorBrush(allReady ? Color.FromRgb(79, 125, 99) : Color.FromRgb(170, 47, 47));

        LlmRuntimeStatusTextBlock.Text = allReady
            ? "LLM runtime status: local setup is ready."
            : "LLM runtime status: local setup incomplete. Fix missing checklist items.";
    }

    private void HookLlmUiEvents()
    {
        if (_llmEventsHooked)
        {
            return;
        }

        _llmEventsHooked = true;
        LlmProviderCombo.SelectionChanged += LlmUiField_OnChanged;
        LlmSeparateModelsCheckBox.Checked += LlmUiField_OnChanged;
        LlmSeparateModelsCheckBox.Unchecked += LlmUiField_OnChanged;
        LlmSplitModelCombo.SelectionChanged += LlmUiField_OnChanged;
        LlmInstructionModelCombo.SelectionChanged += LlmUiField_OnChanged;
        LlmRuntimePathTextBox.TextChanged += LlmUiField_OnChanged;
        LlmModelDirTextBox.TextChanged += LlmUiField_OnChanged;
        LlmSplitTempTextBox.TextChanged += LlmUiField_OnChanged;
        LlmInstructionTempTextBox.TextChanged += LlmUiField_OnChanged;
        LlmSplitMaxTokensTextBox.TextChanged += LlmUiField_OnChanged;
        LlmInstructionMaxTokensTextBox.TextChanged += LlmUiField_OnChanged;
    }

    private void LlmUiField_OnChanged(object? sender, RoutedEventArgs e)
    {
        if (ReferenceEquals(sender, LlmProviderCombo))
        {
            ApplyLlmProviderModelOptions(
                GetSelectedComboText(LlmProviderCombo, "local"),
                GetSelectedComboText(LlmSplitModelCombo, "qwen2.5-7b-instruct"),
                GetSelectedComboText(LlmInstructionModelCombo, "qwen2.5-14b-instruct"));
        }
        UpdateLlmFetchButtonState();
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private void LlmAdvancedToggleCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        UpdateLlmAdvancedVisibility();
    }

    private void UpdateLlmAdvancedVisibility()
    {
        var show = LlmAdvancedToggleCheckBox.IsChecked == true;
        var vis = show ? Visibility.Visible : Visibility.Collapsed;
        if (LlmRuntimePathLabel is not null) LlmRuntimePathLabel.Visibility = vis;
        if (LlmRuntimePathGrid is not null) LlmRuntimePathGrid.Visibility = vis;
        if (LlmModelDirLabel is not null) LlmModelDirLabel.Visibility = vis;
        if (LlmModelDirGrid is not null) LlmModelDirGrid.Visibility = vis;
    }

    private void HookAudioEnhanceUiEvents()
    {
        if (_audioEnhanceEventsHooked)
        {
            return;
        }

        _audioEnhanceEventsHooked = true;
        AudioEnhanceProviderCombo.SelectionChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceVariantCombo.SelectionChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceAutoRunCheckBox.Checked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceAutoRunCheckBox.Unchecked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceLlmDetectCheckBox.Checked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceLlmDetectCheckBox.Unchecked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceLlmRefineCheckBox.Checked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceLlmRefineCheckBox.Unchecked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceExportNarrationOnlyCheckBox.Checked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceExportNarrationOnlyCheckBox.Unchecked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceExportStemsCheckBox.Checked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceExportStemsCheckBox.Unchecked += AudioEnhanceUiField_OnChanged;
        AudioEnhanceAmbienceDbTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceOneShotDbTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceDuckingDbTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceCueMaxPerMinuteTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceRuntimePathTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
        AudioEnhanceModelDirTextBox.TextChanged += AudioEnhanceUiField_OnChanged;
    }

    private void AudioEnhanceUiField_OnChanged(object? sender, RoutedEventArgs e)
    {
        UpdateAudioEnhanceStatus(ReadConfigFromUi());
    }

    private void AudioEnhanceAdvancedToggleCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        UpdateAudioEnhanceAdvancedVisibility();
    }

    private void UpdateAudioEnhanceAdvancedVisibility()
    {
        var show = AudioEnhanceAdvancedToggleCheckBox.IsChecked == true;
        var vis = show ? Visibility.Visible : Visibility.Collapsed;
        if (AudioEnhanceRuntimePathLabel is not null) AudioEnhanceRuntimePathLabel.Visibility = vis;
        if (AudioEnhanceRuntimePathGrid is not null) AudioEnhanceRuntimePathGrid.Visibility = vis;
        if (AudioEnhanceModelDirLabel is not null) AudioEnhanceModelDirLabel.Visibility = vis;
        if (AudioEnhanceModelDirGrid is not null) AudioEnhanceModelDirGrid.Visibility = vis;
    }

    private void BrowseAudioEnhanceRuntimePath_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Filter = "Python executable (python*.exe)|python*.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*",
            Title = "Select AudioX Python Runtime"
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        AudioEnhanceRuntimePathTextBox.Text = dialog.FileName;
        UpdateAudioEnhanceStatus(ReadConfigFromUi());
    }

    private void BrowseAudioEnhanceModelDir_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select AudioX model directory"
        };
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        AudioEnhanceModelDirTextBox.Text = dialog.SelectedPath;
        UpdateAudioEnhanceStatus(ReadConfigFromUi());
    }

    private void UpdateAudioEnhanceStatus(AppConfig cfg)
    {
        NormalizeAudioEnhanceSettings(cfg);
        var provider = (cfg.AudioEnhanceProvider ?? "local_audiox").Trim().ToLowerInvariant();
        var variant = (cfg.AudioEnhanceVariant ?? "audiox_base").Trim().ToLowerInvariant();
        var runtimePath = ResolveAudioEnhanceRuntimeExecutable(cfg);
        var runtimeReady = !string.IsNullOrWhiteSpace(runtimePath) &&
                           (string.Equals(runtimePath, "python", StringComparison.OrdinalIgnoreCase) || File.Exists(runtimePath));
        var modelDir = ResolveAudioEnhanceModelDir(cfg);
        var modelReady = HasAudioEnhanceModelAssets(modelDir);
        var requiredDiskGb = EstimateAudioEnhanceRequiredDiskGb(variant);
        var freeDiskGb = TryGetDiskFreeGb(modelDir);
        var diskReady = freeDiskGb <= 0 || freeDiskGb >= requiredDiskGb;
        var requiredRam = EstimateAudioEnhanceRequiredRamGb(variant);
        var requiredVram = EstimateAudioEnhanceRequiredVramGb(variant);

        AudioEnhanceRuntimeReadyCheckBox.IsChecked = runtimeReady;
        AudioEnhanceRuntimeReadyCheckBox.Content = $"Runtime ready ({Path.GetFileName(runtimePath)})";
        AudioEnhanceModelReadyCheckBox.IsChecked = modelReady;
        AudioEnhanceModelReadyCheckBox.Content = modelReady
            ? $"Model assets ready ({modelDir})"
            : $"Model assets ready (missing in {modelDir})";
        AudioEnhanceDiskReadyCheckBox.IsChecked = diskReady;
        AudioEnhanceDiskReadyCheckBox.Content = freeDiskGb > 0
            ? $"Disk requirement: ~{requiredDiskGb:0.#} GB needed ({freeDiskGb:0.#} GB free)"
            : $"Disk requirement: ~{requiredDiskGb:0.#} GB needed (free space unknown)";

        var allReady = runtimeReady && modelReady && diskReady;
        AudioEnhanceStatusTextBlock.Foreground = new SolidColorBrush(allReady ? Color.FromRgb(79, 125, 99) : Color.FromRgb(170, 47, 47));
        if (provider != "local_audiox")
        {
            AudioEnhanceStatusTextBlock.Text = $"Audio enhancement status: unsupported provider '{provider}'.";
            return;
        }

        AudioEnhanceStatusTextBlock.Text = allReady
            ? $"Audio enhancement status: ready ({variant}). Estimated requirement RAM {requiredRam:0.#}+ GB / VRAM {requiredVram:0.#}+ GB."
            : $"Audio enhancement status: setup incomplete ({variant}). Need RAM {requiredRam:0.#}+ GB / VRAM {requiredVram:0.#}+ GB.";
    }

    private async void InstallAudioEnhanceButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_audioEnhanceInstallCts is not null)
        {
            return;
        }

        var cfg = ReadConfigFromUi();
        NormalizeAudioEnhanceSettings(cfg);
        if (!string.Equals((cfg.AudioEnhanceProvider ?? "local_audiox").Trim(), "local_audiox", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "Only local AudioX provider is supported in v1.", "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var cts = new CancellationTokenSource();
        _audioEnhanceInstallCts = cts;
        try
        {
            InstallAudioEnhanceButton.IsEnabled = false;
            AudioEnhanceInstallProgressBar.IsIndeterminate = false;
            AudioEnhanceInstallProgressBar.Value = 0;
            AudioEnhanceInstallMetricsTextBlock.Text = "Preparing install...";
            AudioEnhanceInstallCurrentFileTextBlock.Text = "Checking runtime and model assets...";

            var runtime = ResolveAudioEnhanceRuntimeExecutable(cfg);
            if (string.IsNullOrWhiteSpace(runtime))
            {
                runtime = "python";
            }
            AudioEnhanceRuntimePathTextBox.Text = runtime;

            var modelDir = ResolveAudioEnhanceModelDir(cfg);
            Directory.CreateDirectory(modelDir);
            AudioEnhanceModelDirTextBox.Text = modelDir;

            var scriptPath = EnsureAudioEnhanceScriptFile();
            ReportAudioEnhanceInstallProgress("Checking Python runtime", runtime, 1, 5, 0, 0);
            var pyCheck = await RunProcessAsync(
                runtime,
                new[] { "-c", "import sys; print(sys.version)" },
                timeout: TimeSpan.FromSeconds(20),
                ct: cts.Token);
            if (pyCheck.ExitCode != 0)
            {
                throw new InvalidOperationException("Python runtime check failed: " + TrimToSingleLine(pyCheck.StdErr));
            }

            ReportAudioEnhanceInstallProgress("Installing dependencies", "torch/torchaudio/huggingface_hub/audiox", 2, 5, 0, 0);
            await EnsureAudioEnhancePythonModuleAsync(runtime, "torch", "torch", cts.Token);
            await EnsureAudioEnhancePythonModuleAsync(runtime, "torchaudio", "torchaudio", cts.Token);
            await EnsureAudioEnhancePythonModuleAsync(runtime, "huggingface_hub", "huggingface_hub", cts.Token);
            await EnsureAudioEnhancePythonModuleAsync(runtime, "audiox", "git+https://github.com/ZeyueT/AudioX.git", cts.Token);

            var repoCandidates = BuildAudioEnhanceRepoCandidates(cfg).ToList();
            var attemptedRepos = new List<string>();
            string? selectedRepo = null;
            string lastDownloadError = string.Empty;
            foreach (var repoCandidate in repoCandidates)
            {
                ReportAudioEnhanceInstallProgress("Downloading AudioX model", repoCandidate, 3, 5, 0, 0);
                var snapshotCode = await RunProcessAsync(
                    runtime,
                    new[]
                    {
                        "-c",
                        "from huggingface_hub import snapshot_download; " +
                        $"snapshot_download(repo_id={ToPythonQuoted(repoCandidate)}, local_dir={ToPythonQuoted(modelDir)}, local_dir_use_symlinks=False, resume_download=True); " +
                        "print('ok')"
                    },
                    timeout: TimeSpan.FromMinutes(90),
                    ct: cts.Token);
                if (snapshotCode.ExitCode == 0)
                {
                    selectedRepo = repoCandidate;
                    break;
                }

                attemptedRepos.Add(repoCandidate);
                var errorParts = new[] { TrimToSingleLine(snapshotCode.StdErr), TrimToSingleLine(snapshotCode.StdOut) }
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToArray();
                lastDownloadError = errorParts.Length > 0 ? string.Join(" | ", errorParts) : "unknown download error";
            }
            if (string.IsNullOrWhiteSpace(selectedRepo))
            {
                var attemptedText = attemptedRepos.Count > 0 ? string.Join(", ", attemptedRepos) : "(none)";
                throw new InvalidOperationException($"AudioX model download failed for repos [{attemptedText}]. Last error: {lastDownloadError}");
            }
            cfg.AudioEnhanceModelRepoId = selectedRepo;
            _original.AudioEnhanceModelRepoId = selectedRepo;

            ReportAudioEnhanceInstallProgress("Running smoke test", Path.GetFileName(scriptPath), 4, 5, 0, 0);
            var smokeOutput = Path.Combine(Path.GetTempPath(), $"audiox_smoke_{Guid.NewGuid():N}.wav");
            var smoke = await RunProcessAsync(
                runtime,
                new[]
                {
                    scriptPath,
                    "--model-dir", modelDir,
                    "--variant", cfg.AudioEnhanceVariant,
                    "--prompt", "soft room tone ambience",
                    "--seconds", "1.5",
                    "--output", smokeOutput
                },
                timeout: TimeSpan.FromMinutes(5),
                ct: cts.Token);
            if (smoke.ExitCode != 0 || !File.Exists(smokeOutput) || new FileInfo(smokeOutput).Length < 512)
            {
                throw new InvalidOperationException("AudioX smoke test failed: " + TrimToSingleLine(smoke.StdErr));
            }
            try
            {
                File.Delete(smokeOutput);
            }
            catch
            {
                // ignore
            }

            ReportAudioEnhanceInstallProgress("Install completed", modelDir, 5, 5, 0, 0);
            AudioEnhanceInstallProgressBar.Value = 100;
            AudioEnhanceInstallMetricsTextBlock.Text = "AudioX install completed.";
            AudioEnhanceInstallCurrentFileTextBlock.Text = "Runtime, dependencies, model, and smoke test verified.";
            UpdateAudioEnhanceStatus(ReadConfigFromUi());
            MessageBox.Show(this, "AudioX enhancement is ready to use.", "Audio Enhancement", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            AudioEnhanceInstallMetricsTextBlock.Text = "AudioX install canceled.";
            AudioEnhanceInstallCurrentFileTextBlock.Text = "Install canceled.";
        }
        catch (Exception ex)
        {
            AudioEnhanceInstallMetricsTextBlock.Text = "AudioX install failed.";
            AudioEnhanceInstallCurrentFileTextBlock.Text = ex.Message;
            UpdateAudioEnhanceStatus(ReadConfigFromUi());
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Audio Enhancement install", ex.Message),
                "Audio Enhancement",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            InstallAudioEnhanceButton.IsEnabled = true;
            _audioEnhanceInstallCts?.Dispose();
            _audioEnhanceInstallCts = null;
        }
    }

    private async void DownloadModelButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            SetRuntimeStatus("Checking model status...", isError: false);
            DownloadModelButton.IsEnabled = false;
            RefreshModelStatusButton.IsEnabled = false;
            DownloadProgressBar.Visibility = Visibility.Visible;
            DownloadProgressBar.IsIndeterminate = true;
            DownloadProgressBar.Value = 0;
            DownloadMetricsTextBlock.Text = "Preparing download...";
            DownloadCurrentFileTextBlock.Text = "Resolving file list...";
            var cfg = ReadConfigFromUi();
            ApplyPresetDefaults(cfg);
            NormalizeModelSettings(cfg);

            var installedBefore = IsSelectedPresetInstalled(cfg, out var beforeSummary);
            if (installedBefore)
            {
                SetRuntimeStatus($"Model already present. Verifying files... {beforeSummary}", isError: false);
            }
            else
            {
                SetRuntimeStatus("Downloading model files...", isError: false);
            }

            var progress = new Progress<string>(msg =>
            {
                StatusTextBlock.Text = msg;
                SetRuntimeStatus(msg, isError: false);
            });
            var telemetry = new Progress<ModelDownloadTelemetry>(OnDownloadTelemetry);
            await _downloader.DownloadAsync(cfg, progress, telemetry, CancellationToken.None, forceRedownload: false);
            if (string.Equals(cfg.LocalModelPreset, "chatterbox_onnx", StringComparison.OrdinalIgnoreCase) &&
                UsesPythonLocalBackend(cfg))
            {
                await EnsureChatterboxPythonBackendReadyAsync(cfg);
            }
            if (IsKittenRepoId(cfg.ModelRepoId) && !UsesPythonLocalBackend(cfg))
            {
                await EnsureKittenEspeakNgInstalledAsync();
            }
            if (string.Equals(cfg.LocalModelPreset, "kitten_tts", StringComparison.OrdinalIgnoreCase) &&
                UsesPythonLocalBackend(cfg))
            {
                await EnsureKittenPythonBackendReadyAsync(cfg);
            }
            if (string.Equals(cfg.LocalModelPreset, "qwen3_tts", StringComparison.OrdinalIgnoreCase))
            {
                await EnsureQwenPythonRuntimeReadyAsync();
            }
            var validationError = ValidateSelectedPresetFiles(cfg);
            if (!string.IsNullOrWhiteSpace(validationError))
            {
                throw new InvalidOperationException(validationError);
            }
            RefreshModelInstallChecklist();
            IsSelectedPresetInstalled(cfg, out var afterSummary);
            SetRuntimeStatus($"Model ready. {afterSummary}", isError: false);
            StatusTextBlock.Text = "Model download completed.";
            DownloadMetricsTextBlock.Text = "Download completed.";
            DownloadCurrentFileTextBlock.Text = "All files verified.";
            var resolvedModelDir = Path.Combine(ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot), "hf-cache");
            MessageBox.Show(
                this,
                $"Model files are ready in:\n{resolvedModelDir}",
                "Model Download",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            SetRuntimeStatus($"Model download failed: {ex.Message}", isError: true);
            StatusTextBlock.Text = $"Model download failed: {ex.Message}";
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Model download", ex.Message),
                "Model Download",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            DownloadModelButton.IsEnabled = true;
            RefreshModelStatusButton.IsEnabled = true;
            DownloadProgressBar.Visibility = Visibility.Collapsed;
            DownloadProgressBar.IsIndeterminate = true;
        }
    }

    private async Task EnsureKittenEspeakNgInstalledAsync()
    {
        if (TryFindEspeakNgExecutable(out _))
        {
            return;
        }

        var toolsDir = Path.Combine(RuntimePaths.AppRoot, "tools", "espeak-ng");
        Directory.CreateDirectory(toolsDir);
        var msiPath = Path.Combine(toolsDir, "espeak-ng.msi");

        using (var client = new HttpClient())
        using (var response = await client.GetAsync(EspeakNgMsiUrl, HttpCompletionOption.ResponseHeadersRead))
        {
            response.EnsureSuccessStatusCode();
            await using var fs = new FileStream(msiPath, FileMode.Create, FileAccess.Write, FileShare.None);
            await response.Content.CopyToAsync(fs);
        }

        SetRuntimeStatus("Installing eSpeak-NG (Kitten TTS requirement)...", isError: false);
        StatusTextBlock.Text = "Installing eSpeak-NG...";
        DownloadCurrentFileTextBlock.Text = "Launching eSpeak-NG installer (UAC may appear)...";

        var psi = new ProcessStartInfo
        {
            FileName = "msiexec.exe",
            UseShellExecute = true,
            Verb = "runas"
        };
        psi.ArgumentList.Add("/i");
        psi.ArgumentList.Add(msiPath);
        psi.ArgumentList.Add("/passive");

        using var proc = Process.Start(psi);
        if (proc is null)
        {
            throw new InvalidOperationException("Failed to start eSpeak-NG installer.");
        }
        await proc.WaitForExitAsync();
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException($"eSpeak-NG installer failed with exit code {proc.ExitCode}.");
        }

        if (!TryFindEspeakNgExecutable(out var installedPath))
        {
            throw new InvalidOperationException("eSpeak-NG installer completed, but espeak-ng.exe was not found. Restart the app and try again.");
        }

        DownloadCurrentFileTextBlock.Text = $"eSpeak-NG ready: {installedPath}";
    }

    private static bool TryFindEspeakNgExecutable(out string path)
    {
        var candidates = new[]
        {
            Path.Combine(RuntimePaths.AppRoot, "tools", "espeak-ng", "espeak-ng.exe"),
            Path.Combine(RuntimePaths.AppRoot, "tools", "espeak-ng", "command_line", "espeak-ng.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "eSpeak NG", "espeak-ng.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "eSpeak NG", "espeak-ng.exe")
        };

        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
            {
                path = candidate;
                return true;
            }
        }

        path = string.Empty;
        return false;
    }

    private void BrowseModelCache_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        ModelCacheTextBox.Text = dialog.SelectedPath;
        RefreshModelInstallChecklist();
    }

    private void BrowseDefaultOutputDir_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        DefaultOutputDirTextBox.Text = dialog.SelectedPath;
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        var cfg = ReadConfigFromUi();
        ApplyPresetDefaults(cfg);
        NormalizeModelSettings(cfg);
        NormalizeApiSettings(cfg);
        NormalizeAudioEnhanceSettings(cfg);
        if (string.IsNullOrWhiteSpace(cfg.ModelCacheDir))
        {
            cfg.ModelCacheDir = "models";
        }
        if (string.IsNullOrWhiteSpace(cfg.DefaultOutputDir))
        {
            cfg.DefaultOutputDir = "output";
        }
        if (string.IsNullOrWhiteSpace(cfg.ModelRepoId))
        {
            cfg.ModelRepoId = "onnx-community/chatterbox-ONNX";
        }
        if (string.IsNullOrWhiteSpace(cfg.PreferDevice))
        {
            cfg.PreferDevice = "auto";
        }
        cfg.OfflineMode = false;

        UpdatedConfig = cfg;
        DialogResult = true;
        Close();
    }

    private void CancelButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private static AppConfig Clone(AppConfig src)
    {
        return new AppConfig
        {
            DefaultOutputDir = src.DefaultOutputDir,
            ModelCacheDir = src.ModelCacheDir,
            PreferDevice = src.PreferDevice,
            OfflineMode = false,
            BackendMode = src.BackendMode,
            LocalModelPreset = src.LocalModelPreset,
            LocalModelBackend = src.LocalModelBackend,
            ModelRepoId = src.ModelRepoId,
            AdditionalModelRepoId = src.AdditionalModelRepoId,
            AutoDownloadModel = src.AutoDownloadModel,
            AutoRemoveCompletedInputFiles = src.AutoRemoveCompletedInputFiles,
            GenerateSrtSubtitles = src.GenerateSrtSubtitles,
            GenerateAssSubtitles = src.GenerateAssSubtitles,
            ApiPreset = src.ApiPreset,
            ApiProvider = src.ApiProvider,
            ApiKey = src.ApiKey,
            ApiKeyAlibaba = src.ApiKeyAlibaba,
            ApiKeyOpenAi = src.ApiKeyOpenAi,
            ApiModelId = src.ApiModelId,
            ApiVoice = src.ApiVoice,
            ApiBaseUrl = src.ApiBaseUrl,
            ApiLanguageType = src.ApiLanguageType,
            ApiVoiceDesignTargetModel = src.ApiVoiceDesignTargetModel,
            CachedOpenAiLlmModels = new List<string>(src.CachedOpenAiLlmModels ?? new List<string>()),
            CachedAlibabaLlmModels = new List<string>(src.CachedAlibabaLlmModels ?? new List<string>()),
            CachedOpenAiTtsModels = new List<string>(src.CachedOpenAiTtsModels ?? new List<string>()),
            CachedAlibabaTtsModels = new List<string>(src.CachedAlibabaTtsModels ?? new List<string>()),
            AudioEnhanceEnabledByDefault = src.AudioEnhanceEnabledByDefault,
            AudioEnhanceAutoRun = src.AudioEnhanceAutoRun,
            AudioEnhanceMode = src.AudioEnhanceMode,
            AudioEnhanceProvider = src.AudioEnhanceProvider,
            AudioEnhanceVariant = src.AudioEnhanceVariant,
            AudioEnhanceRuntimePath = src.AudioEnhanceRuntimePath,
            AudioEnhanceModelDir = src.AudioEnhanceModelDir,
            AudioEnhanceModelRepoId = src.AudioEnhanceModelRepoId,
            AudioEnhanceUseLlmCueDetection = src.AudioEnhanceUseLlmCueDetection,
            AudioEnhanceUseLlmCueRefine = src.AudioEnhanceUseLlmCueRefine,
            AudioEnhanceAmbienceDb = src.AudioEnhanceAmbienceDb,
            AudioEnhanceOneShotDb = src.AudioEnhanceOneShotDb,
            AudioEnhanceDuckingDb = src.AudioEnhanceDuckingDb,
            AudioEnhanceCueMaxPerMinute = src.AudioEnhanceCueMaxPerMinute,
            AudioEnhanceExportNarrationOnly = src.AudioEnhanceExportNarrationOnly,
            AudioEnhanceExportStems = src.AudioEnhanceExportStems,
            LlmPrepProvider = src.LlmPrepProvider,
            LlmPrepSplitModel = src.LlmPrepSplitModel,
            LlmPrepInstructionModel = src.LlmPrepInstructionModel,
            LlmPrepUseSeparateModels = src.LlmPrepUseSeparateModels,
            LlmLocalRuntimePath = src.LlmLocalRuntimePath,
            LlmLocalModelDir = src.LlmLocalModelDir,
            LlmLocalSplitModelFile = src.LlmLocalSplitModelFile,
            LlmLocalInstructionModelFile = src.LlmLocalInstructionModelFile,
            LlmPrepTemperatureSplit = src.LlmPrepTemperatureSplit,
            LlmPrepTemperatureInstruction = src.LlmPrepTemperatureInstruction,
            LlmPrepMaxTokensSplit = src.LlmPrepMaxTokensSplit,
            LlmPrepMaxTokensInstruction = src.LlmPrepMaxTokensInstruction,
            ModelProfiles = CloneProfiles(src.ModelProfiles),
            LastOpenDir = src.LastOpenDir,
            RecentProjects = new List<string>(src.RecentProjects)
        };
    }

    private void LocalModelPresetCombo_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        var preset = SelectedPreset();
        ApplyPresetFieldVisibility(preset);
        if (preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase))
        {
            BackendModeCombo.Text = "api";
            var apiPreset = string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase)
                ? SelectedAlibabaApiPreset()
                : ApiPresetFromTag(preset);
            var cfg = ReadConfigFromUi();
            cfg.ApiPreset = apiPreset;
            ApplyApiPresetDefaults(cfg);
            RefreshModelInstallChecklist();
            _ = AutoFetchApiTtsModelsForActivePresetAsync();
            return;
        }
        if (preset == "chatterbox_onnx")
        {
            SetSelectedLocalBackendChoice(ResolveRecommendedLocalBackend(preset));
            ApplyLocalBackendSelectionToRepo();
            AdditionalRepoTextBox.Text = string.Empty;
        }
        else if (preset == "qwen3_tts")
        {
            ApplySelectedQwenVariantToRepo();
        }
        else if (preset == "kitten_tts")
        {
            SetSelectedLocalBackendChoice(ResolveRecommendedLocalBackend(preset));
            ApplyLocalBackendSelectionToRepo();
            AdditionalRepoTextBox.Text = string.Empty;
        }

        RefreshModelInstallChecklist();
    }

    private void LocalModelBackendCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!SupportsPythonAlternativePreset(SelectedPreset()))
        {
            return;
        }

        ApplyLocalBackendSelectionToRepo();
        RefreshModelInstallChecklist();
    }

    private void ApiKeyPasswordBox_OnChanged(object sender, RoutedEventArgs e)
    {
        UpdateApiKeyStatusUi(
            ResolveApiKeyFromUi(_removeAlibabaApiKeyRequested ? string.Empty : _original.ApiKeyAlibaba, ApiKeyAlibabaPasswordBox.Password),
            ResolveApiKeyFromUi(_removeOpenAiApiKeyRequested ? string.Empty : _original.ApiKeyOpenAi, ApiKeyOpenAiPasswordBox.Password));
        var preset = SelectedPreset();
        if (preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase))
        {
            ApplyPresetFieldVisibility(preset);
            _ = AutoFetchApiTtsModelsForActivePresetAsync();
        }
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private void ClearAlibabaApiKeyButton_OnClick(object sender, RoutedEventArgs e)
    {
        _removeAlibabaApiKeyRequested = true;
        ApiKeyAlibabaPasswordBox.Password = string.Empty;
        UpdateApiKeyStatusUi(string.Empty, ResolveApiKeyFromUi(_removeOpenAiApiKeyRequested ? string.Empty : _original.ApiKeyOpenAi, ApiKeyOpenAiPasswordBox.Password));
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private void ClearOpenAiApiKeyButton_OnClick(object sender, RoutedEventArgs e)
    {
        _removeOpenAiApiKeyRequested = true;
        ApiKeyOpenAiPasswordBox.Password = string.Empty;
        UpdateApiKeyStatusUi(ResolveApiKeyFromUi(_removeAlibabaApiKeyRequested ? string.Empty : _original.ApiKeyAlibaba, ApiKeyAlibabaPasswordBox.Password), string.Empty);
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private void UpdateApiKeyStatusUi(string? alibabaKey, string? openAiKey)
    {
        var hasAlibaba = !string.IsNullOrWhiteSpace(alibabaKey);
        var hasOpenAi = !string.IsNullOrWhiteSpace(openAiKey);
        ApiKeyAlibabaStatusTextBlock.Text = hasAlibaba
            ? "Alibaba key configured. Enter a new key only if you want to replace it."
            : "Alibaba key not configured.";
        ApiKeyOpenAiStatusTextBlock.Text = hasOpenAi
            ? "OpenAI key configured. Enter a new key only if you want to replace it."
            : "OpenAI key not configured.";
        ClearAlibabaApiKeyButton.IsEnabled = hasAlibaba;
        ClearOpenAiApiKeyButton.IsEnabled = hasOpenAi;
    }

    private static string ResolveApiKeyFromUi(string? existingValue, string? enteredValue)
    {
        var trimmedEntered = (enteredValue ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(trimmedEntered))
        {
            return trimmedEntered;
        }

        return (existingValue ?? string.Empty).Trim();
    }

    private void BrowseLlmRuntimePath_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            Title = "Select llama.cpp runtime executable"
        };
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        LlmRuntimePathTextBox.Text = dialog.FileName;
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private void BrowseLlmModelDir_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        LlmModelDirTextBox.Text = dialog.SelectedPath;
        UpdateLlmRuntimeStatus(ReadConfigFromUi());
    }

    private async void InstallLlmNowButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_llmInstallCts is not null)
        {
            return;
        }

        var summary = new LlmInstallSummary();
        try
        {
            InstallLlmNowButton.IsEnabled = false;
            CancelLlmInstallButton.IsEnabled = true;
            _llmInstallCts = new CancellationTokenSource();
            var ct = _llmInstallCts.Token;
            LlmInstallProgressBar.IsIndeterminate = false;
            LlmInstallProgressBar.Value = 0;
            LlmInstallMetricsTextBlock.Text = "Preparing install...";
            LlmInstallCurrentFileTextBlock.Text = "Resolving required runtime/model files...";
            SetRuntimeStatus("Installing LLM runtime + model files...", isError: false);

            var cfg = ReadConfigFromUi();
            if (!string.Equals((cfg.LlmPrepProvider ?? "local").Trim(), "local", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("One-click installer is for local LLM provider. Set provider to 'local' first.");
            }

            var forceRedownload = LlmForceRedownloadCheckBox.IsChecked == true;
            var runtimeExe = await EnsureLlamaRuntimeInstalledAsync(forceRedownload, summary, ct);
            LlmRuntimePathTextBox.Text = runtimeExe;
            var baseCfg = ReadConfigFromUi();

            await EnsureLlmModelInstalledAsync(
                modelName: baseCfg.LlmPrepSplitModel,
                configuredFileName: string.IsNullOrWhiteSpace(baseCfg.LlmLocalSplitModelFile) ? "qwen2.5-7b-instruct.gguf" : baseCfg.LlmLocalSplitModelFile,
                forceRedownload: forceRedownload,
                summary: summary,
                ct: ct);

            if (baseCfg.LlmPrepUseSeparateModels)
            {
                await EnsureLlmModelInstalledAsync(
                    modelName: baseCfg.LlmPrepInstructionModel,
                    configuredFileName: string.IsNullOrWhiteSpace(baseCfg.LlmLocalInstructionModelFile) ? "qwen2.5-14b-instruct.gguf" : baseCfg.LlmLocalInstructionModelFile,
                    forceRedownload: forceRedownload,
                    summary: summary,
                    ct: ct);
            }

            ValidateLlmInstallResults(ReadConfigFromUi());
            UpdateLlmRuntimeStatus(ReadConfigFromUi());
            LlmInstallProgressBar.Value = 100;
            LlmInstallMetricsTextBlock.Text =
                $"Completed | downloaded {summary.DownloadedFiles}, skipped {summary.SkippedFiles}, verified {summary.VerifiedFiles}";
            LlmInstallCurrentFileTextBlock.Text = "All required LLM assets verified.";
            SetRuntimeStatus("LLM installer completed successfully.", isError: false);
            MessageBox.Show(
                this,
                $"LLM install finished.\nDownloaded: {summary.DownloadedFiles}\nSkipped: {summary.SkippedFiles}\nVerified: {summary.VerifiedFiles}",
                "LLM Installer",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            LlmInstallMetricsTextBlock.Text = "Install canceled.";
            LlmInstallCurrentFileTextBlock.Text = "Canceled by user.";
            SetRuntimeStatus("LLM install canceled.", isError: true);
        }
        catch (Exception ex)
        {
            LlmInstallMetricsTextBlock.Text = "Install failed.";
            LlmInstallCurrentFileTextBlock.Text = ex.Message;
            SetRuntimeStatus("LLM install failed: " + ex.Message, isError: true);
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("LLM install", ex.Message),
                "LLM Installer",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            _llmInstallCts?.Dispose();
            _llmInstallCts = null;
            InstallLlmNowButton.IsEnabled = true;
            CancelLlmInstallButton.IsEnabled = false;
        }
    }

    private void CancelLlmInstallButton_OnClick(object sender, RoutedEventArgs e)
    {
        _llmInstallCts?.Cancel();
    }

    private async Task<string> EnsureLlamaRuntimeInstalledAsync(bool forceRedownload, LlmInstallSummary summary, CancellationToken ct)
    {
        var cfg = ReadConfigFromUi();
        var configured = (cfg.LlmLocalRuntimePath ?? string.Empty).Trim();
        if (!forceRedownload && !string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
        {
            summary.SkippedFiles++;
            summary.VerifiedFiles++;
            ReportLlmInstallProgress("Runtime already present", configured, 1, 1, 0, null, 0);
            return configured;
        }

        var toolsDir = Path.Combine(RuntimePathResolver.AppRoot, "tools", "llama.cpp");
        Directory.CreateDirectory(toolsDir);
        var defaultExe = Path.Combine(toolsDir, "llama-cli.exe");
        if (!forceRedownload && File.Exists(defaultExe))
        {
            summary.SkippedFiles++;
            summary.VerifiedFiles++;
            ReportLlmInstallProgress("Runtime already present", defaultExe, 1, 1, 0, null, 0);
            return defaultExe;
        }

        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("AudiobookCreator/1.0");
        using var releaseRes = await http.GetAsync("https://api.github.com/repos/ggerganov/llama.cpp/releases/latest", ct);
        releaseRes.EnsureSuccessStatusCode();
        var releaseJson = await releaseRes.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(releaseJson);
        if (!doc.RootElement.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Could not read llama.cpp release assets.");
        }

        string? assetUrl = null;
        string? assetName = null;
        foreach (var asset in assets.EnumerateArray())
        {
            var name = asset.TryGetProperty("name", out var n) ? (n.GetString() ?? string.Empty) : string.Empty;
            if (name.Contains("win", StringComparison.OrdinalIgnoreCase) &&
                name.Contains("x64", StringComparison.OrdinalIgnoreCase) &&
                name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            {
                assetName = name;
                assetUrl = asset.TryGetProperty("browser_download_url", out var u) ? (u.GetString() ?? string.Empty) : string.Empty;
                if (name.Contains("avx2", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }
        }

        if (string.IsNullOrWhiteSpace(assetUrl))
        {
            throw new InvalidOperationException("No suitable llama.cpp Windows runtime zip found in latest release.");
        }

        var zipPath = Path.Combine(Path.GetTempPath(), $"llama_cpp_{Guid.NewGuid():N}.zip");
        await DownloadFileWithRetriesAsync(
            http,
            assetUrl,
            zipPath,
            stepCompleted: 1,
            stepTotal: 3,
            stepLabel: $"Downloading runtime ({assetName})",
            summary: summary,
            ct: ct);

        try
        {
            ZipFile.ExtractToDirectory(zipPath, toolsDir, overwriteFiles: true);
            summary.VerifiedFiles++;
            ReportLlmInstallProgress("Extracting runtime", assetName ?? "runtime", 2, 3, 0, null, 0);
        }
        finally
        {
            try { File.Delete(zipPath); } catch { }
        }

        var candidate = Directory.EnumerateFiles(toolsDir, "llama-cli.exe", SearchOption.AllDirectories).FirstOrDefault()
            ?? Directory.EnumerateFiles(toolsDir, "main.exe", SearchOption.AllDirectories).FirstOrDefault();
        if (string.IsNullOrWhiteSpace(candidate))
        {
            throw new InvalidOperationException("llama.cpp downloaded but executable not found (llama-cli.exe).");
        }

        summary.DownloadedFiles++;
        summary.VerifiedFiles++;
        ReportLlmInstallProgress("Runtime installed", Path.GetFileName(candidate), 3, 3, 0, null, 0);
        return candidate;
    }

    private async Task EnsureLlmModelInstalledAsync(
        string modelName,
        string configuredFileName,
        bool forceRedownload,
        LlmInstallSummary summary,
        CancellationToken ct)
    {
        var cfg = ReadConfigFromUi();
        var modelDirInput = string.IsNullOrWhiteSpace(cfg.LlmLocalModelDir) ? "models/llm" : cfg.LlmLocalModelDir.Trim();
        var modelDir = Path.IsPathRooted(modelDirInput)
            ? modelDirInput
            : Path.Combine(RuntimePathResolver.AppRoot, modelDirInput);
        Directory.CreateDirectory(modelDir);

        var finalFileName = NormalizeConfiguredGgufFileName(configuredFileName, modelName);
        var targetPath = Path.Combine(modelDir, finalFileName);
        if (!forceRedownload && File.Exists(targetPath) && new FileInfo(targetPath).Length > 1024 * 1024)
        {
            summary.SkippedFiles++;
            summary.VerifiedFiles++;
            ReportLlmInstallProgress("Model already present", finalFileName, 1, 1, 0, null, 0);
            return;
        }

        var spec = await ResolveLlmModelDownloadSpecAsync(modelName, finalFileName);
        var url = $"https://huggingface.co/{spec.RepoId}/resolve/main/{spec.FileName}";
        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("AudiobookCreator/1.0");
        var tmpPath = targetPath + ".download";
        await DownloadFileWithRetriesAsync(
            http,
            url,
            tmpPath,
            stepCompleted: 1,
            stepTotal: 2,
            stepLabel: $"Downloading model ({spec.RepoId}/{spec.FileName})",
            summary: summary,
            ct: ct);

        if (File.Exists(targetPath))
        {
            File.Delete(targetPath);
        }
        File.Move(tmpPath, targetPath);
        summary.DownloadedFiles++;
        summary.VerifiedFiles++;
        ReportLlmInstallProgress("Model installed", finalFileName, 2, 2, 0, null, 0);
    }

    private static string NormalizeConfiguredGgufFileName(string configuredFileName, string modelName)
    {
        var f = (configuredFileName ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(f) && f.EndsWith(".gguf", StringComparison.OrdinalIgnoreCase))
        {
            return f;
        }

        var safe = (modelName ?? "model").Trim().ToLowerInvariant().Replace('/', '-').Replace('\\', '-').Replace(' ', '-');
        return safe + ".gguf";
    }

    private async Task<(string RepoId, string FileName)> ResolveLlmModelDownloadSpecAsync(string modelName, string preferredFileName)
    {
        var normalized = (modelName ?? string.Empty).Trim();
        var lower = normalized.ToLowerInvariant();

        if (normalized.Contains('/', StringComparison.Ordinal) && normalized.Split('/').Length == 2)
        {
            var selected = await PickBestGgufFromRepoAsync(normalized, preferredFileName);
            return (normalized, selected);
        }

        if (lower.Contains("qwen2.5-7b", StringComparison.Ordinal))
        {
            var repo = "unsloth/Qwen2.5-7B-Instruct-GGUF";
            var file = await PickBestGgufFromRepoAsync(repo, preferredFileName, "Q4_K_M");
            return (repo, file);
        }
        if (lower.Contains("qwen2.5-14b", StringComparison.Ordinal))
        {
            var repo = "unsloth/Qwen2.5-14B-Instruct-GGUF";
            var file = await PickBestGgufFromRepoAsync(repo, preferredFileName, "Q4_K_M");
            return (repo, file);
        }
        if (lower.Contains("qwen3-14b", StringComparison.Ordinal))
        {
            var repo = "unsloth/Qwen3-14B-GGUF";
            var file = await PickBestGgufFromRepoAsync(repo, preferredFileName, "Q4_K_M");
            return (repo, file);
        }
        if (lower.Contains("qwen3-8b", StringComparison.Ordinal))
        {
            var repo = "unsloth/Qwen3-8B-GGUF";
            var file = await PickBestGgufFromRepoAsync(repo, preferredFileName, "Q4_K_M");
            return (repo, file);
        }

        // Fallback default for unknown local prep model names.
        {
            var repo = "unsloth/Qwen2.5-7B-Instruct-GGUF";
            var file = await PickBestGgufFromRepoAsync(repo, preferredFileName, "Q4_K_M");
            return (repo, file);
        }
    }

    private static async Task<string> PickBestGgufFromRepoAsync(string repoId, string preferredFileName, string preferredQuantHint = "Q4_K_M")
    {
        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("AudiobookCreator/1.0");
        using var res = await http.GetAsync($"https://huggingface.co/api/models/{repoId}");
        if (!res.IsSuccessStatusCode)
        {
            // fallback to configured filename if API metadata unavailable
            return preferredFileName;
        }

        var payload = await res.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(payload);
        if (!doc.RootElement.TryGetProperty("siblings", out var siblings) || siblings.ValueKind != JsonValueKind.Array)
        {
            return preferredFileName;
        }

        var files = siblings.EnumerateArray()
            .Select(x => x.TryGetProperty("rfilename", out var r) ? (r.GetString() ?? string.Empty) : string.Empty)
            .Where(x => x.EndsWith(".gguf", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (files.Count == 0)
        {
            return preferredFileName;
        }

        var exact = files.FirstOrDefault(x => string.Equals(x, preferredFileName, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(exact))
        {
            return exact;
        }

        var preferred = files.FirstOrDefault(x => x.Contains(preferredQuantHint, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(preferred))
        {
            return preferred;
        }

        return files[0];
    }

    private async Task DownloadFileWithRetriesAsync(
        HttpClient http,
        string url,
        string targetPath,
        int stepCompleted,
        int stepTotal,
        string stepLabel,
        LlmInstallSummary summary,
        CancellationToken ct)
    {
        const int maxAttempts = 3;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var res = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
                if (!res.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"HTTP {(int)res.StatusCode} {res.ReasonPhrase}");
                }

                var total = res.Content.Headers.ContentLength;
                await using var src = await res.Content.ReadAsStreamAsync(ct);
                await using var dst = File.Create(targetPath);
                var buffer = new byte[64 * 1024];
                long done = 0;
                var watch = Stopwatch.StartNew();
                var lastMs = 0L;

                while (true)
                {
                    var read = await src.ReadAsync(buffer, ct);
                    if (read <= 0)
                    {
                        break;
                    }
                    await dst.WriteAsync(buffer.AsMemory(0, read), ct);
                    done += read;
                    if (watch.ElapsedMilliseconds - lastMs >= 220)
                    {
                        var bps = watch.Elapsed.TotalSeconds > 0 ? done / watch.Elapsed.TotalSeconds : 0;
                        ReportLlmInstallProgress(stepLabel, Path.GetFileName(targetPath), stepCompleted, stepTotal, done, total, bps);
                        lastMs = watch.ElapsedMilliseconds;
                    }
                }

                ReportLlmInstallProgress(stepLabel, Path.GetFileName(targetPath), stepCompleted, stepTotal, done, total, 0);
                return;
            }
            catch (Exception ex) when (attempt < maxAttempts && ex is not OperationCanceledException)
            {
                ReportLlmInstallProgress($"Retry {attempt}/{maxAttempts - 1} after error", ex.Message, stepCompleted, stepTotal, 0, null, 0);
                await Task.Delay(400 * attempt, ct);
            }
        }
    }

    private void ReportLlmInstallProgress(
        string stage,
        string fileLabel,
        int stepCompleted,
        int stepTotal,
        long bytesDone,
        long? bytesTotal,
        double bytesPerSecond)
    {
        var percent = bytesTotal.HasValue && bytesTotal.Value > 0
            ? Math.Clamp((double)bytesDone / bytesTotal.Value * 100.0, 0.0, 100.0)
            : 0.0;
        LlmInstallProgressBar.IsIndeterminate = !bytesTotal.HasValue || bytesTotal.Value <= 0;
        if (!LlmInstallProgressBar.IsIndeterminate)
        {
            LlmInstallProgressBar.Value = percent;
        }
        var totalText = bytesTotal.HasValue && bytesTotal.Value > 0 ? FormatBytes(bytesTotal.Value) : "?";
        var speedText = bytesPerSecond > 0 ? $"{FormatBytes((long)bytesPerSecond)}/s" : "-";
        LlmInstallMetricsTextBlock.Text = $"{stage} | Step {stepCompleted}/{stepTotal} | {FormatBytes(bytesDone)} / {totalText} | {speedText}";
        LlmInstallCurrentFileTextBlock.Text = fileLabel;
    }

    private void ValidateLlmInstallResults(AppConfig cfg)
    {
        var provider = (cfg.LlmPrepProvider ?? "local").Trim().ToLowerInvariant();
        if (!string.Equals(provider, "local", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(cfg.LlmLocalRuntimePath) || !File.Exists(cfg.LlmLocalRuntimePath))
        {
            throw new InvalidOperationException("LLM runtime install verification failed (runtime executable missing).");
        }

        var modelDirInput = string.IsNullOrWhiteSpace(cfg.LlmLocalModelDir) ? "models/llm" : cfg.LlmLocalModelDir.Trim();
        var modelDir = Path.IsPathRooted(modelDirInput) ? modelDirInput : Path.Combine(RuntimePathResolver.AppRoot, modelDirInput);
        var splitPath = Path.Combine(modelDir, cfg.LlmLocalSplitModelFile ?? string.Empty);
        if (!File.Exists(splitPath))
        {
            throw new InvalidOperationException("LLM split model verification failed (file missing).");
        }

        if (cfg.LlmPrepUseSeparateModels)
        {
            var instructionPath = Path.Combine(modelDir, cfg.LlmLocalInstructionModelFile ?? string.Empty);
            if (!File.Exists(instructionPath))
            {
                throw new InvalidOperationException("LLM instruction model verification failed (file missing).");
            }
        }
    }

    private sealed class LlmInstallSummary
    {
        public int DownloadedFiles { get; set; }
        public int SkippedFiles { get; set; }
        public int VerifiedFiles { get; set; }
    }

    private static double TryGetDiskFreeGb(string path)
    {
        try
        {
            var full = Path.GetFullPath(path);
            var root = Path.GetPathRoot(full);
            if (string.IsNullOrWhiteSpace(root))
            {
                return -1;
            }
            var drive = new DriveInfo(root);
            return drive.AvailableFreeSpace / (1024d * 1024d * 1024d);
        }
        catch
        {
            return -1;
        }
    }

    private static double EstimateAudioEnhanceRequiredDiskGb(string variant)
    {
        return variant switch
        {
            "audiox_maf_mmdit" => 24.0,
            "audiox_maf" => 16.0,
            _ => 12.0
        };
    }

    private static double EstimateAudioEnhanceRequiredRamGb(string variant)
    {
        return variant switch
        {
            "audiox_maf_mmdit" => 32.0,
            "audiox_maf" => 24.0,
            _ => 16.0
        };
    }

    private static double EstimateAudioEnhanceRequiredVramGb(string variant)
    {
        return variant switch
        {
            "audiox_maf_mmdit" => 16.0,
            "audiox_maf" => 10.0,
            _ => 8.0
        };
    }

    private static string ResolveAudioEnhanceRuntimeExecutable(AppConfig cfg)
    {
        var configured = (cfg.AudioEnhanceRuntimePath ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
        {
            return configured;
        }

        var bundled = Path.Combine(RuntimePaths.AppRoot, "python_qwen", "Scripts", "python.exe");
        if (File.Exists(bundled))
        {
            return bundled;
        }

        return string.IsNullOrWhiteSpace(configured) ? string.Empty : configured;
    }

    private static string ResolveAudioEnhanceModelDir(AppConfig cfg)
    {
        var raw = string.IsNullOrWhiteSpace(cfg.AudioEnhanceModelDir) ? "models/audiox" : cfg.AudioEnhanceModelDir.Trim();
        return Path.IsPathRooted(raw) ? raw : Path.Combine(RuntimePaths.AppRoot, raw);
    }

    private static bool HasAudioEnhanceModelAssets(string modelDir)
    {
        if (string.IsNullOrWhiteSpace(modelDir) || !Directory.Exists(modelDir))
        {
            return false;
        }

        static bool MatchesAsset(string p)
        {
            var name = Path.GetFileName(p).ToLowerInvariant();
            return name.EndsWith(".safetensors", StringComparison.Ordinal) ||
                   name.EndsWith(".ckpt", StringComparison.Ordinal) ||
                   name.EndsWith(".pt", StringComparison.Ordinal) ||
                   name.EndsWith(".pth", StringComparison.Ordinal) ||
                   name.EndsWith(".bin", StringComparison.Ordinal);
        }

        return Directory.EnumerateFiles(modelDir, "*", SearchOption.AllDirectories).Any(MatchesAsset);
    }

    private async Task EnsureAudioEnhancePythonModuleAsync(string pythonExe, string moduleName, string installSpec, CancellationToken ct)
    {
        var check = await RunProcessAsync(
            pythonExe,
            new[] { "-c", $"import {moduleName}; print('ok')" },
            timeout: TimeSpan.FromSeconds(40),
            ct: ct);
        if (check.ExitCode == 0)
        {
            return;
        }

        var install = await RunProcessAsync(
            pythonExe,
            new[] { "-m", "pip", "install", "--upgrade", installSpec },
            timeout: TimeSpan.FromMinutes(20),
            ct: ct);
        if (install.ExitCode != 0)
        {
            throw new InvalidOperationException($"Failed installing python package '{installSpec}': {TrimToSingleLine(install.StdErr)}");
        }
    }

    private static async Task<(int ExitCode, string StdOut, string StdErr)> RunProcessAsync(
        string fileName,
        IReadOnlyList<string> args,
        TimeSpan timeout,
        CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }
        psi.Environment["PYTHONIOENCODING"] = "utf-8";

        using var proc = new Process { StartInfo = psi };
        if (!proc.Start())
        {
            throw new InvalidOperationException($"Failed to start process: {fileName}");
        }

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(timeout);

        try
        {
            await proc.WaitForExitAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            try
            {
                if (!proc.HasExited)
                {
                    proc.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // ignore
            }
            throw new TimeoutException($"Process timed out after {timeout.TotalSeconds:0} seconds: {fileName}");
        }

        var stdout = (await stdoutTask).Trim();
        var stderr = (await stderrTask).Trim();
        return (proc.ExitCode, stdout, stderr);
    }

    private void ReportAudioEnhanceInstallProgress(string stage, string current, int step, int stepTotal, long bytesDone, long bytesTotal)
    {
        var percent = stepTotal > 0 ? Math.Clamp(step / (double)stepTotal * 100.0, 0.0, 100.0) : 0.0;
        AudioEnhanceInstallProgressBar.IsIndeterminate = bytesTotal > 0 && bytesDone <= 0;
        if (!AudioEnhanceInstallProgressBar.IsIndeterminate)
        {
            AudioEnhanceInstallProgressBar.Value = percent;
        }

        var bytesText = bytesTotal > 0 ? $"{FormatBytes(bytesDone)} / {FormatBytes(bytesTotal)}" : "-";
        AudioEnhanceInstallMetricsTextBlock.Text = $"{stage} | Step {step}/{stepTotal} | {bytesText}";
        AudioEnhanceInstallCurrentFileTextBlock.Text = current;
    }

    private static string TrimToSingleLine(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }
        var t = text.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return t.Length <= 260 ? t : t[..260] + "...";
    }

    private static string ToPythonQuoted(string value)
    {
        var s = (value ?? string.Empty)
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("'", "\\'", StringComparison.Ordinal);
        return $"'{s}'";
    }

    private static string EnsureAudioEnhanceScriptFile()
    {
        var scriptDir = Path.Combine(RuntimePaths.AppRoot, "tools", "audiox");
        Directory.CreateDirectory(scriptDir);
        var scriptPath = Path.Combine(scriptDir, "audiox_generate.py");
        var shouldWrite = true;
        if (File.Exists(scriptPath))
        {
            var existing = File.ReadAllText(scriptPath);
            shouldWrite = !string.Equals(existing, AudioEnhanceGenerateScript, StringComparison.Ordinal);
        }

        if (shouldWrite)
        {
            File.WriteAllText(scriptPath, AudioEnhanceGenerateScript, new System.Text.UTF8Encoding(false));
        }

        return scriptPath;
    }

    private static double EstimateLlmRequiredDiskGb(AppConfig cfg)
    {
        static double EstimateForName(string? n)
        {
            var name = (n ?? string.Empty).ToLowerInvariant();
            if (name.Contains("14b", StringComparison.Ordinal)) return 11.0;
            if (name.Contains("8b", StringComparison.Ordinal)) return 6.5;
            if (name.Contains("7b", StringComparison.Ordinal)) return 5.8;
            if (name.Contains("3b", StringComparison.Ordinal)) return 2.8;
            if (name.Contains("1.5b", StringComparison.Ordinal) || name.Contains("1b", StringComparison.Ordinal)) return 1.8;
            return 4.0;
        }

        var split = EstimateForName(cfg.LlmLocalSplitModelFile);
        var instr = cfg.LlmPrepUseSeparateModels ? EstimateForName(cfg.LlmLocalInstructionModelFile) : 0.0;
        return Math.Max(3.0, split + instr + 1.0);
    }

    private static string DeriveLlmModelFileName(string modelName)
    {
        var name = (modelName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            return "qwen2.5-7b-instruct.gguf";
        }

        var safe = name.ToLowerInvariant()
            .Replace("/", "-", StringComparison.Ordinal)
            .Replace("\\", "-", StringComparison.Ordinal)
            .Replace(" ", "-", StringComparison.Ordinal);
        return safe + ".gguf";
    }

    private static double EstimateLlmRequiredRamGb(AppConfig cfg)
    {
        static double EstimateForName(string? n)
        {
            var name = (n ?? string.Empty).ToLowerInvariant();
            if (name.Contains("14b", StringComparison.Ordinal)) return 32.0;
            if (name.Contains("8b", StringComparison.Ordinal)) return 20.0;
            if (name.Contains("7b", StringComparison.Ordinal)) return 18.0;
            if (name.Contains("3b", StringComparison.Ordinal)) return 12.0;
            return 16.0;
        }

        var split = EstimateForName(cfg.LlmLocalSplitModelFile);
        var instr = cfg.LlmPrepUseSeparateModels ? EstimateForName(cfg.LlmLocalInstructionModelFile) : 0.0;
        return Math.Max(split, instr);
    }

    private static double EstimateLlmRequiredVramGb(AppConfig cfg)
    {
        static double EstimateForName(string? n)
        {
            var name = (n ?? string.Empty).ToLowerInvariant();
            if (name.Contains("14b", StringComparison.Ordinal)) return 10.0;
            if (name.Contains("8b", StringComparison.Ordinal)) return 8.0;
            if (name.Contains("7b", StringComparison.Ordinal)) return 6.0;
            if (name.Contains("3b", StringComparison.Ordinal)) return 4.0;
            return 6.0;
        }

        var split = EstimateForName(cfg.LlmLocalSplitModelFile);
        var instr = cfg.LlmPrepUseSeparateModels ? EstimateForName(cfg.LlmLocalInstructionModelFile) : 0.0;
        return Math.Max(split, instr);
    }

    private void ApplyLlmProviderModelOptions(string? providerRaw, string? desiredSplit, string? desiredInstruction)
    {
        var provider = (providerRaw ?? "local").Trim().ToLowerInvariant();
        var options = provider switch
        {
            "openai" => _openAiFetchedLlmModels is { Count: > 0 } ? _openAiFetchedLlmModels : OpenAiLlmModelOptions,
            "alibaba" => _alibabaFetchedLlmModels is { Count: > 0 } ? _alibabaFetchedLlmModels : AlibabaLlmModelOptions,
            _ => LocalLlmModelOptions
        };

        if (options.Count == 0)
        {
            return;
        }

        SetComboOptions(LlmSplitModelCombo, options, desiredSplit, options[0]);
        var instructionFallback = options.Count > 1 ? options[1] : options[0];
        SetComboOptions(LlmInstructionModelCombo, options, desiredInstruction, instructionFallback);
    }

    private static void SetComboOptions(ComboBox combo, IReadOnlyList<string> options, string? desired, string fallback)
    {
        var target = string.IsNullOrWhiteSpace(desired) ? fallback : desired.Trim();
        if (!options.Contains(target, StringComparer.OrdinalIgnoreCase))
        {
            target = fallback;
        }

        combo.Items.Clear();
        foreach (var o in options)
        {
            combo.Items.Add(new ComboBoxItem { Content = o });
        }

        var selected = combo.Items.OfType<ComboBoxItem>()
            .FirstOrDefault(i => string.Equals(i.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase));
        combo.SelectedItem = selected ?? combo.Items.OfType<ComboBoxItem>().FirstOrDefault();
    }

    private static string GetSelectedComboText(ComboBox combo, string fallback)
    {
        if (combo.SelectedItem is ComboBoxItem item &&
            item.Content is string value &&
            !string.IsNullOrWhiteSpace(value))
        {
            return value.Trim();
        }

        var text = (combo.Text ?? string.Empty).Trim();
        return string.IsNullOrWhiteSpace(text) ? fallback : text;
    }

    private static void SetComboSelectionByText(ComboBox combo, string? value, string fallback)
    {
        var target = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        var selected = combo.Items
            .OfType<ComboBoxItem>()
            .FirstOrDefault(i => string.Equals(i.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase));
        combo.SelectedItem = selected ?? combo.Items.OfType<ComboBoxItem>().FirstOrDefault();
    }

    private static string GetSelectedComboTag(ComboBox combo, string fallback)
    {
        if (combo.SelectedItem is ComboBoxItem item)
        {
            var tag = (item.Tag?.ToString() ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(tag))
            {
                return tag;
            }

            var content = (item.Content?.ToString() ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(content))
            {
                return content;
            }
        }

        return fallback;
    }

    private static void SetComboSelectionByTag(ComboBox combo, string? tagValue, string fallback)
    {
        var target = string.IsNullOrWhiteSpace(tagValue) ? fallback : tagValue.Trim();
        var selected = combo.Items
            .OfType<ComboBoxItem>()
            .FirstOrDefault(i => string.Equals(i.Tag?.ToString(), target, StringComparison.OrdinalIgnoreCase))
            ?? combo.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase));
        combo.SelectedItem = selected ?? combo.Items.OfType<ComboBoxItem>().FirstOrDefault();
    }

    private void InitializeApiTtsModelSelectors(AppConfig cfg)
    {
        var openAiOptions = GetOpenAiTtsModelOptions();
        var alibabaOptions = GetAlibabaTtsModelOptions();

        var openAiPreferred = string.Equals((cfg.ApiProvider ?? string.Empty).Trim(), "openai", StringComparison.OrdinalIgnoreCase)
            ? cfg.ApiModelId
            : OpenAiTtsModelOptions[0];
        var alibabaPreferred = string.Equals((cfg.ApiProvider ?? string.Empty).Trim(), "alibaba", StringComparison.OrdinalIgnoreCase)
            ? cfg.ApiModelId
            : AlibabaTtsModelOptions[0];

        SetComboOptions(OpenAiTtsModelCombo, openAiOptions, openAiPreferred, openAiOptions[0]);
        SetComboOptions(AlibabaTtsModelCombo, alibabaOptions, alibabaPreferred, alibabaOptions[0]);
    }

    private IReadOnlyList<string> GetOpenAiTtsModelOptions()
    {
        return _openAiFetchedTtsModels is { Count: > 0 } ? _openAiFetchedTtsModels : OpenAiTtsModelOptions;
    }

    private IReadOnlyList<string> GetAlibabaTtsModelOptions()
    {
        return _alibabaFetchedTtsModels is { Count: > 0 } ? _alibabaFetchedTtsModels : AlibabaTtsModelOptions;
    }

    private string ResolveApiModelIdForPreset(string modelPresetTag)
    {
        var tag = (modelPresetTag ?? string.Empty).Trim().ToLowerInvariant();
        if (tag == "api_openai_default")
        {
            return GetSelectedComboText(OpenAiTtsModelCombo, "gpt-4o-mini-tts");
        }

        if (tag == "api_alibaba")
        {
            return GetSelectedComboText(AlibabaTtsModelCombo, "qwen3-tts-flash");
        }

        return _original.ApiModelId;
    }

    private async void FetchOpenAiTtsModelsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var apiKey = ApiKeyOpenAiPasswordBox.Password.Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            MessageBox.Show(this, "OpenAI key is missing.", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            FetchOpenAiTtsModelsButton.IsEnabled = false;
            FetchOpenAiTtsModelsButton.Content = "Fetching...";
            var models = await FetchTtsModelListAsync("openai", apiKey);
            if (models.Count == 0)
            {
                MessageBox.Show(this, "No TTS-capable OpenAI models were returned for this key.", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _openAiFetchedTtsModels = models;
            await PersistFetchedModelCachesAsync();
            var current = GetSelectedComboText(OpenAiTtsModelCombo, models[0]);
            SetComboOptions(OpenAiTtsModelCombo, models, current, models[0]);
            if (string.Equals(SelectedPreset(), "api_openai_default", StringComparison.OrdinalIgnoreCase))
            {
                ConfigureVariantSelector("api_openai_default");
            }
            MessageBox.Show(this, $"Fetched {models.Count} OpenAI TTS model(s).", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("OpenAI model fetch", ex.Message),
                "API TTS Models",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            FetchOpenAiTtsModelsButton.IsEnabled = true;
            FetchOpenAiTtsModelsButton.Content = "Fetch";
        }
    }

    private async void FetchAlibabaTtsModelsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var apiKey = ApiKeyAlibabaPasswordBox.Password.Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            MessageBox.Show(this, "Alibaba key is missing.", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            FetchAlibabaTtsModelsButton.IsEnabled = false;
            FetchAlibabaTtsModelsButton.Content = "Fetching...";
            var models = await FetchTtsModelListAsync("alibaba", apiKey);
            if (models.Count == 0)
            {
                MessageBox.Show(this, "No TTS-capable Alibaba models were returned for this key.", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _alibabaFetchedTtsModels = models;
            await PersistFetchedModelCachesAsync();
            var current = GetSelectedComboText(AlibabaTtsModelCombo, models[0]);
            SetComboOptions(AlibabaTtsModelCombo, models, current, models[0]);
            if (string.Equals(SelectedPreset(), "api_alibaba", StringComparison.OrdinalIgnoreCase))
            {
                ConfigureVariantSelector("api_alibaba");
            }
            MessageBox.Show(this, $"Fetched {models.Count} Alibaba TTS model(s).", "API TTS Models", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Alibaba model fetch", ex.Message),
                "API TTS Models",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            FetchAlibabaTtsModelsButton.IsEnabled = true;
            FetchAlibabaTtsModelsButton.Content = "Fetch";
        }
    }

    private static async Task<IReadOnlyList<string>> FetchTtsModelListAsync(string provider, string apiKey)
    {
        var url = provider switch
        {
            "openai" => "https://api.openai.com/v1/models",
            "alibaba" => "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/models",
            _ => throw new InvalidOperationException("Unsupported provider.")
        };

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        using var res = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
        var payload = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            var shortPayload = payload.Length <= 300 ? payload : payload[..300] + "...";
            throw new InvalidOperationException($"{(int)res.StatusCode} {res.ReasonPhrase} | {shortPayload}");
        }

        using var doc = JsonDocument.Parse(payload);
        if (!doc.RootElement.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Response does not contain model list.");
        }

        var models = new List<string>();
        foreach (var item in data.EnumerateArray())
        {
            if (!item.TryGetProperty("id", out var idEl) || idEl.ValueKind != JsonValueKind.String)
            {
                continue;
            }

            var id = (idEl.GetString() ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            if (IsLikelyTtsModelId(provider, id))
            {
                models.Add(id);
            }
        }

        return models
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsLikelyTtsModelId(string provider, string modelId)
    {
        var id = (modelId ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(id))
        {
            return false;
        }

        if (provider == "openai")
        {
            if (id.Contains("audio-preview", StringComparison.OrdinalIgnoreCase) ||
                id.Contains("realtime", StringComparison.OrdinalIgnoreCase) ||
                id.Contains("transcribe", StringComparison.OrdinalIgnoreCase) ||
                id.Contains("whisper", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return id.Contains("tts", StringComparison.OrdinalIgnoreCase) ||
                   id.StartsWith("tts-", StringComparison.OrdinalIgnoreCase);
        }

        return provider switch
        {
            "alibaba" => id.Contains("tts", StringComparison.OrdinalIgnoreCase) || id.Contains("voice", StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }

    private void UpdateLlmFetchButtonState()
    {
        if (FetchLlmModelsButton is null)
        {
            return;
        }

        var provider = GetSelectedComboText(LlmProviderCombo, "local").ToLowerInvariant();
        var canFetch = provider == "openai" || provider == "alibaba";
        FetchLlmModelsButton.Visibility = canFetch ? Visibility.Visible : Visibility.Collapsed;
        FetchLlmModelsButton.Content = provider switch
        {
            "openai" => "Fetch OpenAI",
            "alibaba" => "Fetch Alibaba",
            _ => "Fetch Models"
        };
    }

    private async void FetchLlmModelsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var provider = GetSelectedComboText(LlmProviderCombo, "local").ToLowerInvariant();
        if (provider == "local")
        {
            MessageBox.Show(this, "Local provider uses local model files, not API model listing.", "LLM Models", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var apiKey = provider switch
        {
            "openai" => ApiKeyOpenAiPasswordBox.Password.Trim(),
            "alibaba" => ApiKeyAlibabaPasswordBox.Password.Trim(),
            _ => string.Empty
        };
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            MessageBox.Show(this, "API key is missing for the selected provider.", "LLM Models", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            FetchLlmModelsButton.IsEnabled = false;
            FetchLlmModelsButton.Content = "Fetching...";
            var models = await FetchLlmModelListAsync(provider, apiKey);
            if (models.Count == 0)
            {
                MessageBox.Show(this, "No compatible text models were returned by this API key.", "LLM Models", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (provider == "openai")
            {
                _openAiFetchedLlmModels = models;
            }
            else if (provider == "alibaba")
            {
                _alibabaFetchedLlmModels = models;
            }
            await PersistFetchedModelCachesAsync();

            var currentSplit = GetSelectedComboText(LlmSplitModelCombo, models[0]);
            var currentInstruction = GetSelectedComboText(LlmInstructionModelCombo, models[0]);
            ApplyLlmProviderModelOptions(provider, currentSplit, currentInstruction);
            UpdateLlmRuntimeStatus(ReadConfigFromUi());
            MessageBox.Show(this, $"Fetched {models.Count} models from {provider}.", "LLM Models", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("LLM model fetch", ex.Message),
                "LLM Models",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            FetchLlmModelsButton.IsEnabled = true;
            UpdateLlmFetchButtonState();
        }
    }

    private async Task PersistFetchedModelCachesAsync()
    {
        try
        {
            var store = new JsonConfigStore();
            var cfg = await store.LoadAsync();
            cfg.CachedOpenAiLlmModels = (_openAiFetchedLlmModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            cfg.CachedAlibabaLlmModels = (_alibabaFetchedLlmModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            cfg.CachedOpenAiTtsModels = (_openAiFetchedTtsModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            cfg.CachedAlibabaTtsModels = (_alibabaFetchedTtsModels ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            await store.SaveAsync(cfg);
        }
        catch
        {
            // Non-blocking cache persistence.
        }
    }

    private static async Task<IReadOnlyList<string>> FetchLlmModelListAsync(string provider, string apiKey)
    {
        var url = provider switch
        {
            "openai" => "https://api.openai.com/v1/models",
            "alibaba" => "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/models",
            _ => throw new InvalidOperationException("Unsupported provider for model fetch.")
        };

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        using var res = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
        var payload = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            var shortPayload = payload.Length <= 300 ? payload : payload[..300] + "...";
            throw new InvalidOperationException($"{(int)res.StatusCode} {res.ReasonPhrase} | {shortPayload}");
        }

        using var doc = JsonDocument.Parse(payload);
        if (!doc.RootElement.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Response does not contain a 'data' model list.");
        }

        var items = new List<string>();
        foreach (var model in data.EnumerateArray())
        {
            if (!model.TryGetProperty("id", out var idEl) || idEl.ValueKind != JsonValueKind.String)
            {
                continue;
            }

            var id = (idEl.GetString() ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            if (IsLikelyTextLlmModelId(provider, id))
            {
                items.Add(id);
            }
        }

        return items
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsLikelyTextLlmModelId(string provider, string modelId)
    {
        var id = (modelId ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(id))
        {
            return false;
        }

        if (id.Contains("audio", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("tts", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("transcribe", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("whisper", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("embedding", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("moderation", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("image", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("vision", StringComparison.OrdinalIgnoreCase) ||
            id.Contains("realtime", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return provider switch
        {
            "openai" => id.StartsWith("gpt-", StringComparison.OrdinalIgnoreCase) || id.StartsWith("o", StringComparison.OrdinalIgnoreCase),
            "alibaba" => id.Contains("qwen", StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }

    private void QwenVariantCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingQwenVariant)
        {
            return;
        }

        if (string.Equals(SelectedPreset(), "api_openai_default", StringComparison.OrdinalIgnoreCase))
        {
            if (QwenVariantCombo.SelectedItem is ComboBoxItem { Tag: string modelId } &&
                !string.IsNullOrWhiteSpace(modelId))
            {
                SetComboSelectionByText(OpenAiTtsModelCombo, modelId, OpenAiTtsModelOptions[0]);
            }
            RefreshModelInstallChecklist();
            return;
        }

        if (string.Equals(SelectedPreset(), "api_alibaba", StringComparison.OrdinalIgnoreCase))
        {
            RefreshModelInstallChecklist();
            return;
        }

        if (!string.Equals(SelectedPreset(), "qwen3_tts", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        ApplySelectedQwenVariantToRepo();
        RefreshModelInstallChecklist();
    }

    private void ApplySelectedQwenVariantToRepo()
    {
        var variant = ResolveEffectiveQwenVariant(SelectedQwenVariantKey());
        _isApplyingQwenVariant = true;
        try
        {
            ModelRepoTextBox.Text = variant.RepoId;
            AdditionalRepoTextBox.Text = variant.TokenizerRepo ?? string.Empty;
        }
        finally
        {
            _isApplyingQwenVariant = false;
        }
    }

    private void SelectRuntimePresetFromConfig(AppConfig cfg)
    {
        var backendMode = (cfg.BackendMode ?? "auto").Trim().ToLowerInvariant();
        var preset = backendMode == "api"
            ? BuildApiPresetTag(cfg.ApiPreset)
            : (cfg.LocalModelPreset ?? "chatterbox_onnx");

        foreach (var item in LocalModelPresetCombo.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), preset, StringComparison.OrdinalIgnoreCase))
            {
                LocalModelPresetCombo.SelectedItem = item;
                return;
            }
        }

        LocalModelPresetCombo.SelectedIndex = 0;
    }

    private string SelectedPreset()
    {
        if (LocalModelPresetCombo.SelectedItem is System.Windows.Controls.ComboBoxItem { Tag: string tag } && !string.IsNullOrWhiteSpace(tag))
        {
            return tag;
        }
        return "chatterbox_onnx";
    }

    private static string BuildApiPresetTag(string? apiPreset)
    {
        var key = (apiPreset ?? string.Empty).Trim().ToLowerInvariant();
        return key switch
        {
            "alibaba_qwen3_tts" => "api_alibaba",
            "alibaba_qwen3_tts_instruct" => "api_alibaba",
            "alibaba_qwen3_tts_vc" => "api_alibaba",
            "alibaba_voice_design" => "api_alibaba",
            "alibaba_qwen_tts_latest" => "api_alibaba",
            "openai_default" => "api_openai_default",
            _ => "api_openai_default"
        };
    }

    private static string ApiPresetFromTag(string presetTag)
    {
        var key = (presetTag ?? string.Empty).Trim().ToLowerInvariant();
        return key switch
        {
            "api_alibaba" => "alibaba_qwen3_tts",
            "api_openai_default" => "openai_default",
            _ => "openai_default"
        };
    }

    private string SelectedAlibabaApiPreset()
    {
        if (QwenVariantCombo.SelectedItem is ComboBoxItem { Tag: string tag } && !string.IsNullOrWhiteSpace(tag))
        {
            var match = AlibabaApiVariants.FirstOrDefault(x => string.Equals(x.Key, tag, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(match.ApiPreset))
            {
                return match.ApiPreset;
            }
        }

        return "alibaba_qwen3_tts";
    }

    private static string GuessApiPreset(AppConfig cfg)
    {
        var provider = (cfg.ApiProvider ?? string.Empty).Trim().ToLowerInvariant();
        var model = (cfg.ApiModelId ?? string.Empty).Trim().ToLowerInvariant();
        if (provider == "alibaba")
        {
            if (model.Contains("tts-instruct-flash", StringComparison.OrdinalIgnoreCase))
            {
                return "alibaba_qwen3_tts_instruct";
            }
            if (model.Contains("tts-vc", StringComparison.OrdinalIgnoreCase))
            {
                return "alibaba_qwen3_tts_vc";
            }
            if (model.StartsWith("qwen-tts", StringComparison.OrdinalIgnoreCase))
            {
                return "alibaba_qwen_tts_latest";
            }
            return model.Contains("voice-design", StringComparison.OrdinalIgnoreCase)
                ? "alibaba_voice_design"
                : "alibaba_qwen3_tts";
        }
        if (provider == "openai")
        {
            return "openai_default";
        }

        return "openai_default";
    }

    private void InitializeQwenVariantOptions()
    {
        QwenVariantCombo.Items.Clear();
        foreach (var option in QwenVariants)
        {
            QwenVariantCombo.Items.Add(new ComboBoxItem
            {
                Tag = option.Key,
                Content = option.DisplayName
            });
        }

        if (QwenVariantCombo.Items.Count > 0)
        {
            QwenVariantCombo.SelectedIndex = 0;
        }
    }

    private void SelectQwenVariantFromConfig(AppConfig cfg)
    {
        var repo = (cfg.ModelRepoId ?? string.Empty).Trim();
        var matched = QwenVariants.FirstOrDefault(x =>
            !string.IsNullOrWhiteSpace(x.RepoId) &&
            string.Equals(x.RepoId, repo, StringComparison.OrdinalIgnoreCase));
        var key = string.IsNullOrWhiteSpace(matched?.Key) ? "auto" : matched.Key;

        foreach (var item in QwenVariantCombo.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), key, StringComparison.OrdinalIgnoreCase))
            {
                QwenVariantCombo.SelectedItem = item;
                return;
            }
        }

        QwenVariantCombo.SelectedIndex = 0;
    }

    private string SelectedQwenVariantKey()
    {
        if (QwenVariantCombo.SelectedItem is ComboBoxItem { Tag: string key } && !string.IsNullOrWhiteSpace(key))
        {
            return key;
        }

        return "auto";
    }

    private QwenVariantOption ResolveEffectiveQwenVariant(string selectedKey)
    {
        var key = string.IsNullOrWhiteSpace(selectedKey) ? "auto" : selectedKey.Trim().ToLowerInvariant();
        if (key == "auto")
        {
            var recommendedKey = PickRecommendedQwenVariantKey(_systemProfile);
            var recommended = QwenVariants.FirstOrDefault(x => string.Equals(x.Key, recommendedKey, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(recommended?.Key))
            {
                return recommended;
            }
        }

        var direct = QwenVariants.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(direct?.Key))
        {
            return direct;
        }

        return QwenVariants.First(x => x.Key == "qwen_py_06b");
    }

    private static string PickRecommendedQwenVariantKey(SystemProfile profile)
    {
        if (profile.RamGb >= 24 && profile.GpuVramGb >= 10)
        {
            return "qwen_py_17b";
        }

        return "qwen_py_06b";
    }

    private void ApplyPresetFieldVisibility(string preset)
    {
        var isQwen = string.Equals(preset, "qwen3_tts", StringComparison.OrdinalIgnoreCase);
        var isOpenAiApi = string.Equals(preset, "api_openai_default", StringComparison.OrdinalIgnoreCase);
        var isAlibabaApi = string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase);
        var isApi = preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase);
        ConfigureVariantSelector(preset);

        ModelRepoLabel.Visibility = Visibility.Collapsed;
        ModelRepoTextBox.Visibility = Visibility.Collapsed;
        TokenizerRepoLabel.Visibility = Visibility.Collapsed;
        AdditionalRepoTextBox.Visibility = Visibility.Collapsed;

        if (QwenPythonRuntimePanel is not null)
        {
            QwenPythonRuntimePanel.Visibility = isQwen ? Visibility.Visible : Visibility.Collapsed;
        }

        if (isQwen)
        {
            UpdateQwenRuntimeNote();
        }

        var localVisibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        LocalModelFolderLabel.Visibility = localVisibility;
        if (LocalModelFolderGrid is not null)
        {
            LocalModelFolderGrid.Visibility = localVisibility;
        }
        ModelCacheTextBox.IsEnabled = !isApi;
        LocalDeviceLabel.Visibility = localVisibility;
        LocalDeviceCombo.Visibility = localVisibility;
        LocalDeviceCombo.IsEnabled = !isApi;
        AutoDownloadModelCheckBox.Visibility = localVisibility;
        AutoRemoveInputFilesCheckBox.Visibility = Visibility.Visible;
        GenerateSrtSubtitlesCheckBox.Visibility = Visibility.Visible;
        GenerateAssSubtitlesCheckBox.Visibility = Visibility.Visible;

        ModelControlsButton.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        if (!isAlibabaApi && !isOpenAiApi)
        {
            QwenVariantLabel.Visibility = isApi ? Visibility.Collapsed : QwenVariantLabel.Visibility;
            QwenVariantCombo.Visibility = isApi ? Visibility.Collapsed : QwenVariantCombo.Visibility;
        }

        DownloadModelButton.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        RefreshModelStatusButton.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        RemoveModelCacheButton.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        DownloadProgressBar.Visibility = isApi ? Visibility.Collapsed : DownloadProgressBar.Visibility;
        DownloadMetricsTextBlock.Visibility = isApi ? Visibility.Collapsed : DownloadMetricsTextBlock.Visibility;
        DownloadCurrentFileTextBlock.Visibility = isApi ? Visibility.Collapsed : DownloadCurrentFileTextBlock.Visibility;

        RecommendedLabel.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        ChecklistLabel.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        MainRepoInstalledCheckBox.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        TokenizerInstalledCheckBox.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        RamRequirementCheckBox.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        VramRequirementCheckBox.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;
        DiskRequirementCheckBox.Visibility = isApi ? Visibility.Collapsed : Visibility.Visible;

        if (isApi)
        {
            RecommendedTierTextBlock.Visibility = Visibility.Visible;
            QwenChunkLimitsTextBlock.Visibility = Visibility.Collapsed;
            RecommendedTierTextBlock.Text = BuildApiReadinessText(preset);
        }

        RefreshLocalBackendUi(ReadConfigFromUi());
    }

    private void ConfigureVariantSelector(string preset)
    {
        if (string.Equals(preset, "qwen3_tts", StringComparison.OrdinalIgnoreCase))
        {
            if (QwenVariantLabel.Text != "Qwen Variant" || QwenVariantCombo.Items.Count != QwenVariants.Count)
            {
                _isApplyingQwenVariant = true;
                try
                {
                    QwenVariantLabel.Text = "Qwen Variant";
                    InitializeQwenVariantOptions();
                    SelectQwenVariantFromConfig(ReadConfigFromUi());
                }
                finally
                {
                    _isApplyingQwenVariant = false;
                }
            }
            QwenVariantLabel.Visibility = Visibility.Visible;
            QwenVariantCombo.Visibility = Visibility.Visible;
            return;
        }

        if (string.Equals(preset, "api_openai_default", StringComparison.OrdinalIgnoreCase))
        {
            var options = GetOpenAiTtsModelOptions();
            var currentModel = GetSelectedComboText(OpenAiTtsModelCombo, options[0]);
            _isApplyingQwenVariant = true;
            try
            {
                QwenVariantLabel.Text = "OpenAI TTS Model";
                QwenVariantCombo.Items.Clear();
                foreach (var modelId in options)
                {
                    QwenVariantCombo.Items.Add(new ComboBoxItem
                    {
                        Tag = modelId,
                        Content = modelId
                    });
                }

                var selected = QwenVariantCombo.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(i => string.Equals(i.Tag?.ToString(), currentModel, StringComparison.OrdinalIgnoreCase));
                QwenVariantCombo.SelectedItem = selected ?? QwenVariantCombo.Items.OfType<ComboBoxItem>().FirstOrDefault();
            }
            finally
            {
                _isApplyingQwenVariant = false;
            }
            QwenVariantLabel.Visibility = Visibility.Visible;
            QwenVariantCombo.Visibility = Visibility.Visible;
            return;
        }

        if (string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase))
        {
            var currentPreset = GuessApiPreset(ReadConfigFromUi());
            _isApplyingQwenVariant = true;
            try
            {
                QwenVariantLabel.Text = "Alibaba Variant";
                QwenVariantCombo.Items.Clear();
                foreach (var variant in AlibabaApiVariants)
                {
                    QwenVariantCombo.Items.Add(new ComboBoxItem
                    {
                        Tag = variant.Key,
                        Content = variant.DisplayName
                    });
                }

                var preferredKey = AlibabaApiVariants.FirstOrDefault(x => x.ApiPreset == currentPreset).Key;
                if (string.IsNullOrWhiteSpace(preferredKey))
                {
                    preferredKey = AlibabaApiVariants[0].Key;
                }
                var selected = QwenVariantCombo.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(i => string.Equals(i.Tag?.ToString(), preferredKey, StringComparison.OrdinalIgnoreCase));
                QwenVariantCombo.SelectedItem = selected ?? QwenVariantCombo.Items.OfType<ComboBoxItem>().FirstOrDefault();
            }
            finally
            {
                _isApplyingQwenVariant = false;
            }
            QwenVariantLabel.Visibility = Visibility.Visible;
            QwenVariantCombo.Visibility = Visibility.Visible;
            return;
        }

        QwenVariantLabel.Visibility = Visibility.Collapsed;
        QwenVariantCombo.Visibility = Visibility.Collapsed;
    }

    private void ApiTtsModelCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingQwenVariant)
        {
            return;
        }

        var preset = SelectedPreset();
        if (string.Equals(preset, "api_openai_default", StringComparison.OrdinalIgnoreCase))
        {
            ConfigureVariantSelector("api_openai_default");
        }
        else if (string.Equals(preset, "api_alibaba", StringComparison.OrdinalIgnoreCase))
        {
            ConfigureVariantSelector("api_alibaba");
        }

        if (preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase))
        {
            RecommendedTierTextBlock.Visibility = Visibility.Visible;
            RecommendedTierTextBlock.Text = BuildApiReadinessText(preset);
        }
    }

    private string BuildApiReadinessText(string runtimePresetTag)
    {
        var apiPreset = ApiPresetFromTag(runtimePresetTag);
        var provider = apiPreset switch
        {
            "alibaba_qwen3_tts" => "Alibaba",
            "alibaba_qwen3_tts_instruct" => "Alibaba",
            "alibaba_qwen3_tts_vc" => "Alibaba",
            "alibaba_voice_design" => "Alibaba",
            "alibaba_qwen_tts_latest" => "Alibaba",
            "openai_default" => "OpenAI",
            _ => "OpenAI"
        };

        var hasKey = provider switch
        {
            "Alibaba" => !string.IsNullOrWhiteSpace(ApiKeyAlibabaPasswordBox.Password),
            "OpenAI" => !string.IsNullOrWhiteSpace(ApiKeyOpenAiPasswordBox.Password),
            _ => !string.IsNullOrWhiteSpace(ApiKeyOpenAiPasswordBox.Password)
        };

        return hasKey
            ? $"API readiness: {provider} key detected. Ready to use."
            : $"API readiness: {provider} key missing. Add key in API Settings.";
    }

    private void UpdateQwenRuntimeNote()
    {
        if (QwenPythonRuntimeNoteTextBlock is null)
        {
            return;
        }

        var appRoot = RuntimePathResolver.AppRoot;
        var path = TryFindBundledPythonExe(appRoot, "python_qwen") ?? Path.Combine(appRoot, "python_qwen", "Scripts", "python.exe");
        QwenPythonRuntimeNoteTextBlock.Text =
            "Qwen uses bundled Python runtime (isolated from system Python).\n" +
            $"Path: {path}";
    }

    private async void VerifyQwenRuntimeButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            VerifyQwenRuntimeButton.IsEnabled = false;
            QwenPythonRuntimeVerifyTextBlock.Text = "Checking bundled Qwen runtime...";
            QwenPythonRuntimeVerifyTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(79, 125, 99));

            var appRoot = RuntimePathResolver.AppRoot;
            var pythonExe = TryFindBundledPythonExe(appRoot, "python_qwen");
            if (string.IsNullOrWhiteSpace(pythonExe))
            {
                throw new FileNotFoundException("Bundled Qwen Python runtime not found (customer_release\\python_qwen).");
            }

            var probe = await ProbeQwenRuntimeAsync(pythonExe, appRoot);
            var torchVer = probe.TorchVersion;
            var cudaLine = probe.CudaLine;
            QwenPythonRuntimeVerifyTextBlock.Text = $"Bundled Qwen runtime OK | torch {torchVer} | {cudaLine}";
            QwenPythonRuntimeVerifyTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(79, 125, 99));

            if (!string.IsNullOrWhiteSpace(probe.Stderr))
            {
                // qwen-tts commonly prints optional warnings (flash-attn/sox); surface a short note without failing.
                SetRuntimeStatus("Qwen runtime verified (warnings from qwen-tts are optional).", isError: false);
            }
            else
            {
                SetRuntimeStatus("Qwen runtime verified.", isError: false);
            }
        }
        catch (Exception ex)
        {
            QwenPythonRuntimeVerifyTextBlock.Text = "Qwen runtime check failed: " + ex.Message;
            QwenPythonRuntimeVerifyTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(170, 47, 47));
            SetRuntimeStatus("Qwen runtime check failed: " + ex.Message, isError: true);
        }
        finally
        {
            VerifyQwenRuntimeButton.IsEnabled = true;
        }
    }

    private async Task EnsureQwenPythonRuntimeReadyAsync()
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var pythonExe = TryFindBundledPythonExe(appRoot, "python_qwen");
        if (string.IsNullOrWhiteSpace(pythonExe))
        {
            SetRuntimeStatus("Qwen runtime missing. Preparing bundled runtime...", isError: false);
            DownloadCurrentFileTextBlock.Text = "Installing bundled Qwen runtime...";
            pythonExe = await BootstrapBundledPythonRuntimeAsync(appRoot, "python_qwen");
        }

        DownloadCurrentFileTextBlock.Text = "Checking Qwen runtime dependencies...";
        (string TorchVersion, string CudaLine, string Stderr) probe;
        try
        {
            probe = await ProbeQwenRuntimeAsync(pythonExe, appRoot);
        }
        catch
        {
            SetRuntimeStatus("Qwen runtime found, installing missing dependencies...", isError: false);
            DownloadCurrentFileTextBlock.Text = "Installing missing Python packages (qwen-tts / torch / torchaudio)...";
            await EnsureQwenPythonDependenciesAsync(pythonExe, appRoot);
            probe = await ProbeQwenRuntimeAsync(pythonExe, appRoot);
        }

        QwenPythonRuntimeVerifyTextBlock.Text = $"Bundled Qwen runtime OK | torch {probe.TorchVersion} | {probe.CudaLine}";
        QwenPythonRuntimeVerifyTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(79, 125, 99));
        if (!string.IsNullOrWhiteSpace(probe.Stderr))
        {
            SetRuntimeStatus("Qwen runtime verified (warnings from qwen-tts are optional).", isError: false);
        }
    }

    private async Task<string> EnsureBundledPythonRuntimeReadyAsync(string envName, string displayName)
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var normalizedEnv = (envName ?? string.Empty).Trim().ToLowerInvariant();
        var pythonExe = TryFindBundledPythonExe(appRoot, normalizedEnv);
        if (!string.IsNullOrWhiteSpace(pythonExe))
        {
            return pythonExe;
        }

        SetRuntimeStatus($"{displayName} Python runtime missing. Preparing local runtime...", isError: false);
        DownloadCurrentFileTextBlock.Text = $"Installing bundled {displayName} runtime...";
        return await BootstrapBundledPythonRuntimeAsync(appRoot, normalizedEnv);
    }

    private async Task EnsureTorchRuntimeAsync(string pythonExe, string workingDir, bool preferCudaTorch)
    {
        var hasTorch = await TryRunProcessAsync(pythonExe, "-m pip show torch torchaudio", workingDir);
        if (hasTorch)
        {
            return;
        }

        await RunProcessOrThrowAsync(pythonExe, "-m pip install --upgrade pip", workingDir, "Failed to update pip for bundled Python runtime.");
        if (preferCudaTorch)
        {
            var installedCudaTorch = await TryRunProcessAsync(
                pythonExe,
                "-m pip install --index-url https://download.pytorch.org/whl/cu124 torch==2.6.0+cu124 torchaudio==2.6.0+cu124",
                workingDir);
            if (installedCudaTorch)
            {
                return;
            }
        }

        await RunProcessOrThrowAsync(
            pythonExe,
            "-m pip install torch==2.6.0 torchaudio==2.6.0",
            workingDir,
            "Failed to install torch/torchaudio for bundled Python runtime.");
    }

    private static async Task EnsurePythonModuleAsync(string pythonExe, string workingDir, string moduleName, string installSpec)
    {
        var hasModule = await TryRunProcessAsync(
            pythonExe,
            $"-c \"import {moduleName}; print('ok')\"",
            workingDir);
        if (hasModule)
        {
            return;
        }

        await RunProcessOrThrowAsync(
            pythonExe,
            $"-m pip install --upgrade {installSpec}",
            workingDir,
            $"Failed to install Python package for module '{moduleName}'.");
    }

    private async Task EnsureChatterboxPythonBackendReadyAsync(AppConfig cfg)
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var pythonExe = await EnsureBundledPythonRuntimeReadyAsync("python_chatterbox", "Chatterbox");
        var preferCudaTorch = string.Equals(_systemProfile.GpuVendor, "nvidia", StringComparison.OrdinalIgnoreCase) &&
                              !string.Equals(NormalizePreferDevice(cfg.PreferDevice), "cpu", StringComparison.OrdinalIgnoreCase);

        DownloadCurrentFileTextBlock.Text = "Installing Chatterbox Python dependencies...";
        await EnsureTorchRuntimeAsync(pythonExe, appRoot, preferCudaTorch);
        await EnsurePythonModuleAsync(pythonExe, appRoot, "chatterbox", "chatterbox-tts");
        await RunProcessOrThrowAsync(
            pythonExe,
            "-c \"from chatterbox.tts import ChatterboxTTS; print('ok')\"",
            appRoot,
            "Chatterbox Python backend verification failed.");
    }

    private async Task EnsureKittenPythonBackendReadyAsync(AppConfig cfg)
    {
        var appRoot = RuntimePathResolver.AppRoot;
        var pythonExe = await EnsureBundledPythonRuntimeReadyAsync("python_kitten", "Kitten");

        DownloadCurrentFileTextBlock.Text = "Installing Kitten Python dependencies...";
        await EnsurePythonModuleAsync(pythonExe, appRoot, "soundfile", "soundfile");
        await EnsurePythonModuleAsync(pythonExe, appRoot, "kittentts", KittenTtsWheelUrl);
        await RunProcessOrThrowAsync(
            pythonExe,
            "-c \"from kittentts import KittenTTS; print('ok')\"",
            appRoot,
            "Kitten Python backend verification failed.");
    }

    private static Task<string> BootstrapBundledPythonRuntimeAsync(string appRoot, string envName)
    {
        var targetDir = Path.Combine(appRoot, envName);
        var normalizedEnv = (envName ?? string.Empty).Trim().ToLowerInvariant();
        var baseEnvName = ResolvePythonBootstrapBaseName(normalizedEnv);
        var candidates = new List<string>
        {
            Path.Combine(appRoot, "tools", "python_backends", normalizedEnv),
            Path.Combine(appRoot, "tools", normalizedEnv),
            Path.Combine(appRoot, normalizedEnv)
        };

        if (string.Equals(normalizedEnv, "python_qwen", StringComparison.OrdinalIgnoreCase))
        {
            candidates.Add(Path.Combine(appRoot, "tools", "python_backends", "python_qwen"));
            candidates.Add(Path.Combine(appRoot, ".venv"));
            candidates.Add(Path.GetFullPath(Path.Combine(appRoot, "..", ".venv")));
        }
        else
        {
            candidates.Add(Path.Combine(appRoot, "tools", "python_backends", baseEnvName));
            candidates.Add(Path.Combine(appRoot, "tools", baseEnvName));
            candidates.Add(Path.Combine(appRoot, baseEnvName));
        }

        var sourceDir = candidates.FirstOrDefault(IsValidPythonRuntimeDir);
        if (string.IsNullOrWhiteSpace(sourceDir))
        {
            throw new FileNotFoundException(
                $"Python runtime package is missing for '{envName}'. Include a bundled runtime or a clean base runtime '{baseEnvName}'.");
        }

        if (Directory.Exists(targetDir))
        {
            Directory.Delete(targetDir, recursive: true);
        }

        CopyDirectory(sourceDir, targetDir);
        var pythonExe = TryFindBundledPythonExe(appRoot, normalizedEnv);
        if (string.IsNullOrWhiteSpace(pythonExe))
        {
            throw new FileNotFoundException($"Python runtime copy finished but python.exe was not found in {envName}.");
        }

        return Task.FromResult(pythonExe);
    }

    private static bool IsValidPythonRuntimeDir(string? dir)
    {
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
        {
            return false;
        }

        return File.Exists(Path.Combine(dir, "python.exe")) ||
               File.Exists(Path.Combine(dir, "Scripts", "python.exe"));
    }

    private static string? TryFindBundledPythonExe(string appRoot, string envName)
    {
        var candidates = new[]
        {
            Path.Combine(appRoot, envName, "python.exe"),
            Path.Combine(appRoot, envName, "Scripts", "python.exe"),
            Path.Combine(appRoot, "tools", envName, "python.exe"),
            Path.Combine(appRoot, "tools", envName, "Scripts", "python.exe")
        };

        return candidates.FirstOrDefault(File.Exists);
    }

    private static async Task<(string TorchVersion, string CudaLine, string Stderr)> ProbeQwenRuntimeAsync(string pythonExe, string workingDir)
    {
        var probeCode = "import torch, qwen_tts; print('OK'); print(torch.__version__); print('cuda=' + str(bool(torch.cuda.is_available())))";
        var psi = new ProcessStartInfo
        {
            FileName = pythonExe,
            Arguments = $"-c \"{probeCode}\"",
            WorkingDirectory = workingDir,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.Environment["PYTHONIOENCODING"] = "utf-8";

        using var proc = new Process { StartInfo = psi };
        if (!proc.Start())
        {
            throw new InvalidOperationException("Failed to start bundled Qwen Python runtime.");
        }

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync();
        var stdout = (await stdoutTask).Trim();
        var stderr = (await stderrTask).Trim();

        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr) ? $"Python runtime check failed (exit {proc.ExitCode})." : stderr);
        }

        var outLines = stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var torchVer = outLines.FirstOrDefault(x => x.StartsWith("2.", StringComparison.OrdinalIgnoreCase)) ?? "unknown";
        var cudaLine = outLines.FirstOrDefault(x => x.StartsWith("cuda=", StringComparison.OrdinalIgnoreCase)) ?? "cuda=unknown";
        return (torchVer, cudaLine, stderr);
    }

    private static async Task EnsureQwenPythonDependenciesAsync(string pythonExe, string workingDir)
    {
        var hasCorePackages = await TryRunProcessAsync(
            pythonExe,
            "-m pip show qwen-tts torch torchaudio",
            workingDir);
        if (hasCorePackages)
        {
            return;
        }

        await RunProcessOrThrowAsync(pythonExe, "-m pip install --upgrade pip", workingDir, "Failed to update pip for Qwen runtime.");
        await RunProcessOrThrowAsync(
            pythonExe,
            "-m pip install --upgrade qwen-tts==0.1.1",
            workingDir,
            "Failed to install qwen-tts package for Qwen runtime.");

        var hasTorch = await TryRunProcessAsync(pythonExe, "-m pip show torch torchaudio", workingDir);
        if (!hasTorch)
        {
            var installedCudaTorch = await TryRunProcessAsync(
                pythonExe,
                "-m pip install --index-url https://download.pytorch.org/whl/cu124 torch==2.6.0+cu124 torchaudio==2.6.0+cu124",
                workingDir);
            if (!installedCudaTorch)
            {
                await RunProcessOrThrowAsync(
                    pythonExe,
                    "-m pip install torch==2.6.0 torchaudio==2.6.0",
                    workingDir,
                    "Failed to install torch/torchaudio for Qwen runtime.");
            }
        }
    }

    private static async Task<bool> TryRunProcessAsync(string fileName, string arguments, string workingDir)
    {
        try
        {
            await RunProcessOrThrowAsync(fileName, arguments, workingDir, null);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static async Task RunProcessOrThrowAsync(string fileName, string arguments, string workingDir, string? errorPrefix)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = workingDir,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.Environment["PYTHONIOENCODING"] = "utf-8";

        using var proc = new Process { StartInfo = psi };
        if (!proc.Start())
        {
            throw new InvalidOperationException(errorPrefix ?? $"Failed to start process: {fileName} {arguments}");
        }

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync();
        var stdout = (await stdoutTask).Trim();
        var stderr = (await stderrTask).Trim();

        if (proc.ExitCode == 0)
        {
            return;
        }

        var detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        if (string.IsNullOrWhiteSpace(detail))
        {
            detail = $"Exit code {proc.ExitCode}.";
        }

        if (string.IsNullOrWhiteSpace(errorPrefix))
        {
            throw new InvalidOperationException(detail);
        }

        throw new InvalidOperationException($"{errorPrefix} {detail}");
    }

    private static void CopyDirectory(string sourceDir, string targetDir)
    {
        Directory.CreateDirectory(targetDir);
        foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.TopDirectoryOnly))
        {
            var destFile = Path.Combine(targetDir, Path.GetFileName(file));
            File.Copy(file, destFile, overwrite: true);
        }

        foreach (var dir in Directory.GetDirectories(sourceDir, "*", SearchOption.TopDirectoryOnly))
        {
            var destSubDir = Path.Combine(targetDir, Path.GetFileName(dir));
            CopyDirectory(dir, destSubDir);
        }
    }

    private static void ApplyPresetDefaults(AppConfig cfg)
    {
        var preset = (cfg.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        if (preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase))
        {
            cfg.BackendMode = "api";
            if (preset == "api_alibaba")
            {
                if (string.IsNullOrWhiteSpace(cfg.ApiPreset) || !cfg.ApiPreset.StartsWith("alibaba_", StringComparison.OrdinalIgnoreCase))
                {
                    cfg.ApiPreset = "alibaba_qwen3_tts";
                }
            }
            else
            {
                cfg.ApiPreset = ApiPresetFromTag(preset);
            }
            ApplyApiPresetDefaults(cfg);
            return;
        }
        if (preset == "qwen3_tts")
        {
            if (string.IsNullOrWhiteSpace(cfg.ModelRepoId))
            {
                cfg.ModelRepoId = "Qwen/Qwen3-TTS-12Hz-1.7B-Base";
            }
            cfg.AdditionalModelRepoId = string.Empty;
            return;
        }

        if (preset == "kitten_tts")
        {
            cfg.LocalModelBackend = NormalizeLocalModelBackendChoice(cfg.LocalModelBackend);
            if (string.IsNullOrWhiteSpace(cfg.ModelRepoId))
            {
                cfg.ModelRepoId = ResolveRepoForPresetAndBackend(preset, cfg.LocalModelBackend);
            }
            cfg.AdditionalModelRepoId = string.Empty;
            return;
        }

        if (preset == "chatterbox_onnx")
        {
            cfg.LocalModelBackend = NormalizeLocalModelBackendChoice(cfg.LocalModelBackend);
            if (string.IsNullOrWhiteSpace(cfg.ModelRepoId))
            {
                cfg.ModelRepoId = ResolveRepoForPresetAndBackend(preset, cfg.LocalModelBackend);
            }
            cfg.AdditionalModelRepoId = string.Empty;
        }

        ApplyApiPresetDefaults(cfg);
    }

    private static void NormalizeModelSettings(AppConfig cfg)
    {
        cfg.ModelCacheDir = ModelCachePath.NormalizeInput(cfg.ModelCacheDir);
        cfg.PreferDevice = NormalizePreferDevice(cfg.PreferDevice);
        var mode = (cfg.BackendMode ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(mode) || mode == "auto")
        {
            cfg.BackendMode = "local";
        }
        if (string.Equals(cfg.ModelRepoId?.Trim(), "Qwen/Qwen3-TTS-0.6B", StringComparison.OrdinalIgnoreCase))
        {
            cfg.ModelRepoId = "Qwen/Qwen3-TTS-12Hz-0.6B-Base";
        }
        if (string.Equals(cfg.ModelRepoId?.Trim(), "zukky/Qwen3-TTS-ONNX-DLL", StringComparison.OrdinalIgnoreCase))
        {
            cfg.ModelRepoId = "Qwen/Qwen3-TTS-12Hz-0.6B-Base";
            cfg.AdditionalModelRepoId = string.Empty;
        }
        if (string.Equals(cfg.ModelRepoId?.Trim(), "xkos/Qwen3-TTS-12Hz-1.7B-ONNX", StringComparison.OrdinalIgnoreCase))
        {
            cfg.ModelRepoId = "Qwen/Qwen3-TTS-12Hz-1.7B-Base";
            cfg.AdditionalModelRepoId = string.Empty;
        }
        if (string.Equals((cfg.LocalModelPreset ?? string.Empty).Trim(), "chatterbox_onnx", StringComparison.OrdinalIgnoreCase))
        {
            cfg.LocalModelBackend = (cfg.ModelRepoId ?? string.Empty).Contains("onnx", StringComparison.OrdinalIgnoreCase)
                ? "onnx"
                : "python";
            cfg.ModelRepoId = ResolveRepoForPresetAndBackend(cfg.LocalModelPreset ?? "chatterbox_onnx", cfg.LocalModelBackend);
            cfg.AdditionalModelRepoId = string.Empty;
        }
        else if (string.Equals((cfg.LocalModelPreset ?? string.Empty).Trim(), "kitten_tts", StringComparison.OrdinalIgnoreCase))
        {
            cfg.LocalModelBackend = NormalizeLocalModelBackendChoice(cfg.LocalModelBackend);
            cfg.ModelRepoId = ResolveRepoForPresetAndBackend(cfg.LocalModelPreset ?? "kitten_tts", cfg.LocalModelBackend);
            cfg.AdditionalModelRepoId = string.Empty;
        }
        else
        {
            cfg.LocalModelBackend = string.Empty;
        }

        NormalizeApiSettings(cfg);
        NormalizeAudioEnhanceSettings(cfg);
    }

    private static void NormalizeAudioEnhanceSettings(AppConfig cfg)
    {
        cfg.AudioEnhanceMode = string.IsNullOrWhiteSpace(cfg.AudioEnhanceMode) ? "sfx_ambience" : cfg.AudioEnhanceMode.Trim().ToLowerInvariant();
        cfg.AudioEnhanceProvider = string.IsNullOrWhiteSpace(cfg.AudioEnhanceProvider) ? "local_audiox" : cfg.AudioEnhanceProvider.Trim().ToLowerInvariant();
        cfg.AudioEnhanceVariant = string.IsNullOrWhiteSpace(cfg.AudioEnhanceVariant) ? "audiox_base" : cfg.AudioEnhanceVariant.Trim().ToLowerInvariant();
        if (!AudioEnhanceVariantOptions.Contains(cfg.AudioEnhanceVariant, StringComparer.OrdinalIgnoreCase))
        {
            cfg.AudioEnhanceVariant = "audiox_base";
        }
        cfg.AudioEnhanceModelDir = string.IsNullOrWhiteSpace(cfg.AudioEnhanceModelDir) ? "models/audiox" : cfg.AudioEnhanceModelDir.Trim();
        var repo = (cfg.AudioEnhanceModelRepoId ?? string.Empty).Trim();
        if (string.Equals(repo, "audiox", StringComparison.OrdinalIgnoreCase))
        {
            repo = string.Empty;
        }
        var defaultVariantRepo = GetAudioEnhanceDefaultRepoForVariant(cfg.AudioEnhanceVariant);
        if (string.IsNullOrWhiteSpace(repo) || string.Equals(repo, "ZeyueT/AudioX", StringComparison.OrdinalIgnoreCase))
        {
            repo = defaultVariantRepo;
        }
        cfg.AudioEnhanceModelRepoId = repo;
        cfg.AudioEnhanceAmbienceDb = Math.Clamp(cfg.AudioEnhanceAmbienceDb, -40.0, 0.0);
        cfg.AudioEnhanceOneShotDb = Math.Clamp(cfg.AudioEnhanceOneShotDb, -30.0, 0.0);
        cfg.AudioEnhanceDuckingDb = Math.Clamp(cfg.AudioEnhanceDuckingDb, -24.0, 0.0);
        cfg.AudioEnhanceCueMaxPerMinute = Math.Clamp(cfg.AudioEnhanceCueMaxPerMinute <= 0 ? 10 : cfg.AudioEnhanceCueMaxPerMinute, 1, 60);
    }

    private static string GetAudioEnhanceDefaultRepoForVariant(string? variant)
    {
        var variantKey = (variant ?? "audiox_base").Trim();
        if (AudioEnhanceVariantRepoCandidates.TryGetValue(variantKey, out var repos)
            && repos is { Length: > 0 })
        {
            return repos[0];
        }
        return "HKUSTAudio/AudioX";
    }

    private static IEnumerable<string> BuildAudioEnhanceRepoCandidates(AppConfig cfg)
    {
        var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        static string NormalizeRepoCandidate(string repo)
        {
            if (string.Equals(repo, "audiox", StringComparison.OrdinalIgnoreCase))
            {
                return "HKUSTAudio/AudioX";
            }
            return string.Equals(repo, "ZeyueT/AudioX", StringComparison.OrdinalIgnoreCase)
                ? "HKUSTAudio/AudioX"
                : repo;
        }

        var configured = NormalizeRepoCandidate((cfg.AudioEnhanceModelRepoId ?? string.Empty).Trim());
        if (!string.IsNullOrWhiteSpace(configured) && yielded.Add(configured))
        {
            yield return configured;
        }

        var variantKey = (cfg.AudioEnhanceVariant ?? "audiox_base").Trim();
        if (!AudioEnhanceVariantRepoCandidates.TryGetValue(variantKey, out var defaults)
            || defaults is not { Length: > 0 })
        {
            defaults = AudioEnhanceVariantRepoCandidates["audiox_base"];
        }

        foreach (var repo in defaults.Select(NormalizeRepoCandidate))
        {
            if (!string.IsNullOrWhiteSpace(repo) && yielded.Add(repo))
            {
                yield return repo;
            }
        }
    }

    private static void ApplyApiPresetDefaults(AppConfig cfg)
    {
        var preset = (cfg.ApiPreset ?? string.Empty).Trim().ToLowerInvariant();
        switch (preset)
        {
            case "alibaba_qwen3_tts":
                cfg.ApiProvider = "alibaba";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "qwen3-tts-flash";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoice))
                {
                    cfg.ApiVoice = "Cherry";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
                {
                    cfg.ApiBaseUrl = "https://dashscope-intl.aliyuncs.com/api/v1";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
                {
                    cfg.ApiLanguageType = "Auto";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
                {
                    cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
                }
                break;
            case "alibaba_qwen3_tts_instruct":
                cfg.ApiProvider = "alibaba";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "qwen3-tts-instruct-flash";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoice))
                {
                    cfg.ApiVoice = "Cherry";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
                {
                    cfg.ApiBaseUrl = "https://dashscope-intl.aliyuncs.com/api/v1";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
                {
                    cfg.ApiLanguageType = "Auto";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
                {
                    cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
                }
                break;
            case "alibaba_qwen3_tts_vc":
                cfg.ApiProvider = "alibaba";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "qwen3-tts-vc-2026-01-22";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
                {
                    cfg.ApiBaseUrl = "https://dashscope-intl.aliyuncs.com/api/v1";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
                {
                    cfg.ApiLanguageType = "Auto";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
                {
                    cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
                }
                break;
            case "alibaba_qwen_tts_latest":
                cfg.ApiProvider = "alibaba";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "qwen-tts-latest";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoice))
                {
                    cfg.ApiVoice = "Cherry";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
                {
                    cfg.ApiBaseUrl = "https://dashscope-intl.aliyuncs.com/api/v1";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
                {
                    cfg.ApiLanguageType = "Auto";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
                {
                    cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
                }
                break;
            case "alibaba_voice_design":
                cfg.ApiProvider = "alibaba";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "qwen3-tts-vd-2026-01-26";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoice))
                {
                    cfg.ApiVoice = "Cherry";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
                {
                    cfg.ApiBaseUrl = "https://dashscope-intl.aliyuncs.com/api/v1";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
                {
                    cfg.ApiLanguageType = "Auto";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
                {
                    cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
                }
                break;
            case "openai_default":
                cfg.ApiProvider = "openai";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "gpt-4o-mini-tts";
                }
                if (string.IsNullOrWhiteSpace(cfg.ApiVoice))
                {
                    cfg.ApiVoice = "alloy";
                }
                break;
            case "custom":
                break;
            default:
                cfg.ApiPreset = "openai_default";
                cfg.ApiProvider = "openai";
                if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
                {
                    cfg.ApiModelId = "gpt-4o-mini-tts";
                }
                break;
        }
    }

    private static void NormalizeApiSettings(AppConfig cfg)
    {
        if (string.IsNullOrWhiteSpace(cfg.ApiPreset))
        {
            cfg.ApiPreset = GuessApiPreset(cfg);
        }

        ApplyApiPresetDefaults(cfg);

        if (string.IsNullOrWhiteSpace(cfg.ApiProvider))
        {
            cfg.ApiProvider = "openai";
        }
        if (string.IsNullOrWhiteSpace(cfg.ApiModelId))
        {
            cfg.ApiModelId = cfg.ApiProvider.Equals("openai", StringComparison.OrdinalIgnoreCase)
                ? "gpt-4o-mini-tts"
                : cfg.ApiProvider.Equals("alibaba", StringComparison.OrdinalIgnoreCase)
                    ? "qwen3-tts-flash"
                    : "gpt-4o-mini-tts";
        }
        if (string.IsNullOrWhiteSpace(cfg.ApiBaseUrl))
        {
            cfg.ApiBaseUrl = cfg.ApiProvider.Equals("alibaba", StringComparison.OrdinalIgnoreCase)
                ? "https://dashscope-intl.aliyuncs.com/api/v1"
                : string.Empty;
        }
        if (string.IsNullOrWhiteSpace(cfg.ApiLanguageType))
        {
            cfg.ApiLanguageType = "Auto";
        }
        if (string.IsNullOrWhiteSpace(cfg.ApiVoiceDesignTargetModel))
        {
            cfg.ApiVoiceDesignTargetModel = "qwen3-tts-vd-2026-01-26";
        }
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

    private void RefreshModelStatusButton_OnClick(object sender, RoutedEventArgs e)
    {
        RefreshModelInstallChecklist();
        SetRuntimeStatus("Model status refreshed.", isError: false);
    }

    private void RemoveModelCacheButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var cfg = ReadConfigFromUi();
            ApplyPresetDefaults(cfg);
            NormalizeModelSettings(cfg);

            var modelCacheDir = ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot);
            var reposToRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(cfg.ModelRepoId))
            {
                reposToRemove.Add(cfg.ModelRepoId.Trim());
            }
            if (!string.IsNullOrWhiteSpace(cfg.AdditionalModelRepoId))
            {
                reposToRemove.Add(cfg.AdditionalModelRepoId.Trim());
            }

            var removed = 0;
            var failed = new List<string>();
            foreach (var repo in reposToRemove)
            {
                if (!TryResolveRepoFolder(modelCacheDir, repo, out var folder) || !Directory.Exists(folder))
                {
                    continue;
                }

                try
                {
                    Directory.Delete(folder, recursive: true);
                    removed++;
                }
                catch (Exception ex)
                {
                    failed.Add($"{repo}: {ex.Message}");
                }
            }

            RefreshModelInstallChecklist();
            if (failed.Count > 0)
            {
                var detail = string.Join(Environment.NewLine, failed.Take(3));
                SetRuntimeStatus("Failed to remove one or more model caches (file in use).", isError: true);
                MessageBox.Show(
                    this,
                    $"Some model folders could not be removed because files are in use.\n\n{detail}\n\nClose generation and retry.",
                    "Remove Cache",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var message = removed > 0
                ? $"Removed {removed} model cache folder(s)."
                : "No selected model cache folder found.";
            SetRuntimeStatus(message, isError: false);
        }
        catch (Exception ex)
        {
            SetRuntimeStatus($"Failed to remove cache: {ex.Message}", isError: true);
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Remove cache", ex.Message),
                "Remove Cache",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void ModelConfigTextChanged_OnChanged(object sender, TextChangedEventArgs e)
    {
        RefreshModelInstallChecklist();
    }

    private void SetRuntimeStatus(string message, bool isError)
    {
        var text = string.IsNullOrWhiteSpace(message) ? string.Empty : message.Trim();
        RuntimeStatusTextBlock.Text = text;
        RuntimeStatusTextBlock.Foreground = isError
            ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(170, 47, 47))
            : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(79, 125, 99));
    }

    private void OnDownloadTelemetry(ModelDownloadTelemetry t)
    {
        var totalText = t.BytesTotal.HasValue && t.BytesTotal.Value > 0
            ? FormatBytes(t.BytesTotal.Value)
            : "unknown";
        var doneText = FormatBytes(t.BytesDownloaded);
        var speedText = t.BytesPerSecond > 0 ? $"{FormatBytes((long)t.BytesPerSecond)}/s" : "-";
        var percent = t.BytesTotal.HasValue && t.BytesTotal.Value > 0
            ? Math.Clamp((double)t.BytesDownloaded / t.BytesTotal.Value * 100.0, 0.0, 100.0)
            : 0.0;

        if (t.BytesTotal.HasValue && t.BytesTotal.Value > 0)
        {
            DownloadProgressBar.IsIndeterminate = false;
            DownloadProgressBar.Value = percent;
        }
        else
        {
            DownloadProgressBar.IsIndeterminate = true;
        }

        DownloadMetricsTextBlock.Text =
            $"Files {t.FilesCompleted}/{Math.Max(t.FilesTotal, 1)} | {doneText} / {totalText} | {speedText}";
        DownloadCurrentFileTextBlock.Text =
            $"{t.RepoId} :: {t.FilePath} ({FormatBytes(t.FileBytesDownloaded)}"
            + (t.FileBytesTotal.HasValue && t.FileBytesTotal.Value > 0 ? $" / {FormatBytes(t.FileBytesTotal.Value)}" : string.Empty)
            + ")";
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes <= 0)
        {
            return "0 B";
        }

        string[] units = { "B", "KB", "MB", "GB", "TB" };
        var value = (double)bytes;
        var index = 0;
        while (value >= 1024 && index < units.Length - 1)
        {
            value /= 1024;
            index++;
        }

        return $"{value:0.##} {units[index]}";
    }

    private bool HasBundledPythonRuntime(string envName)
    {
        return !string.IsNullOrWhiteSpace(TryFindBundledPythonExe(RuntimePathResolver.AppRoot, envName));
    }

    private static bool UsesPythonLocalBackend(AppConfig cfg)
    {
        return SupportsPythonAlternativePreset(cfg.LocalModelPreset) &&
               string.Equals(NormalizeLocalModelBackendChoice(cfg.LocalModelBackend), "python", StringComparison.OrdinalIgnoreCase);
    }

    private void RefreshModelInstallChecklist()
    {
        var cfg = ReadConfigFromUi();
        ApplyPresetDefaults(cfg);
        NormalizeModelSettings(cfg);
        var preset = (cfg.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        if (preset.StartsWith("api_", StringComparison.OrdinalIgnoreCase))
        {
            // API mode should not show local disk/RAM/model requirements.
            RecommendedTierTextBlock.Text = BuildApiReadinessText(preset);
            QwenChunkLimitsTextBlock.Visibility = Visibility.Collapsed;
            QwenChunkLimitsTextBlock.Text = string.Empty;
            return;
        }

        var chatterRepo = ResolveRepoForPresetAndBackend("chatterbox_onnx", "onnx");
        var chatterPythonRepo = ResolveRepoForPresetAndBackend("chatterbox_onnx", "python");
        var kittenRepo = "KittenML/kitten-tts-mini-0.8";
        var qwenTokenizerRepo = "zukky/Qwen3-TTS-ONNX-DLL";
        var modelCacheDir = ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot);
        var usesPythonBackend = UsesPythonLocalBackend(cfg);
        var pythonRuntimeReady = HasBundledPythonRuntime(ResolvePythonEnvNameForPreset(cfg.LocalModelPreset ?? string.Empty));

        var chatterInstalled = IsRepoInstalled(modelCacheDir, chatterRepo, requiresTokenizer: false);
        var chatterPythonInstalled = IsRepoInstalled(modelCacheDir, chatterPythonRepo, requiresTokenizer: false);
        var kittenModelInstalled = IsRepoInstalled(modelCacheDir, kittenRepo, requiresTokenizer: false);
        var kittenEspeakInstalled = TryFindEspeakNgExecutable(out _);
        var kittenInstalled = kittenModelInstalled && kittenEspeakInstalled;
        var selectedQwenMainInstalled = IsRepoInstalled(modelCacheDir, cfg.ModelRepoId, requiresTokenizer: false);
        var qwenNeedsTokenizerRepo = !string.IsNullOrWhiteSpace(cfg.AdditionalModelRepoId);
        var selectedQwenTokenizerInstalled = !qwenNeedsTokenizerRepo || IsRepoInstalled(modelCacheDir, cfg.AdditionalModelRepoId, requiresTokenizer: true);
        var anyKnownQwenMainInstalled = QwenVariants
            .Where(v => !string.IsNullOrWhiteSpace(v.RepoId))
            .Select(v => v.RepoId)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Any(repo => IsRepoInstalled(modelCacheDir, repo, requiresTokenizer: false));
        var anyQwenTokenizerInstalled = selectedQwenTokenizerInstalled || IsRepoInstalled(modelCacheDir, qwenTokenizerRepo, requiresTokenizer: true);

        var qwenInstalled = preset == "qwen3_tts"
            ? (selectedQwenMainInstalled && selectedQwenTokenizerInstalled)
            : (anyKnownQwenMainInstalled && anyQwenTokenizerInstalled);

        ChatterboxPresetItem.Content = $"{((chatterInstalled || (chatterPythonInstalled && pythonRuntimeReady)) ? "Installed" : "Not Installed")} - Chatterbox";
        QwenPresetItem.Content = $"{(qwenInstalled ? "Installed" : "Not Installed")} - Qwen3 TTS";
        KittenPresetItem.Content = $"{((kittenInstalled || (kittenModelInstalled && pythonRuntimeReady)) ? "Installed" : "Not Installed")} - Kitten TTS Mini 0.8";

        var variant = preset == "qwen3_tts"
            ? ResolveEffectiveQwenVariant(SelectedQwenVariantKey())
            : preset == "kitten_tts"
                ? new QwenVariantOption("kitten", "Kitten TTS Mini 0.8", cfg.ModelRepoId, string.Empty, 2.0, 0.0, 2, 1)
            : new QwenVariantOption("chatterbox", "Chatterbox", cfg.ModelRepoId, cfg.AdditionalModelRepoId, 4.2, 9.8, 12, 6);

        var hasMain = IsRepoInstalled(modelCacheDir, cfg.ModelRepoId, requiresTokenizer: false);
        var needsTokenizer = preset == "qwen3_tts" && qwenNeedsTokenizerRepo;
        var hasTokenizer = !needsTokenizer || IsRepoInstalled(modelCacheDir, cfg.AdditionalModelRepoId, requiresTokenizer: true);

        var freeDiskGb = GetFreeDiskGb(modelCacheDir);
        var ramOk = _systemProfile.RamGb >= variant.MinRamGb;
        var vramRequired = variant.MinVramGb > 0;
        var vramOk = !vramRequired || _systemProfile.GpuVramGb >= variant.MinVramGb;
        var diskOk = !freeDiskGb.HasValue || freeDiskGb.Value >= variant.EstDiskGb;

        MainRepoInstalledCheckBox.IsChecked = hasMain;
        TokenizerInstalledCheckBox.IsChecked = hasTokenizer;
        TokenizerInstalledCheckBox.Visibility = needsTokenizer ? Visibility.Visible : Visibility.Collapsed;
        if (string.Equals(preset, "kitten_tts", StringComparison.OrdinalIgnoreCase))
        {
            MainRepoInstalledCheckBox.Content = $"Main model files installed{(kittenModelInstalled ? string.Empty : " (Kitten model missing)")}";
            TokenizerInstalledCheckBox.Visibility = Visibility.Visible;
            if (usesPythonBackend)
            {
                TokenizerInstalledCheckBox.Content = "Bundled Python runtime ready";
                TokenizerInstalledCheckBox.IsChecked = pythonRuntimeReady;
            }
            else
            {
                TokenizerInstalledCheckBox.Content = "eSpeak-NG installed";
                TokenizerInstalledCheckBox.IsChecked = kittenEspeakInstalled;
            }
        }
        else if (string.Equals(preset, "chatterbox_onnx", StringComparison.OrdinalIgnoreCase) && usesPythonBackend)
        {
            MainRepoInstalledCheckBox.Content = "Main model files installed";
            TokenizerInstalledCheckBox.Visibility = Visibility.Visible;
            TokenizerInstalledCheckBox.Content = "Bundled Python runtime ready";
            TokenizerInstalledCheckBox.IsChecked = pythonRuntimeReady;
        }
        else
        {
            MainRepoInstalledCheckBox.Content = "Main model files installed";
            TokenizerInstalledCheckBox.Content = "Tokenizer files installed";
        }

        RamRequirementCheckBox.IsChecked = ramOk;
        RamRequirementCheckBox.Content = $"RAM: {variant.MinRamGb:0.#}+ GB required (Detected: {_systemProfile.RamGb:0.#} GB)";

        VramRequirementCheckBox.IsChecked = vramOk;
        VramRequirementCheckBox.Visibility = vramRequired ? Visibility.Visible : Visibility.Collapsed;
        VramRequirementCheckBox.Content = $"VRAM: {variant.MinVramGb:0.#}+ GB required (Detected: {_systemProfile.GpuVramGb:0.#} GB)";

        DiskRequirementCheckBox.IsChecked = diskOk;
        var freeDiskText = freeDiskGb.HasValue ? $"{freeDiskGb.Value:0.#} GB free" : "unknown free space";
        DiskRequirementCheckBox.Content = $"Disk: {variant.EstDiskGb:0.#} GB required ({freeDiskText})";

        var recommendedVariant = ResolveEffectiveQwenVariant("auto");
        var isVoiceDesignVariant = variant.RepoId.Contains("voicedesign", StringComparison.OrdinalIgnoreCase);
        var recommendation = preset == "qwen3_tts"
            ? isVoiceDesignVariant
                ? $"VoiceDesign selected: {variant.DisplayName}. Text-only voice design (no source audio) via Python worker. Approx download {variant.EstDownloadGb:0.#} GB, disk {variant.EstDiskGb:0.#} GB. Device: {NormalizePreferDevice(cfg.PreferDevice).ToUpperInvariant()}."
                : $"Recommended variant for this PC: {recommendedVariant.DisplayName}. Approx download {recommendedVariant.EstDownloadGb:0.#} GB, disk {recommendedVariant.EstDiskGb:0.#} GB. Device: {NormalizePreferDevice(cfg.PreferDevice).ToUpperInvariant()}."
            : preset == "kitten_tts"
                ? usesPythonBackend
                    ? "Kitten TTS Mini 0.8 will use the Python backend on this PC."
                    : $"Kitten TTS Mini 0.8 uses built-in model voices (no voice cloning). eSpeak-NG required: {(kittenEspeakInstalled ? "Installed" : "Missing")}."
                : usesPythonBackend
                    ? "Chatterbox will use the Python backend on this PC."
                    : "Recommended tier: Chatterbox ONNX (stable native inference).";
        RecommendedTierTextBlock.Text = recommendation;
        UpdateQwenChunkLimitInfo(cfg, preset);
    }

    private void UpdateQwenChunkLimitInfo(AppConfig cfg, string preset)
    {
        if (!string.Equals(preset, "qwen3_tts", StringComparison.OrdinalIgnoreCase))
        {
            QwenChunkLimitsTextBlock.Visibility = Visibility.Collapsed;
            QwenChunkLimitsTextBlock.Text = string.Empty;
            return;
        }

        if ((cfg.ModelRepoId ?? string.Empty).Contains("voicedesign", StringComparison.OrdinalIgnoreCase))
        {
            QwenChunkLimitsTextBlock.Visibility = Visibility.Visible;
            QwenChunkLimitsTextBlock.Text =
                "Qwen VoiceDesign Notes\n" +
                "Inputs: voice text + style prompt + language (auto/en/zh).\n" +
                "No source audio is required.\n" +
                "Generation uses Python worker and Qwen decode settings from Model Controls.";
            return;
        }

        var settings = GetEffectiveModelSettingsForConfig(cfg);
        var narrationTarget = Math.Clamp(settings.NarrationTargetSec, 2.0, 12.0);
        var dialogueTarget = Math.Clamp(settings.DialogueTargetSec, 2.0, 10.0);

        // Match the Qwen runtime repack cap logic used in MainWindow.
        var recommendedSecondsCap = Math.Clamp(Math.Max(narrationTarget * 3.0, 22.0), 16.0, 48.0);
        var recommendedWordCap = (int)Math.Clamp(Math.Round(recommendedSecondsCap * 3.2), 70, 170);

        var maxNewTokens = ResolveLocalMaxNewTokensEstimate(settings);
        var positionIdsPerSec = TryReadQwenPositionIdsPerSecond(cfg) ?? 13.0;
        var hardCapSeconds = Math.Clamp(maxNewTokens / Math.Max(1.0, positionIdsPerSec), 4.0, 240.0);
        var hardCapWords = (int)Math.Round(hardCapSeconds * 3.2);

        QwenChunkLimitsTextBlock.Visibility = Visibility.Visible;
        QwenChunkLimitsTextBlock.Text =
            "Qwen Chunk Limits (UI estimate)\n" +
            $"Recommended words/chunk: ~{recommendedWordCap} (about {Math.Max(50, recommendedWordCap - 35)}-{Math.Min(220, recommendedWordCap + 25)} depending on punctuation/dialogue)\n" +
            $"Recommended seconds/chunk: ~{recommendedSecondsCap:0.#}s (dialogue base {dialogueTarget:0.#}s, narration base {narrationTarget:0.#}s)\n" +
            $"Hard cap estimate from MaxNewTokens ({maxNewTokens}): ~{hardCapSeconds:0.#}s (~{hardCapWords} words) at {positionIdsPerSec:0.#} pos/sec";
    }

    private SynthesisSettings GetEffectiveModelSettingsForConfig(AppConfig cfg)
    {
        var key = BuildModelProfileKey(cfg.LocalModelPreset, cfg.ModelRepoId);
        if (_modelProfiles.TryGetValue(key, out var perModel) && perModel is not null)
        {
            return CloneSynthesisSettings(perModel);
        }

        return new SynthesisSettings();
    }

    private static int ResolveLocalMaxNewTokensEstimate(SynthesisSettings settings)
    {
        var mode = (settings.ChunkMode ?? "auto").Trim().ToLowerInvariant();
        var chunkChars = mode == "manual"
            ? Math.Clamp(Math.Max(settings.MinChars, Math.Min(settings.MaxChars, settings.ManualMaxChars)), 80, 2400)
            : Math.Clamp(settings.MaxChars, 120, 2400);

        var estimated = chunkChars * 2;
        return Math.Clamp(estimated, 320, 1024);
    }

    private static double? TryReadQwenPositionIdsPerSecond(AppConfig cfg)
    {
        try
        {
            var modelCacheDir = ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot);
            if (!TryResolveRepoFolder(modelCacheDir, cfg.ModelRepoId, out var repoFolder) || !Directory.Exists(repoFolder))
            {
                return null;
            }

            var splitCfg = Path.Combine(repoFolder, "voice_clone_config.json");
            if (File.Exists(splitCfg))
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(splitCfg));
                if (doc.RootElement.TryGetProperty("talker_config", out var talker) &&
                    talker.TryGetProperty("position_id_per_seconds", out var posPerSec))
                {
                    return Math.Clamp(posPerSec.GetDouble(), 1.0, 1000.0);
                }
            }

            foreach (var cfgPath in Directory.GetFiles(repoFolder, "config.json", SearchOption.AllDirectories))
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(cfgPath));
                if (doc.RootElement.TryGetProperty("talker_config", out var talker) &&
                    talker.TryGetProperty("position_id_per_seconds", out var posPerSec))
                {
                    return Math.Clamp(posPerSec.GetDouble(), 1.0, 1000.0);
                }
            }
        }
        catch
        {
            // Fallback in caller.
        }

        return null;
    }

    private static bool IsSelectedPresetInstalled(AppConfig cfg, out string summary)
    {
        var preset = (cfg.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        var modelCacheDir = ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot);
        var mainRepo = string.IsNullOrWhiteSpace(cfg.ModelRepoId) ? "onnx-community/chatterbox-ONNX" : cfg.ModelRepoId.Trim();
        var tokenizerRepo = string.IsNullOrWhiteSpace(cfg.AdditionalModelRepoId) ? string.Empty : cfg.AdditionalModelRepoId.Trim();
        var pythonRuntimeReady = !string.IsNullOrWhiteSpace(TryFindBundledPythonExe(RuntimePathResolver.AppRoot, ResolvePythonEnvNameForPreset(cfg.LocalModelPreset ?? string.Empty)));

        if (preset == "qwen3_tts")
        {
            var baseOk = IsRepoInstalled(modelCacheDir, mainRepo, requiresTokenizer: false);
            var needsTokenizerRepo = !string.IsNullOrWhiteSpace(tokenizerRepo);
            var tokOk = !needsTokenizerRepo || IsRepoInstalled(modelCacheDir, tokenizerRepo, requiresTokenizer: true);
            summary = needsTokenizerRepo
                ? $"Base: {(baseOk ? "[x]" : "[ ]")} | Tokenizer: {(tokOk ? "[x]" : "[ ]")}"
                : $"Repo: {(baseOk ? "[x]" : "[ ]")} (Python Qwen model)";
            return baseOk && tokOk;
        }
        if (preset == "kitten_tts")
        {
            var kittenRepoOk = IsRepoInstalled(modelCacheDir, mainRepo, requiresTokenizer: false);
            if (UsesPythonLocalBackend(cfg))
            {
                summary = $"Repo: {(kittenRepoOk ? "[x]" : "[ ]")} | Python: {(pythonRuntimeReady ? "[x]" : "[ ]")}";
                return kittenRepoOk && pythonRuntimeReady;
            }
            summary = $"Repo: {(kittenRepoOk ? "[x]" : "[ ]")} (config + onnx + voices.npz)";
            return kittenRepoOk;
        }
        if (preset == "chatterbox_onnx" && UsesPythonLocalBackend(cfg))
        {
            var chatterRepoOk = IsRepoInstalled(modelCacheDir, mainRepo, requiresTokenizer: false);
            summary = $"Repo: {(chatterRepoOk ? "[x]" : "[ ]")} | Python: {(pythonRuntimeReady ? "[x]" : "[ ]")}";
            return chatterRepoOk && pythonRuntimeReady;
        }

        var repoOk = IsRepoInstalled(modelCacheDir, mainRepo, requiresTokenizer: false);
        summary = $"Repo: {(repoOk ? "[x]" : "[ ]")}";
        return repoOk;
    }

    private static string? ValidateSelectedPresetFiles(AppConfig cfg)
    {
        var preset = (cfg.LocalModelPreset ?? string.Empty).Trim().ToLowerInvariant();
        var modelCacheDir = ModelCachePath.ResolveAbsolute(cfg.ModelCacheDir, RuntimePaths.AppRoot);
        if (preset == "qwen3_tts")
        {
            var repoId = string.IsNullOrWhiteSpace(cfg.ModelRepoId) ? "Qwen/Qwen3-TTS-12Hz-1.7B-Base" : cfg.ModelRepoId.Trim();
            var tokenizerRepoId = string.IsNullOrWhiteSpace(cfg.AdditionalModelRepoId) ? string.Empty : cfg.AdditionalModelRepoId.Trim();
            if (repoId.Contains("onnx", StringComparison.OrdinalIgnoreCase))
            {
                return ValidateQwenOnnxDllFiles(modelCacheDir, repoId, string.IsNullOrWhiteSpace(tokenizerRepoId) ? "zukky/Qwen3-TTS-ONNX-DLL" : tokenizerRepoId);
            }
            return ValidateQwenPythonBaseFiles(modelCacheDir, repoId);
        }

        if (preset == "chatterbox_onnx")
        {
            var repoId = string.IsNullOrWhiteSpace(cfg.ModelRepoId) ? "onnx-community/chatterbox-ONNX" : cfg.ModelRepoId.Trim();
            if (UsesPythonLocalBackend(cfg))
            {
                return ValidateChatterboxPythonFiles(modelCacheDir, repoId);
            }
            return ValidateChatterboxOnnxFiles(modelCacheDir, repoId);
        }
        if (preset == "kitten_tts")
        {
            var repoId = string.IsNullOrWhiteSpace(cfg.ModelRepoId) ? "KittenML/kitten-tts-mini-0.8" : cfg.ModelRepoId.Trim();
            if (UsesPythonLocalBackend(cfg))
            {
                return ValidateKittenPythonFiles(modelCacheDir, repoId);
            }
            return ValidateKittenFiles(modelCacheDir, repoId);
        }

        return null;
    }

    private static string? ValidateKittenFiles(string modelCacheDir, string repoId)
    {
        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "Kitten model cache folder not found. Download again.";
        }

        var configPath = Path.Combine(repoFolder, "config.json");
        var onnxPath = Path.Combine(repoFolder, "kitten_tts_mini_v0_8.onnx");
        var voicesPath = Path.Combine(repoFolder, "voices.npz");

        foreach (var path in new[] { configPath, onnxPath, voicesPath })
        {
            var sanity = ValidateBinaryModelFile(path, minBytes: 64);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        try
        {
            using var fs = File.OpenRead(voicesPath);
            using var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: false);
            if (zip.Entries.Count == 0 || !zip.Entries.Any(e => e.FullName.EndsWith(".npy", StringComparison.OrdinalIgnoreCase)))
            {
                return "Kitten voices.npz does not contain expected voice embeddings.";
            }
        }
        catch (Exception ex)
        {
            return $"Kitten voices.npz failed validation: {ex.Message}";
        }

        using var sessionOptions = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_DISABLE_ALL };
        try
        {
            using var session = new InferenceSession(onnxPath, sessionOptions);
        }
        catch (Exception ex)
        {
            return $"Kitten ONNX file failed validation: {ex.Message}";
        }

        return null;
    }

    private static string? ValidateKittenPythonFiles(string modelCacheDir, string repoId)
    {
        var runtime = TryFindBundledPythonExe(RuntimePathResolver.AppRoot, "python_kitten");
        if (string.IsNullOrWhiteSpace(runtime))
        {
            return "Bundled Python runtime not found for Kitten Python backend.";
        }

        return ValidateKittenFiles(modelCacheDir, repoId);
    }

    private static string? ValidateChatterboxPythonFiles(string modelCacheDir, string repoId)
    {
        var runtime = TryFindBundledPythonExe(RuntimePathResolver.AppRoot, "python_chatterbox");
        if (string.IsNullOrWhiteSpace(runtime))
        {
            return "Bundled Python runtime not found for Chatterbox Python backend.";
        }

        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "Chatterbox model cache folder not found. Download again.";
        }

        var tokenizerPath = Path.Combine(repoFolder, "tokenizer.json");
        var tokenizerCheck = ValidateTextModelFile(tokenizerPath, minBytes: 16);
        if (!string.IsNullOrWhiteSpace(tokenizerCheck))
        {
            return tokenizerCheck;
        }

        var hasWeights = Directory.GetFiles(repoFolder, "*.safetensors", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.pt", SearchOption.AllDirectories).Length > 0;
        if (!hasWeights)
        {
            return "Chatterbox Python model weights missing (.safetensors/.pt).";
        }

        return null;
    }

    private static string? ValidateQwenPythonBaseFiles(string modelCacheDir, string repoId)
    {
        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "Qwen model cache folder not found. Download again.";
        }

        var mustExist = new[]
        {
            Path.Combine(repoFolder, "config.json"),
            Path.Combine(repoFolder, "tokenizer_config.json")
        };

        foreach (var path in mustExist)
        {
            var sanity = ValidateTextModelFile(path, minBytes: 16);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        var hasWeights = Directory.GetFiles(repoFolder, "*.safetensors", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.bin", SearchOption.AllDirectories).Length > 0;
        if (!hasWeights)
        {
            return "Qwen Python model weights missing (.safetensors/.bin).";
        }

        var hasTokenizerAssets =
            Directory.GetFiles(repoFolder, "tokenizer.json", SearchOption.AllDirectories).Length > 0 ||
            (Directory.GetFiles(repoFolder, "vocab.json", SearchOption.AllDirectories).Length > 0 &&
             Directory.GetFiles(repoFolder, "merges.txt", SearchOption.AllDirectories).Length > 0);
        if (!hasTokenizerAssets)
        {
            return "Qwen tokenizer files missing. Download again.";
        }

        return null;
    }

    private static string? ValidateQwenOnnxDllFiles(string modelCacheDir, string repoId, string tokenizerRepoId)
    {
        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "Qwen model cache folder not found. Download again.";
        }

        if (IsQwenSplitRepoId(repoId) ||
            (Directory.Exists(Path.Combine(repoFolder, "onnx", "shared")) && Directory.Exists(Path.Combine(repoFolder, "onnx", "voice_clone"))))
        {
            return ValidateQwenSplitOnnxFiles(modelCacheDir, repoFolder, tokenizerRepoId);
        }

        var onnx06Dir = Path.Combine(repoFolder, "onnx_kv_06b");
        var onnx17Dir = Path.Combine(repoFolder, "onnx_kv");
        var onnxDir = Directory.Exists(onnx06Dir) ? onnx06Dir : onnx17Dir;
        if (!Directory.Exists(onnxDir))
        {
            return "Qwen ONNX folder missing (onnx_kv/onnx_kv_06b). Use Remove Cache then Download Model Now.";
        }

        var required = new[]
        {
            "speaker_encoder.onnx",
            "talker_prefill.onnx",
            "code_predictor.onnx",
            "text_project.onnx",
            "codec_embed.onnx",
            "code_predictor_embed.onnx",
            "tokenizer12hz_decode.onnx"
        };

        // 1.7B ONNX files exceed protobuf practical limits in ORT C# load path; 0.6B set is the strict native target.
        if (string.Equals(Path.GetFileName(onnxDir), "onnx_kv", StringComparison.OrdinalIgnoreCase))
        {
            var prefill = Path.Combine(onnxDir, "talker_prefill.onnx");
            if (File.Exists(prefill) && new FileInfo(prefill).Length > int.MaxValue)
            {
                if (Directory.Exists(onnx06Dir))
                {
                    onnxDir = onnx06Dir;
                }
                else
                {
                    return "Installed Qwen ONNX set is 1.7B (onnx_kv) which cannot be loaded by strict native ORT due file size. Download onnx_kv_06b variant.";
                }
            }
        }

        foreach (var file in required)
        {
            var path = Path.Combine(onnxDir, file);
            var sanity = ValidateBinaryModelFile(path, minBytes: 4096);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        using var sessionOptions = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_DISABLE_ALL };
        foreach (var file in required)
        {
            var path = Path.Combine(onnxDir, file);
            try
            {
                using var session = new InferenceSession(path, sessionOptions);
            }
            catch (Exception ex)
            {
                return $"Qwen ONNX file failed validation: {file}. {ex.Message}";
            }
        }

        return null;
    }

    private static string? ValidateQwenSplitOnnxFiles(string modelCacheDir, string repoFolder, string tokenizerRepoId)
    {
        var onnxShared = Path.Combine(repoFolder, "onnx", "shared");
        var onnxVoiceClone = Path.Combine(repoFolder, "onnx", "voice_clone");
        if (!Directory.Exists(onnxShared) || !Directory.Exists(onnxVoiceClone))
        {
            return "Qwen split ONNX layout missing (onnx/shared + onnx/voice_clone).";
        }

        var required = new[]
        {
            Path.Combine(repoFolder, "voice_clone_config.json"),
            Path.Combine(onnxShared, "speaker_encoder.onnx"),
            Path.Combine(onnxShared, "speech_tokenizer_decoder.onnx"),
            Path.Combine(onnxVoiceClone, "talker_decode.onnx"),
            Path.Combine(onnxVoiceClone, "text_embedding.onnx"),
            Path.Combine(onnxVoiceClone, "codec_embedding.onnx")
        };

        foreach (var path in required)
        {
            var sanity = ValidateBinaryModelFile(path, minBytes: 4096);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        var hasPredictor = File.Exists(Path.Combine(onnxVoiceClone, "code_predictor_kv.onnx")) ||
                           File.Exists(Path.Combine(onnxVoiceClone, "code_predictor.onnx"));
        if (!hasPredictor)
        {
            return "Qwen split ONNX code predictor file missing (code_predictor.onnx/code_predictor_kv.onnx).";
        }

        var hasCodeEmbed = Directory.GetFiles(onnxVoiceClone, "code_predictor_embed_g*.onnx", SearchOption.TopDirectoryOnly).Length > 0;
        if (!hasCodeEmbed)
        {
            return "Qwen split ONNX code_predictor_embed_g*.onnx files missing.";
        }

        if (!TryResolveQwenTokenizerFiles(modelCacheDir, repoFolder, tokenizerRepoId, out var vocabPath, out var mergesPath, out var tokenizerConfigPath))
        {
            return "Qwen tokenizer files missing. Download tokenizer repo (zukky/Qwen3-TTS-ONNX-DLL) and refresh.";
        }

        var tokenizerTextFiles = new[] { vocabPath, mergesPath, tokenizerConfigPath };
        foreach (var path in tokenizerTextFiles)
        {
            var sanity = ValidateTextModelFile(path, minBytes: 16);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        using var sessionOptions = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_DISABLE_ALL };
        var probeFiles = new[]
        {
            Path.Combine(onnxShared, "speaker_encoder.onnx"),
            Path.Combine(onnxShared, "speech_tokenizer_decoder.onnx"),
            Path.Combine(onnxVoiceClone, "text_embedding.onnx"),
            Path.Combine(onnxVoiceClone, "codec_embedding.onnx"),
            Path.Combine(onnxVoiceClone, "talker_decode.onnx")
        };
        foreach (var path in probeFiles)
        {
            try
            {
                using var session = new InferenceSession(path, sessionOptions);
            }
            catch (Exception ex)
            {
                return $"Qwen split ONNX file failed validation: {Path.GetFileName(path)}. {ex.Message}";
            }
        }

        return null;
    }

    private static bool TryResolveQwenTokenizerFiles(
        string modelCacheDir,
        string repoFolder,
        string tokenizerRepoId,
        out string vocabPath,
        out string mergesPath,
        out string tokenizerConfigPath)
    {
        vocabPath = string.Empty;
        mergesPath = string.Empty;
        tokenizerConfigPath = string.Empty;

        if (TryFindTokenizerTriplet(repoFolder, out vocabPath, out mergesPath, out tokenizerConfigPath))
        {
            return true;
        }

        if (TryResolveRepoFolder(modelCacheDir, tokenizerRepoId, out var tokenizerRepoFolder) &&
            Directory.Exists(tokenizerRepoFolder) &&
            TryFindTokenizerTriplet(tokenizerRepoFolder, out vocabPath, out mergesPath, out tokenizerConfigPath))
        {
            return true;
        }

        var hfRoot = Path.Combine(modelCacheDir, "hf-cache");
        if (!Directory.Exists(hfRoot))
        {
            return false;
        }

        var candidates = new List<(string VocabPath, string MergesPath, string TokenizerConfigPath)>();
        foreach (var foundVocab in Directory.GetFiles(hfRoot, "vocab.json", SearchOption.AllDirectories))
        {
            var dir = Path.GetDirectoryName(foundVocab);
            if (string.IsNullOrWhiteSpace(dir))
            {
                continue;
            }

            var foundMerges = Path.Combine(dir, "merges.txt");
            var foundCfg = Path.Combine(dir, "tokenizer_config.json");
            if (File.Exists(foundMerges) && File.Exists(foundCfg))
            {
                candidates.Add((foundVocab, foundMerges, foundCfg));
            }
        }

        if (candidates.Count == 0)
        {
            return false;
        }

        var preferred = candidates.FirstOrDefault(c =>
            c.VocabPath.IndexOf("1.7B", StringComparison.OrdinalIgnoreCase) >= 0 ||
            c.VocabPath.IndexOf("12Hz-1.7B-Base", StringComparison.OrdinalIgnoreCase) >= 0);
        var selected = string.IsNullOrWhiteSpace(preferred.VocabPath) ? candidates[0] : preferred;
        vocabPath = selected.VocabPath;
        mergesPath = selected.MergesPath;
        tokenizerConfigPath = selected.TokenizerConfigPath;
        return true;
    }

    private static bool TryFindTokenizerTriplet(
        string searchRoot,
        out string vocabPath,
        out string mergesPath,
        out string tokenizerConfigPath)
    {
        vocabPath = string.Empty;
        mergesPath = string.Empty;
        tokenizerConfigPath = string.Empty;

        var directTokenizerDir = Path.Combine(searchRoot, "tokenizer");
        var directVocab = Path.Combine(directTokenizerDir, "vocab.json");
        var directMerges = Path.Combine(directTokenizerDir, "merges.txt");
        var directCfg = Path.Combine(directTokenizerDir, "tokenizer_config.json");
        if (File.Exists(directVocab) && File.Exists(directMerges) && File.Exists(directCfg))
        {
            vocabPath = directVocab;
            mergesPath = directMerges;
            tokenizerConfigPath = directCfg;
            return true;
        }

        foreach (var foundVocab in Directory.GetFiles(searchRoot, "vocab.json", SearchOption.AllDirectories))
        {
            var dir = Path.GetDirectoryName(foundVocab);
            if (string.IsNullOrWhiteSpace(dir))
            {
                continue;
            }

            var foundMerges = Path.Combine(dir, "merges.txt");
            var foundCfg = Path.Combine(dir, "tokenizer_config.json");
            if (File.Exists(foundMerges) && File.Exists(foundCfg))
            {
                vocabPath = foundVocab;
                mergesPath = foundMerges;
                tokenizerConfigPath = foundCfg;
                return true;
            }
        }

        return false;
    }

    private static string? ValidateChatterboxOnnxFiles(string modelCacheDir, string repoId)
    {
        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder) || !Directory.Exists(repoFolder))
        {
            return "Chatterbox model cache folder not found. Download again.";
        }

        var required = new[]
        {
            "onnx/conditional_decoder.onnx",
            "onnx/conditional_decoder.onnx_data",
            "onnx/embed_tokens.onnx",
            "onnx/embed_tokens.onnx_data",
            "onnx/language_model.onnx",
            "onnx/language_model.onnx_data",
            "onnx/speech_encoder.onnx",
            "onnx/speech_encoder.onnx_data"
        };
        foreach (var rel in required)
        {
            var path = Path.Combine(repoFolder, rel.Replace('/', Path.DirectorySeparatorChar));
            var minBytes = rel.EndsWith(".onnx_data", StringComparison.OrdinalIgnoreCase) ? 1024L : 4096L;
            var sanity = ValidateBinaryModelFile(path, minBytes);
            if (!string.IsNullOrWhiteSpace(sanity))
            {
                return sanity;
            }
        }

        using var sessionOptions = new SessionOptions { GraphOptimizationLevel = GraphOptimizationLevel.ORT_DISABLE_ALL };
        var onnxToLoad = new[]
        {
            "onnx/conditional_decoder.onnx",
            "onnx/embed_tokens.onnx",
            "onnx/language_model.onnx",
            "onnx/speech_encoder.onnx"
        };
        foreach (var rel in onnxToLoad)
        {
            var path = Path.Combine(repoFolder, rel.Replace('/', Path.DirectorySeparatorChar));
            try
            {
                using var session = new InferenceSession(path, sessionOptions);
            }
            catch (Exception ex)
            {
                return $"Chatterbox ONNX file failed validation: {Path.GetFileName(path)}. {ex.Message}";
            }
        }

        return null;
    }

    private static string? ValidateBinaryModelFile(string path, long minBytes)
    {
        if (!File.Exists(path))
        {
            return $"Required model file missing: {path}";
        }

        var info = new FileInfo(path);
        if (info.Length < minBytes)
        {
            return $"Model file looks incomplete (too small): {path}";
        }

        try
        {
            var probeLength = (int)Math.Min(512, info.Length);
            var buffer = new byte[probeLength];
            using var fs = File.OpenRead(path);
            var read = fs.Read(buffer, 0, probeLength);
            if (read > 0)
            {
                var head = System.Text.Encoding.UTF8.GetString(buffer, 0, read);
                if (head.Contains("git-lfs.github.com/spec/v1", StringComparison.OrdinalIgnoreCase) ||
                    head.Contains("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) ||
                    head.Contains("<html", StringComparison.OrdinalIgnoreCase))
                {
                    return $"Model file is invalid text/html pointer instead of binary: {path}";
                }
            }
        }
        catch (Exception ex)
        {
            return $"Failed reading model file {path}: {ex.Message}";
        }

        return null;
    }

    private static string? ValidateTextModelFile(string path, long minBytes)
    {
        if (!File.Exists(path))
        {
            return $"Required model file missing: {path}";
        }

        var info = new FileInfo(path);
        if (info.Length < minBytes)
        {
            return $"Model file looks incomplete (too small): {path}";
        }

        try
        {
            var probeLength = (int)Math.Min(512, info.Length);
            var buffer = new byte[probeLength];
            using var fs = File.OpenRead(path);
            var read = fs.Read(buffer, 0, probeLength);
            if (read > 0)
            {
                var head = System.Text.Encoding.UTF8.GetString(buffer, 0, read);
                if (head.Contains("git-lfs.github.com/spec/v1", StringComparison.OrdinalIgnoreCase) ||
                    head.Contains("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) ||
                    head.Contains("<html", StringComparison.OrdinalIgnoreCase))
                {
                    return $"Model file is invalid text/html pointer instead of tokenizer asset: {path}";
                }
            }
        }
        catch (Exception ex)
        {
            return $"Failed reading model file {path}: {ex.Message}";
        }

        return null;
    }

    private static double? GetFreeDiskGb(string modelCacheDir)
    {
        try
        {
            var root = Path.GetPathRoot(modelCacheDir);
            if (string.IsNullOrWhiteSpace(root))
            {
                return null;
            }

            var drive = new DriveInfo(root);
            if (!drive.IsReady)
            {
                return null;
            }

            return Math.Round(drive.AvailableFreeSpace / 1024d / 1024d / 1024d, 1);
        }
        catch
        {
            return null;
        }
    }

    private static bool IsRepoInstalled(string modelCacheDir, string repoId, bool requiresTokenizer)
    {
        if (!TryResolveRepoFolder(modelCacheDir, repoId, out var repoFolder))
        {
            return false;
        }

        if (!Directory.Exists(repoFolder))
        {
            return false;
        }

        var hasConfig = File.Exists(Path.Combine(repoFolder, "config.json")) ||
                        File.Exists(Path.Combine(repoFolder, "generation_config.json")) ||
                        File.Exists(Path.Combine(repoFolder, "README.md"));

        var hasTokenizer = File.Exists(Path.Combine(repoFolder, "tokenizer.json")) ||
                           File.Exists(Path.Combine(repoFolder, "tokenizer_config.json")) ||
                           Directory.GetFiles(repoFolder, "tokenizer.*", SearchOption.AllDirectories).Length > 0 ||
                           Directory.GetFiles(repoFolder, "tokenizer*.json", SearchOption.AllDirectories).Length > 0;
        var hasTokenizerAssets = Directory.GetFiles(repoFolder, "vocab.json", SearchOption.AllDirectories).Length > 0 &&
                                 Directory.GetFiles(repoFolder, "merges.txt", SearchOption.AllDirectories).Length > 0;
        var isQwenOnnxDll = repoId.IndexOf("qwen3-tts-onnx-dll", StringComparison.OrdinalIgnoreCase) >= 0;
        var isQwenSplit = IsQwenSplitRepoId(repoId);
        var isKitten = IsKittenRepoId(repoId);
        var hasQwenRustDll = Directory.GetFiles(repoFolder, "qwen3_tts_rust.dll", SearchOption.AllDirectories).Length > 0;
        var onnxKvDir = Path.Combine(repoFolder, "onnx_kv");
        var onnxKv06Dir = Path.Combine(repoFolder, "onnx_kv_06b");
        static bool HasRequiredQwenOnnxSet(string dir)
        {
            if (!Directory.Exists(dir))
            {
                return false;
            }

            var required = new[]
            {
                "speaker_encoder.onnx",
                "talker_prefill.onnx",
                "code_predictor.onnx",
                "text_project.onnx",
                "codec_embed.onnx",
                "code_predictor_embed.onnx",
                "tokenizer12hz_decode.onnx"
            };

            return required.All(name => File.Exists(Path.Combine(dir, name)));
        }

        var hasQwenOnnxKv = HasRequiredQwenOnnxSet(onnxKvDir) || HasRequiredQwenOnnxSet(onnxKv06Dir);
        var hasQwenSplitOnnx = HasRequiredQwenSplitSet(repoFolder);
        var hasKittenSet = HasRequiredKittenSet(repoFolder);

        var hasWeights = Directory.GetFiles(repoFolder, "*.onnx", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.safetensors", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.bin", SearchOption.AllDirectories).Length > 0 ||
                         Directory.GetFiles(repoFolder, "*.pt", SearchOption.AllDirectories).Length > 0;

        if (requiresTokenizer)
        {
            // Some tokenizer repos (for example Qwen3 tokenizer) package tokenizer weights as model.safetensors
            // and config files without a tokenizer.json. Treat that structure as installed as well.
            var looksLikeTokenizerRepo = repoId.IndexOf("tokenizer", StringComparison.OrdinalIgnoreCase) >= 0;
            var hasTokenizerModelLayout = File.Exists(Path.Combine(repoFolder, "preprocessor_config.json")) ||
                                          (looksLikeTokenizerRepo && (hasConfig || hasWeights));
            return hasTokenizer || hasTokenizerModelLayout || hasTokenizerAssets;
        }

        if (isQwenOnnxDll)
        {
            return hasQwenRustDll && hasQwenOnnxKv && hasTokenizerAssets;
        }
        if (isQwenSplit)
        {
            return hasQwenSplitOnnx;
        }
        if (isKitten)
        {
            return hasKittenSet;
        }

        return hasConfig && hasWeights;
    }

    private static bool IsQwenSplitRepoId(string repoId)
    {
        var normalized = (repoId ?? string.Empty).Trim();
        return normalized.IndexOf("qwen3-tts", StringComparison.OrdinalIgnoreCase) >= 0 &&
               normalized.IndexOf("onnx", StringComparison.OrdinalIgnoreCase) >= 0 &&
               normalized.IndexOf("onnx-dll", StringComparison.OrdinalIgnoreCase) < 0;
    }

    private static bool IsKittenRepoId(string repoId)
    {
        return (repoId ?? string.Empty).IndexOf("kitten-tts", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool HasRequiredKittenSet(string repoFolder)
    {
        return File.Exists(Path.Combine(repoFolder, "config.json")) &&
               File.Exists(Path.Combine(repoFolder, "kitten_tts_mini_v0_8.onnx")) &&
               File.Exists(Path.Combine(repoFolder, "voices.npz"));
    }

    private static bool HasRequiredQwenSplitSet(string repoFolder)
    {
        var onnxShared = Path.Combine(repoFolder, "onnx", "shared");
        var onnxVoiceClone = Path.Combine(repoFolder, "onnx", "voice_clone");
        if (!Directory.Exists(onnxShared) || !Directory.Exists(onnxVoiceClone))
        {
            return false;
        }

        var required = new[]
        {
            Path.Combine(repoFolder, "voice_clone_config.json"),
            Path.Combine(onnxShared, "speaker_encoder.onnx"),
            Path.Combine(onnxShared, "speech_tokenizer_decoder.onnx"),
            Path.Combine(onnxVoiceClone, "talker_decode.onnx"),
            Path.Combine(onnxVoiceClone, "text_embedding.onnx"),
            Path.Combine(onnxVoiceClone, "codec_embedding.onnx")
        };
        if (required.Any(path => !File.Exists(path)))
        {
            return false;
        }

        var hasPredictor = File.Exists(Path.Combine(onnxVoiceClone, "code_predictor_kv.onnx")) ||
                           File.Exists(Path.Combine(onnxVoiceClone, "code_predictor.onnx"));
        if (!hasPredictor)
        {
            return false;
        }

        return Directory.GetFiles(onnxVoiceClone, "code_predictor_embed_g*.onnx", SearchOption.TopDirectoryOnly).Length > 0;
    }

    private static bool TryResolveRepoFolder(string modelCacheDir, string repoId, out string folder)
    {
        folder = string.Empty;
        if (string.IsNullOrWhiteSpace(repoId))
        {
            return false;
        }

        var tokens = repoId.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2)
        {
            return false;
        }

        folder = Path.Combine(modelCacheDir, "hf-cache", $"models--{tokens[0]}--{tokens[1]}");
        return true;
    }

    private void ModelControlsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var cfg = ReadConfigFromUi();
        ApplyPresetDefaults(cfg);
        NormalizeModelSettings(cfg);
        var key = BuildModelProfileKey(cfg.LocalModelPreset, cfg.ModelRepoId);
        if (!_modelProfiles.TryGetValue(key, out var modelSettings))
        {
            modelSettings = new SynthesisSettings();
        }
        var pauseDefaults = GetDefaultPauseRangesForModel(cfg.LocalModelPreset, cfg.ModelRepoId);
        var isQwenModel = (cfg.LocalModelPreset ?? string.Empty).Equals("qwen3_tts", StringComparison.OrdinalIgnoreCase) ||
                          (cfg.ModelRepoId ?? string.Empty).IndexOf("qwen3-tts", StringComparison.OrdinalIgnoreCase) >= 0;

        var dlg = new Window
        {
            Title = $"Model Settings: {cfg.ModelRepoId}",
            Width = 560,
            Height = 760,
            MinHeight = 560,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.CanResize,
            Background = System.Windows.Media.Brushes.White
        };

        var root = new System.Windows.Controls.Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var header = new System.Windows.Controls.TextBlock
        {
            Text = "Per-Model Generation Settings",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold
        };
        System.Windows.Controls.Grid.SetRow(header, 0);
        root.Children.Add(header);

        var panel = new System.Windows.Controls.Grid { Margin = new Thickness(0, 12, 0, 0) };
        panel.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(180) });
        panel.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var i = 0; i < 15; i++)
        {
            panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        }

        var chunkModeLabel = new System.Windows.Controls.TextBlock { Text = "Chunk Mode", VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(chunkModeLabel, 0);
        panel.Children.Add(chunkModeLabel);
        var chunkModeCombo = new System.Windows.Controls.ComboBox { Height = 32 };
        chunkModeCombo.Items.Add("auto");
        chunkModeCombo.Items.Add("manual");
        chunkModeCombo.Text = string.IsNullOrWhiteSpace(modelSettings.ChunkMode) ? "auto" : modelSettings.ChunkMode;
        System.Windows.Controls.Grid.SetRow(chunkModeCombo, 0);
        System.Windows.Controls.Grid.SetColumn(chunkModeCombo, 1);
        panel.Children.Add(chunkModeCombo);

        var minCharsLabel = new System.Windows.Controls.TextBlock { Text = "Min Chars", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(minCharsLabel, 1);
        panel.Children.Add(minCharsLabel);
        var minCharsBox = new System.Windows.Controls.TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = modelSettings.MinChars.ToString() };
        System.Windows.Controls.Grid.SetRow(minCharsBox, 1);
        System.Windows.Controls.Grid.SetColumn(minCharsBox, 1);
        panel.Children.Add(minCharsBox);

        var maxCharsLabel = new System.Windows.Controls.TextBlock { Text = "Max Chars (Auto)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(maxCharsLabel, 2);
        panel.Children.Add(maxCharsLabel);
        var maxCharsBox = new System.Windows.Controls.TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = modelSettings.MaxChars.ToString() };
        System.Windows.Controls.Grid.SetRow(maxCharsBox, 2);
        System.Windows.Controls.Grid.SetColumn(maxCharsBox, 1);
        panel.Children.Add(maxCharsBox);

        var manualCharsLabel = new System.Windows.Controls.TextBlock { Text = "Max Chars (Manual)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(manualCharsLabel, 3);
        panel.Children.Add(manualCharsLabel);
        var manualCharsBox = new System.Windows.Controls.TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = modelSettings.ManualMaxChars.ToString() };
        System.Windows.Controls.Grid.SetRow(manualCharsBox, 3);
        System.Windows.Controls.Grid.SetColumn(manualCharsBox, 1);
        panel.Children.Add(manualCharsBox);

        var chunkPauseLabel = new System.Windows.Controls.TextBlock { Text = "Chunk Pause (ms)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(chunkPauseLabel, 4);
        panel.Children.Add(chunkPauseLabel);
        var chunkPauseBox = new System.Windows.Controls.TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = modelSettings.ChunkPauseMs.ToString() };
        System.Windows.Controls.Grid.SetRow(chunkPauseBox, 4);
        System.Windows.Controls.Grid.SetColumn(chunkPauseBox, 1);
        panel.Children.Add(chunkPauseBox);

        var paragraphPauseLabel = new System.Windows.Controls.TextBlock { Text = "Paragraph Pause (ms)", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(paragraphPauseLabel, 5);
        panel.Children.Add(paragraphPauseLabel);
        var paragraphPauseBox = new System.Windows.Controls.TextBox { Height = 32, Margin = new Thickness(0, 10, 0, 0), Text = modelSettings.ParagraphPauseMs.ToString() };
        System.Windows.Controls.Grid.SetRow(paragraphPauseBox, 5);
        System.Windows.Controls.Grid.SetColumn(paragraphPauseBox, 1);
        panel.Children.Add(paragraphPauseBox);

        var effectiveClauseRange = ResolvePauseRange(modelSettings.ClausePauseMinMs, modelSettings.ClausePauseMaxMs, pauseDefaults.Clause);
        var effectiveSentenceRange = ResolvePauseRange(modelSettings.SentencePauseMinMs, modelSettings.SentencePauseMaxMs, pauseDefaults.Sentence);
        var effectiveEllipsisRange = ResolvePauseRange(modelSettings.EllipsisPauseMinMs, modelSettings.EllipsisPauseMaxMs, pauseDefaults.Ellipsis);
        var effectiveParagraphRange = ResolvePauseRange(modelSettings.ParagraphPauseMinMs, modelSettings.ParagraphPauseMaxMs, pauseDefaults.Paragraph);

        AddPauseRangeRow(panel, 6, "Clause Pause Range", effectiveClauseRange, out var clausePauseMinBox, out var clausePauseMaxBox);
        AddPauseRangeRow(panel, 7, "Sentence Pause Range", effectiveSentenceRange, out var sentencePauseMinBox, out var sentencePauseMaxBox);
        AddPauseRangeRow(panel, 8, "Ellipsis Pause Range", effectiveEllipsisRange, out var ellipsisPauseMinBox, out var ellipsisPauseMaxBox);
        AddPauseRangeRow(panel, 9, "Paragraph Pause Range", effectiveParagraphRange, out var paragraphPauseMinBox, out var paragraphPauseMaxBox);

        var outFmtLabel = new System.Windows.Controls.TextBlock { Text = "Output Format", Margin = new Thickness(0, 10, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        System.Windows.Controls.Grid.SetRow(outFmtLabel, 10);
        panel.Children.Add(outFmtLabel);
        var outFmtCombo = new System.Windows.Controls.ComboBox { Height = 32, Margin = new Thickness(0, 10, 0, 0) };
        outFmtCombo.Items.Add("wav");
        outFmtCombo.Items.Add("mp3");
        outFmtCombo.Text = NormalizeLocalOutputExtension(modelSettings.OutputFormat);
        System.Windows.Controls.Grid.SetRow(outFmtCombo, 10);
        System.Windows.Controls.Grid.SetColumn(outFmtCombo, 1);
        panel.Children.Add(outFmtCombo);

        var protectBracket = new System.Windows.Controls.CheckBox
        {
            Content = "Preserve [ ... ] directives while chunking",
            Margin = new Thickness(0, 12, 0, 0),
            IsChecked = modelSettings.ProtectBracketDirectives
        };
        System.Windows.Controls.Grid.SetRow(protectBracket, 11);
        System.Windows.Controls.Grid.SetColumnSpan(protectBracket, 2);
        panel.Children.Add(protectBracket);

        var hintPanel = new System.Windows.Controls.Grid { Margin = new Thickness(0, 10, 0, 0) };
        hintPanel.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(180) });
        hintPanel.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var hintLabel = new System.Windows.Controls.TextBlock { Text = "Instruction Hint", VerticalAlignment = VerticalAlignment.Center };
        var hintBox = new System.Windows.Controls.TextBox
        {
            Height = 32,
            Text = modelSettings.LocalInstructionHint ?? string.Empty,
            ToolTip = "For Qwen models, this is prefixed as [hint] before each chunk."
        };
        System.Windows.Controls.Grid.SetColumn(hintBox, 1);
        hintPanel.Children.Add(hintLabel);
        hintPanel.Children.Add(hintBox);
        System.Windows.Controls.Grid.SetRow(hintPanel, 12);
        System.Windows.Controls.Grid.SetColumnSpan(hintPanel, 2);
        panel.Children.Add(hintPanel);

        var qwenUseRefTextCheck = new System.Windows.Controls.CheckBox
        {
            Content = "Qwen: Use ref_text (ICL clone mode)",
            Margin = new Thickness(0, 12, 0, 0),
            IsChecked = modelSettings.QwenUseRefText,
            IsEnabled = isQwenModel
        };
        System.Windows.Controls.Grid.SetRow(qwenUseRefTextCheck, 13);
        System.Windows.Controls.Grid.SetColumnSpan(qwenUseRefTextCheck, 2);
        panel.Children.Add(qwenUseRefTextCheck);

        var qwenRefTextNote = new System.Windows.Controls.TextBlock
        {
            Text = "Turn OFF if voice sample transcript leaks into output. OFF uses x-vector-only clone (safer for long generation).",
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = System.Windows.Media.Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            Visibility = isQwenModel ? Visibility.Visible : Visibility.Collapsed
        };
        System.Windows.Controls.Grid.SetRow(qwenRefTextNote, 14);
        System.Windows.Controls.Grid.SetColumnSpan(qwenRefTextNote, 2);
        panel.Children.Add(qwenRefTextNote);

        var panelScroll = new System.Windows.Controls.ScrollViewer
        {
            Content = panel,
            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled
        };
        System.Windows.Controls.Grid.SetRow(panelScroll, 1);
        root.Children.Add(panelScroll);

        var buttons = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        var cancelBtn = new System.Windows.Controls.Button { Content = "Cancel", Width = 90, Height = 32, Margin = new Thickness(0, 0, 8, 0) };
        cancelBtn.Click += (_, _) => { dlg.DialogResult = false; dlg.Close(); };
        var applyBtn = new System.Windows.Controls.Button { Content = "Apply", Width = 100, Height = 34 };
        applyBtn.Click += (_, _) =>
        {
            var updated = CloneSynthesisSettings(modelSettings);
            updated.ChunkMode = (chunkModeCombo.Text ?? "auto").Trim().ToLowerInvariant() == "manual" ? "manual" : "auto";
            updated.MinChars = Math.Clamp(ParseIntOrFallback(minCharsBox.Text, updated.MinChars), 40, 3000);
            updated.MaxChars = Math.Clamp(ParseIntOrFallback(maxCharsBox.Text, updated.MaxChars), 80, 4000);
            updated.ManualMaxChars = Math.Clamp(ParseIntOrFallback(manualCharsBox.Text, updated.ManualMaxChars), 80, 6000);
            updated.ChunkPauseMs = Math.Clamp(ParseIntOrFallback(chunkPauseBox.Text, updated.ChunkPauseMs), 0, 5000);
            updated.ParagraphPauseMs = Math.Clamp(ParseIntOrFallback(paragraphPauseBox.Text, updated.ParagraphPauseMs), 0, 10000);
            updated.ClausePauseMinMs = Math.Clamp(ParseIntOrFallback(clausePauseMinBox.Text, effectiveClauseRange.Min), 0, 5000);
            updated.ClausePauseMaxMs = Math.Clamp(ParseIntOrFallback(clausePauseMaxBox.Text, effectiveClauseRange.Max), 0, 5000);
            updated.SentencePauseMinMs = Math.Clamp(ParseIntOrFallback(sentencePauseMinBox.Text, effectiveSentenceRange.Min), 0, 10000);
            updated.SentencePauseMaxMs = Math.Clamp(ParseIntOrFallback(sentencePauseMaxBox.Text, effectiveSentenceRange.Max), 0, 10000);
            updated.EllipsisPauseMinMs = Math.Clamp(ParseIntOrFallback(ellipsisPauseMinBox.Text, effectiveEllipsisRange.Min), 0, 10000);
            updated.EllipsisPauseMaxMs = Math.Clamp(ParseIntOrFallback(ellipsisPauseMaxBox.Text, effectiveEllipsisRange.Max), 0, 10000);
            updated.ParagraphPauseMinMs = Math.Clamp(ParseIntOrFallback(paragraphPauseMinBox.Text, effectiveParagraphRange.Min), 0, 15000);
            updated.ParagraphPauseMaxMs = Math.Clamp(ParseIntOrFallback(paragraphPauseMaxBox.Text, effectiveParagraphRange.Max), 0, 15000);
            updated.OutputFormat = NormalizeLocalOutputExtension(outFmtCombo.Text);
            updated.ProtectBracketDirectives = protectBracket.IsChecked == true;
            updated.LocalInstructionHint = (hintBox.Text ?? string.Empty).Trim();
            updated.QwenUseRefText = qwenUseRefTextCheck.IsChecked != false;
            if (updated.MinChars > updated.MaxChars)
            {
                updated.MinChars = updated.MaxChars;
            }
            NormalizePauseRangeOrder(updated);

            OverwriteModelProfileEverywhere(cfg.LocalModelPreset, cfg.ModelRepoId, updated);
            dlg.DialogResult = true;
            dlg.Close();
        };
        buttons.Children.Add(cancelBtn);
        buttons.Children.Add(applyBtn);
        System.Windows.Controls.Grid.SetRow(buttons, 2);
        root.Children.Add(buttons);

        dlg.Content = root;
        if (dlg.ShowDialog() == true)
        {
            StatusTextBlock.Text = $"Saved per-model settings for: {cfg.ModelRepoId}";
        }
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

    private static IEnumerable<string> BuildEquivalentModelProfileKeys(string? preset, string? repoId)
    {
        var repo = (repoId ?? string.Empty).Trim();
        var p = (preset ?? string.Empty).Trim();
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

        Add(BuildModelProfileKey(preset, repoId));
        Add(repo.ToLowerInvariant());
        Add(p.ToLowerInvariant());

        if (DetectModelFamily(repo) == ModelFamily.Qwen || DetectModelFamily(p) == ModelFamily.Qwen)
        {
            Add("qwen3_tts");
            Add("qwen/qwen3-tts-12hz-1.7b-base");
            Add("xkos/qwen3-tts-12hz-1.7b-onnx");
            Add("zukky/qwen3-tts-onnx-dll");
        }
        else if (DetectModelFamily(repo) == ModelFamily.Chatterbox || DetectModelFamily(p) == ModelFamily.Chatterbox)
        {
            Add("chatterbox_onnx");
            Add("onnx-community/chatterbox-onnx");
        }
        else if (DetectModelFamily(repo) == ModelFamily.Kitten || DetectModelFamily(p) == ModelFamily.Kitten)
        {
            Add("kitten_tts");
            Add("kittenml/kitten-tts-mini-0.8");
        }

        return keys;
    }

    private void OverwriteModelProfileEverywhere(string? preset, string? repoId, SynthesisSettings settings)
    {
        var snapshot = CloneSynthesisSettings(settings);
        var family = DetectModelFamily(repoId);
        var keysToWrite = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var key in BuildEquivalentModelProfileKeys(preset, repoId))
        {
            keysToWrite.Add(key);
        }

        foreach (var existingKey in _modelProfiles.Keys.ToList())
        {
            if (IsFamilyKey(family, existingKey))
            {
                keysToWrite.Add(existingKey);
            }
        }

        foreach (var key in keysToWrite)
        {
            _modelProfiles[key] = CloneSynthesisSettings(snapshot);
        }
    }

    private sealed record QwenVariantOption(
        string Key,
        string DisplayName,
        string RepoId,
        string TokenizerRepo,
        double EstDownloadGb,
        double EstDiskGb,
        double MinRamGb,
        double MinVramGb);

    private static int ParseIntOrFallback(string? raw, int fallback)
    {
        return int.TryParse(raw?.Trim(), out var value) ? value : fallback;
    }

    private static double ParseDoubleOrFallback(string? raw, double fallback)
    {
        var s = raw?.Trim();
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
        {
            return value;
        }
        return fallback;
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

    private static string BuildModelProfileKey(string? preset, string? repoId)
    {
        var repo = (repoId ?? string.Empty).Trim().ToLowerInvariant();
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

        var p = (preset ?? string.Empty).Trim().ToLowerInvariant();
        return string.IsNullOrWhiteSpace(p) ? "chatterbox_onnx" : p;
    }

    private static ((int Min, int Max) Clause, (int Min, int Max) Sentence, (int Min, int Max) Paragraph, (int Min, int Max) Ellipsis)
        GetDefaultPauseRangesForModel(string? preset, string? repoId)
    {
        var p = (preset ?? string.Empty).Trim();
        var repo = (repoId ?? string.Empty).Trim();
        var isQwen = p.Equals("qwen3_tts", StringComparison.OrdinalIgnoreCase) ||
                     repo.IndexOf("qwen3-tts", StringComparison.OrdinalIgnoreCase) >= 0;
        var isKitten = p.Equals("kitten_tts", StringComparison.OrdinalIgnoreCase) ||
                       repo.IndexOf("kitten-tts", StringComparison.OrdinalIgnoreCase) >= 0;

        if (isQwen)
        {
            return ((180, 260), (420, 620), (820, 1200), (520, 760));
        }

        if (isKitten)
        {
            return ((170, 240), (400, 580), (780, 1120), (500, 720));
        }

        return ((190, 280), (480, 700), (900, 1350), (560, 820));
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
            if (max > 0 && min == 0) min = max;
            if (min > 0 && max == 0) max = min;
            if (max < min) (min, max) = (max, min);
        }
    }

    private static void AddPauseRangeRow(System.Windows.Controls.Grid panel, int row, string label, (int Min, int Max) range,
        out System.Windows.Controls.TextBox minBox, out System.Windows.Controls.TextBox maxBox)
    {
        var textLabel = new System.Windows.Controls.TextBlock
        {
            Text = $"{label} (ms)",
            Margin = new Thickness(0, 10, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        System.Windows.Controls.Grid.SetRow(textLabel, row);
        panel.Children.Add(textLabel);

        var host = new System.Windows.Controls.Grid { Margin = new Thickness(0, 10, 0, 0) };
        host.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        host.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
        host.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        minBox = new System.Windows.Controls.TextBox { Height = 32, Text = range.Min.ToString(System.Globalization.CultureInfo.InvariantCulture) };
        maxBox = new System.Windows.Controls.TextBox { Height = 32, Text = range.Max.ToString(System.Globalization.CultureInfo.InvariantCulture) };
        var sep = new System.Windows.Controls.TextBlock
        {
            Text = "to",
            Margin = new Thickness(8, 0, 8, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

        System.Windows.Controls.Grid.SetColumn(minBox, 0);
        System.Windows.Controls.Grid.SetColumn(sep, 1);
        System.Windows.Controls.Grid.SetColumn(maxBox, 2);
        host.Children.Add(minBox);
        host.Children.Add(sep);
        host.Children.Add(maxBox);

        System.Windows.Controls.Grid.SetRow(host, row);
        System.Windows.Controls.Grid.SetColumn(host, 1);
        panel.Children.Add(host);
    }

    private const string AudioEnhanceGenerateScript = """
import argparse
import os
import sys

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model-dir", required=False, default="")
    parser.add_argument("--variant", required=False, default="audiox_base")
    parser.add_argument("--prompt", required=True)
    parser.add_argument("--seconds", required=True, type=float)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    model_dir = args.model_dir or ""
    if model_dir:
        os.makedirs(model_dir, exist_ok=True)
        os.environ.setdefault("HF_HOME", model_dir)
        os.environ.setdefault("HUGGINGFACE_HUB_CACHE", model_dir)
        os.environ.setdefault("TRANSFORMERS_CACHE", model_dir)

    try:
        import torch
        import torchaudio
    except Exception as ex:
        print(f"ERR: missing torch/torchaudio: {ex}", file=sys.stderr)
        return 2

    model_loader = None
    generate_fn = None
    try:
        from audiox import get_pretrained_model
        from audiox.inference.generation import generate_diffusion_cond
        model_loader = get_pretrained_model
        generate_fn = generate_diffusion_cond
    except Exception as ex:
        print(f"ERR: AudioX import failed: {ex}", file=sys.stderr)
        return 3

    model_name = {
        "audiox_base": "HKUSTAudio/AudioX",
        "audiox_maf": "HKUSTAudio/AudioX-MAF",
        "audiox_maf_mmdit": "HKUSTAudio/AudioX-MAF-MMDiT",
    }.get((args.variant or "audiox_base").lower(), "HKUSTAudio/AudioX")

    try:
        model, _model_config = model_loader(model_name)
    except Exception as ex:
        print(f"ERR: failed loading model '{model_name}': {ex}", file=sys.stderr)
        return 4

    model.eval()
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model = model.to(device)

    requested_seconds = max(1.0, min(8.0, float(args.seconds)))
    model_seconds = 10.0
    sample_rate = int(getattr(model, "sample_rate", 44100))
    target_fps = 5
    if isinstance(_model_config, dict):
        try:
            target_fps = int(_model_config.get("video_fps", target_fps))
        except Exception:
            target_fps = 5

    video_tensors = torch.zeros((int(target_fps * model_seconds), 3, 224, 224), dtype=torch.float32, device=device)
    video_sync_frames = torch.zeros((1, 240, 768), dtype=torch.float32, device=device)
    audio_prompt = torch.zeros((1, 2, int(sample_rate * model_seconds)), dtype=torch.float32, device=device)
    conditioning = [{
        "video_prompt": {
            "video_tensors": video_tensors.unsqueeze(0),
            "video_sync_frames": video_sync_frames
        },
        "text_prompt": args.prompt or "",
        "audio_prompt": audio_prompt,
        "seconds_start": 0,
        "seconds_total": model_seconds
    }]

    try:
        output = generate_fn(
            model,
            steps=30,
            cfg_scale=7,
            conditioning=conditioning,
            sample_size=int(sample_rate * model_seconds),
            sigma_min=0.3,
            sigma_max=500,
            sampler_type="dpmpp-3m-sde",
            device=device,
        )
    except Exception as ex:
        print(f"ERR: generation failed: {ex}", file=sys.stderr)
        return 5

    if output is None:
        print("ERR: generation returned no output", file=sys.stderr)
        return 6

    output_dir = os.path.dirname(os.path.abspath(args.output))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    try:
        audio = output[0].to(torch.float32).detach().cpu()
        target_len = int(sample_rate * requested_seconds)
        if audio.shape[-1] > target_len:
            audio = audio[..., :target_len]
        elif audio.shape[-1] < target_len:
            pad = target_len - audio.shape[-1]
            audio = torch.nn.functional.pad(audio, (0, pad))
        if audio.dim() == 1:
            audio = audio.unsqueeze(0)
        torchaudio.save(args.output, audio, sample_rate)
    except Exception as ex:
        print(f"ERR: failed writing wav: {ex}", file=sys.stderr)
        return 7

    print("OK")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
""";

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

    private static Dictionary<string, SynthesisSettings> CloneProfiles(Dictionary<string, SynthesisSettings>? source)
    {
        var copy = new Dictionary<string, SynthesisSettings>(StringComparer.OrdinalIgnoreCase);
        if (source is null)
        {
            return copy;
        }

        foreach (var kv in source)
        {
            if (string.IsNullOrWhiteSpace(kv.Key) || kv.Value is null)
            {
                continue;
            }
            copy[kv.Key.Trim().ToLowerInvariant()] = CloneSynthesisSettings(kv.Value);
        }
        return copy;
    }
}
