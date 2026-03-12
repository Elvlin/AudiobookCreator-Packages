using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Input;
using System.Windows.Media;
using App.Core.Models;

namespace AudiobookCreator.UI;

public partial class ScriptPrepWindow : Window
{
    private readonly AppConfig _llmConfig;
    private readonly ScriptPrepLlmClient _llmClient;
    private readonly string _sourcePath;
    private readonly string _sourceText;
    private readonly string _sourceSignature;
    private readonly List<string> _voiceOptions;
    private readonly bool _allowInstructions;
    private readonly bool _allowManualSfx;
    private readonly ObservableCollection<PreparedScriptPartView> _parts = new();
    private readonly ObservableCollection<VoiceChoice> _voiceChoices = new();
    private readonly ObservableCollection<VoiceChoice> _recentVoiceChoices = new();
    private readonly Dictionary<string, TextBox> _partTextEditors = new(StringComparer.OrdinalIgnoreCase);
    private bool _isBusy;
    private readonly Stopwatch _llmBusyStopwatch = new();
    private TextBox? _lastFocusedPartTextEditor;
    private const int MaxRecentVoiceChoices = 12;

    public PreparedScriptDocument? Result { get; private set; }
    public IReadOnlyList<VoiceChoice> VoiceChoices => _voiceChoices;
    public IReadOnlyList<VoiceChoice> RecentVoiceChoices => _recentVoiceChoices;
    public Visibility InstructionEditorVisibility => _allowInstructions ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ManualSfxControlsVisibility => _allowManualSfx ? Visibility.Visible : Visibility.Collapsed;

    public ScriptPrepWindow(
        AppConfig llmConfig,
        string sourcePath,
        string sourceText,
        string sourceSignature,
        PreparedScriptDocument? existing,
        IEnumerable<string> voicePaths,
        bool allowInstructions)
    {
        _llmConfig = llmConfig;
        _llmClient = new ScriptPrepLlmClient(llmConfig);
        _sourcePath = sourcePath;
        _sourceText = sourceText;
        _sourceSignature = sourceSignature;
        _voiceOptions = voicePaths.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        _allowInstructions = allowInstructions;
        _allowManualSfx = AppFeatureFlags.AudioEnhancementEnabled;
        InitializeComponent();

        SourceFileText.Text = "Prepare Script: " + Path.GetFileName(sourcePath);
        ScriptStatusText.Text = existing is null
            ? "No prepared script yet."
            : $"Loaded {existing.Parts.Count} part(s) | Updated {existing.UpdatedAt.LocalDateTime:yyyy-MM-dd HH:mm:ss}";

        BuildVoiceCombo();
        ApplyInstructionModeUi();
        LoadParts(existing);
        PartsList.ItemsSource = _parts;
        SyncRecentVoiceChoicesFromParts();
        if (_parts.Count > 0)
        {
            PartsList.SelectedIndex = 0;
        }
        ConfigureLocalServerButton();
        UpdatePartButtonsState();
        UpdateAutoPrepareStatus();
    }

    private void ApplyInstructionModeUi()
    {
        if (RegenerateInstructionsButton is not null)
        {
            RegenerateInstructionsButton.Visibility = _allowInstructions ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void ConfigureLocalServerButton()
    {
        var isLocal = string.Equals((_llmConfig.LlmPrepProvider ?? "local").Trim(), "local", StringComparison.OrdinalIgnoreCase);
        if (LocalServerToggleButton is null)
        {
            return;
        }

        LocalServerToggleButton.Visibility = isLocal ? Visibility.Visible : Visibility.Collapsed;
        RefreshLocalServerButtonText();
    }

    private void RefreshLocalServerButtonText()
    {
        if (LocalServerToggleButton is null || LocalServerToggleButton.Visibility != Visibility.Visible)
        {
            return;
        }

        LocalServerToggleButton.Content = _llmClient.IsLocalServerRunning
            ? "Stop LLM Server"
            : "Start LLM Server";
    }

    private void BuildVoiceCombo()
    {
        _voiceChoices.Clear();
        _voiceChoices.Add(new VoiceChoice("(Use project default voice)", string.Empty));
        foreach (var path in _voiceOptions)
        {
            EnsureVoiceChoiceExists(path);
        }
    }

    private void EnsureVoiceChoiceExists(string? path)
    {
        var normalized = (path ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        if (_voiceChoices.Any(v => string.Equals(v.Path, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        _voiceChoices.Add(new VoiceChoice(GetVoiceDisplayName(normalized), normalized));
    }

    private void SyncRecentVoiceChoicesFromParts()
    {
        _recentVoiceChoices.Clear();
        foreach (var path in _parts
                     .Select(p => p.VoicePath)
                     .Where(p => !string.IsNullOrWhiteSpace(p))
                     .Reverse())
        {
            RegisterRecentVoice(path, updateExistingPosition: false);
        }
    }

    private void RegisterRecentVoice(string? path, bool updateExistingPosition = true)
    {
        var normalized = (path ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        EnsureVoiceChoiceExists(normalized);
        var existing = _recentVoiceChoices.FirstOrDefault(v =>
            string.Equals(v.Path, normalized, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            if (!updateExistingPosition)
            {
                return;
            }

            var existingIndex = _recentVoiceChoices.IndexOf(existing);
            if (existingIndex > 0)
            {
                _recentVoiceChoices.Move(existingIndex, 0);
            }
            return;
        }

        var display = GetVoiceDisplayName(normalized);
        _recentVoiceChoices.Insert(0, new VoiceChoice(display, normalized));
        while (_recentVoiceChoices.Count > MaxRecentVoiceChoices)
        {
            _recentVoiceChoices.RemoveAt(_recentVoiceChoices.Count - 1);
        }
        RefreshRecentVoiceChoicesForAllParts();
    }

    private static string GetVoiceDisplayName(string rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            return string.Empty;
        }

        var text = rawPath.Trim();
        const string customPrefix = "qwen-customvoice://";
        if (text.StartsWith(customPrefix, StringComparison.OrdinalIgnoreCase))
        {
            var speaker = text[customPrefix.Length..].Trim();
            return string.IsNullOrWhiteSpace(speaker) ? text : speaker;
        }

        return Path.GetFileName(text);
    }

    private void LoadParts(PreparedScriptDocument? existing)
    {
        _parts.Clear();
        var source = existing?.Parts is { Count: > 0 }
            ? existing.Parts.OrderBy(p => p.Order).ToList()
            : ScriptPrepEngine.ReSplit(_sourceText);

        for (var i = 0; i < source.Count; i++)
        {
            var part = source[i];
            EnsureVoiceChoiceExists(part.VoicePath);
            var view = new PreparedScriptPartView
            {
                Id = string.IsNullOrWhiteSpace(part.Id) ? Guid.NewGuid().ToString("N") : part.Id,
                Order = i,
                SpeakerTag = string.IsNullOrWhiteSpace(part.SpeakerTag) ? "Narrator" : part.SpeakerTag.Trim(),
                VoicePath = part.VoicePath ?? string.Empty,
                Instruction = part.Instruction ?? string.Empty,
                Text = part.Text ?? string.Empty,
                Locked = part.Locked,
                DisableAutoSfxDetection = _allowManualSfx && part.DisableAutoSfxDetection,
                SoundEffectPrompt = _allowManualSfx ? (part.SoundEffectPrompt ?? string.Empty) : string.Empty
            };
            WirePart(view);
            _parts.Add(view);
        }
        RefreshFooter();
    }

    private void RefreshFooter()
    {
        FooterText.Text = $"{_parts.Count} part(s). Empty parts are skipped during generation.";
        for (var i = 0; i < _parts.Count; i++)
        {
            _parts[i].Order = i;
        }
    }

    private void PartsList_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdatePartButtonsState();
    }

    private void AddPartButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        var insertAt = _parts.Count;
        var template = _parts.LastOrDefault();
        var defaultSpeaker = string.IsNullOrWhiteSpace(template?.SpeakerTag) ? "Narrator" : template!.SpeakerTag.Trim();
        var defaultVoice = template?.VoicePath ?? string.Empty;
        var defaultInstruction = template?.Instruction ?? string.Empty;

        var newPart = new PreparedScriptPartView
        {
            Id = Guid.NewGuid().ToString("N"),
            Order = insertAt,
            SpeakerTag = defaultSpeaker,
            VoicePath = defaultVoice,
            Instruction = defaultInstruction,
            Text = string.Empty,
            Locked = false,
            DisableAutoSfxDetection = false,
            SoundEffectPrompt = string.Empty
        };

        WirePart(newPart);
        _parts.Insert(insertAt, newPart);
        EnsureVoiceChoiceExists(newPart.VoicePath);
        RefreshFooter();
        PartsList.SelectedItem = newPart;
        UpdatePartButtonsState();
    }

    private void AddPartAfterButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || sender is not FrameworkElement { Tag: PreparedScriptPartView sourcePart })
        {
            return;
        }

        AddPartAfterSelected(sourcePart);
    }

    private void AddPartAfterSelected(PreparedScriptPartView sourcePart)
    {
        var sourceIndex = _parts.IndexOf(sourcePart);
        if (sourceIndex < 0)
        {
            return;
        }

        var insertAt = sourceIndex + 1;
        var newPart = new PreparedScriptPartView
        {
            Id = Guid.NewGuid().ToString("N"),
            Order = insertAt,
            SpeakerTag = string.IsNullOrWhiteSpace(sourcePart.SpeakerTag) ? "Narrator" : sourcePart.SpeakerTag.Trim(),
            VoicePath = sourcePart.VoicePath ?? string.Empty,
            Instruction = _allowInstructions ? (sourcePart.Instruction ?? string.Empty) : string.Empty,
            Text = string.Empty,
            Locked = false,
            DisableAutoSfxDetection = false,
            SoundEffectPrompt = string.Empty
        };

        WirePart(newPart);
        _parts.Insert(insertAt, newPart);
        EnsureVoiceChoiceExists(newPart.VoicePath);
        RefreshFooter();
        PartsList.SelectedItem = newPart;
        PartsList.ScrollIntoView(newPart);
        UpdatePartButtonsState();
    }

    private void SplitPartButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || sender is not FrameworkElement { Tag: PreparedScriptPartView sourcePart })
        {
            return;
        }

        TrySplitPartAtCaret(sourcePart);
    }

    private void PartTextEditor_OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_isBusy || e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.Control)
        {
            return;
        }

        if (sender is not TextBox { Tag: PreparedScriptPartView sourcePart })
        {
            return;
        }

        if (TrySplitPartAtCaret(sourcePart))
        {
            e.Handled = true;
        }
    }

    private bool TrySplitPartAtCaret(PreparedScriptPartView sourcePart)
    {
        if (_isBusy)
        {
            return false;
        }

        var sourceIndex = _parts.IndexOf(sourcePart);
        if (sourceIndex < 0)
        {
            return false;
        }

        var editor = GetEditorForPart(sourcePart);
        if (editor is null)
        {
            MessageBox.Show(this,
                "Click inside the part text first, place the caret where you want the split, then click Split Here.",
                "Prepare Script",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }

        var currentText = editor.Text ?? string.Empty;
        var caretIndex = Math.Clamp(editor.CaretIndex, 0, currentText.Length);
        if (caretIndex <= 0 || caretIndex >= currentText.Length)
        {
            MessageBox.Show(this,
                "Place the text cursor inside the part before splitting. The split point cannot be at the very start or end.",
                "Prepare Script",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }

        var leftText = currentText[..caretIndex].TrimEnd();
        var rightText = currentText[caretIndex..].TrimStart();
        if (string.IsNullOrWhiteSpace(leftText) || string.IsNullOrWhiteSpace(rightText))
        {
            MessageBox.Show(this,
                "That split point would leave one side empty. Move the caret to a point where both parts contain text.",
                "Prepare Script",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }

        sourcePart.Text = leftText;
        var insertAt = sourceIndex + 1;
        var newPart = new PreparedScriptPartView
        {
            Id = Guid.NewGuid().ToString("N"),
            Order = insertAt,
            SpeakerTag = string.IsNullOrWhiteSpace(sourcePart.SpeakerTag) ? "Narrator" : sourcePart.SpeakerTag.Trim(),
            VoicePath = sourcePart.VoicePath ?? string.Empty,
            Instruction = _allowInstructions ? (sourcePart.Instruction ?? string.Empty) : string.Empty,
            Text = rightText,
            Locked = _allowInstructions && sourcePart.Locked,
            DisableAutoSfxDetection = _allowManualSfx && sourcePart.DisableAutoSfxDetection,
            SoundEffectPrompt = _allowManualSfx && sourcePart.DisableAutoSfxDetection ? (sourcePart.SoundEffectPrompt ?? string.Empty) : string.Empty
        };

        WirePart(newPart);
        _parts.Insert(insertAt, newPart);
        EnsureVoiceChoiceExists(newPart.VoicePath);
        RefreshFooter();
        PartsList.SelectedItem = newPart;
        PartsList.ScrollIntoView(newPart);
        UpdatePartButtonsState();

        Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
        {
            if (GetEditorForPart(newPart) is { } newEditor)
            {
                newEditor.Focus();
                newEditor.CaretIndex = 0;
            }
        }));
        return true;
    }

    private void DeletePartButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || PartsList.SelectedItem is not PreparedScriptPartView selected)
        {
            return;
        }

        if (_parts.Count <= 1)
        {
            MessageBox.Show(this, "At least one part must remain.", "Prepare Script", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var index = _parts.IndexOf(selected);
        if (index < 0)
        {
            return;
        }

        UnwirePart(selected);
        _parts.RemoveAt(index);
        RefreshFooter();
        var nextIndex = Math.Clamp(index, 0, _parts.Count - 1);
        if (_parts.Count > 0)
        {
            PartsList.SelectedIndex = nextIndex;
            PartsList.ScrollIntoView(_parts[nextIndex]);
        }
        UpdatePartButtonsState();
    }

    private void DeletePartForItemButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || sender is not FrameworkElement { Tag: PreparedScriptPartView part })
        {
            return;
        }

        if (_parts.Count <= 1)
        {
            MessageBox.Show(this, "At least one part must remain.", "Prepare Script", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var index = _parts.IndexOf(part);
        if (index < 0)
        {
            return;
        }

        UnwirePart(part);
        _parts.RemoveAt(index);
        RefreshFooter();
        if (_parts.Count > 0)
        {
            var nextIndex = Math.Clamp(index, 0, _parts.Count - 1);
            PartsList.SelectedIndex = nextIndex;
            PartsList.ScrollIntoView(_parts[nextIndex]);
        }
        UpdatePartButtonsState();
    }

    private void PartTextEditor_OnLoaded(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.Tag is PreparedScriptPartView part)
        {
            _partTextEditors[part.Id] = textBox;
        }
    }

    private void PartTextEditor_OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.Tag is PreparedScriptPartView part)
        {
            if (_partTextEditors.TryGetValue(part.Id, out var existing) && ReferenceEquals(existing, textBox))
            {
                _partTextEditors.Remove(part.Id);
            }

            if (ReferenceEquals(_lastFocusedPartTextEditor, textBox))
            {
                _lastFocusedPartTextEditor = null;
            }
        }
    }

    private void PartTextEditor_OnGotKeyboardFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            _lastFocusedPartTextEditor = textBox;
            if (textBox.Tag is PreparedScriptPartView part)
            {
                PartsList.SelectedItem = part;
            }
        }
    }

    private void VoiceComboForPart_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is not ComboBox { Tag: PreparedScriptPartView part })
        {
            return;
        }

        if (sender is ComboBox combo)
        {
            var selectedPath = combo.SelectedValue as string ?? part.VoicePath;
            if (!string.IsNullOrWhiteSpace(selectedPath))
            {
                RegisterRecentVoice(selectedPath);
            }
            else
            {
                RefreshRecentVoiceChoicesForPart(part);
            }
        }
    }

    private void RecentVoiceButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || sender is not Button { Tag: PreparedScriptPartView part, CommandParameter: RecentVoiceChipView choice })
        {
            return;
        }

        part.VoicePath = choice.Path;
        RegisterRecentVoice(choice.Path);
    }

    private void InspectorVoiceCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PartsList.SelectedItem is not PreparedScriptPartView part)
        {
            return;
        }

        var selectedPath = part.VoicePath;
        if (!string.IsNullOrWhiteSpace(selectedPath))
        {
            RegisterRecentVoice(selectedPath);
        }
        else
        {
            RefreshRecentVoiceChoicesForPart(part);
        }
    }

    private void InspectorRecentVoiceButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || PartsList.SelectedItem is not PreparedScriptPartView part || sender is not Button { CommandParameter: RecentVoiceChipView choice })
        {
            return;
        }

        part.VoicePath = choice.Path;
        RegisterRecentVoice(choice.Path);
    }

    private void InspectorAddPartButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || PartsList.SelectedItem is not PreparedScriptPartView part)
        {
            return;
        }

        AddPartAfterSelected(part);
    }

    private void InspectorSplitButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy || PartsList.SelectedItem is not PreparedScriptPartView part)
        {
            return;
        }

        TrySplitPartAtCaret(part);
    }

    private void WirePart(PreparedScriptPartView part)
    {
        part.PropertyChanged += PreparedPartView_OnPropertyChanged;
        RefreshRecentVoiceChoicesForPart(part);
    }

    private void UnwirePart(PreparedScriptPartView part)
    {
        part.PropertyChanged -= PreparedPartView_OnPropertyChanged;
        _partTextEditors.Remove(part.Id);
    }

    private void PreparedPartView_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not PreparedScriptPartView part)
        {
            return;
        }

        if (string.Equals(e.PropertyName, nameof(PreparedScriptPartView.VoicePath), StringComparison.Ordinal))
        {
            RefreshRecentVoiceChoicesForPart(part);
        }
    }

    private void RefreshRecentVoiceChoicesForAllParts()
    {
        foreach (var part in _parts)
        {
            RefreshRecentVoiceChoicesForPart(part);
        }
    }

    private void RefreshRecentVoiceChoicesForPart(PreparedScriptPartView part)
    {
        part.RecentVoiceChoices.Clear();
        foreach (var recent in _recentVoiceChoices)
        {
            part.RecentVoiceChoices.Add(new RecentVoiceChipView(
                recent.Display,
                recent.Path,
                string.Equals(recent.Path, part.VoicePath, StringComparison.OrdinalIgnoreCase)));
        }
    }

    private TextBox? GetEditorForPart(PreparedScriptPartView part)
    {
        if (_lastFocusedPartTextEditor?.Tag is PreparedScriptPartView focusedPart &&
            string.Equals(focusedPart.Id, part.Id, StringComparison.OrdinalIgnoreCase))
        {
            return _lastFocusedPartTextEditor;
        }

        return _partTextEditors.TryGetValue(part.Id, out var editor) ? editor : null;
    }

    private void UpdatePartButtonsState()
    {
        if (AddPartButton is not null)
        {
            AddPartButton.IsEnabled = !_isBusy;
        }

        if (DeletePartButton is not null)
        {
            DeletePartButton.IsEnabled = !_isBusy && _parts.Count > 1;
        }
    }

    private async void AutoPrepareButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }
        if (!await EnsureLocalServerReadyForActionAsync("Auto Prepare (LLM)"))
        {
            return;
        }

        await RunBusyAsync("Auto preparing with LLM...", async () =>
        {
            var progress = BuildLlmProgressAdapter();
            var parts = _allowInstructions
                ? await _llmClient.AutoPrepareAsync(_sourceText, default, progress)
                : await _llmClient.SplitAsync(_sourceText, default, progress);
            ReplaceParts(parts);
            ShowSplitFallbackWarningIfNeeded("Auto Prepare");
            UpdateAutoPrepareStatus(prefix: "Auto-prepared");
        });
    }

    private async void ResplitButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }
        if (!await EnsureLocalServerReadyForActionAsync("Re-split"))
        {
            return;
        }

        await RunBusyAsync("Re-splitting with LLM...", async () =>
        {
            var progress = BuildLlmProgressAdapter();
            var split = await _llmClient.SplitAsync(_sourceText, default, progress);
            var existingBySpeaker = _parts
                .Where(p => !string.IsNullOrWhiteSpace(p.SpeakerTag) && !string.IsNullOrWhiteSpace(p.VoicePath))
                .GroupBy(p => p.SpeakerTag.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().VoicePath, StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < split.Count; i++)
            {
                if (existingBySpeaker.TryGetValue(split[i].SpeakerTag, out var mappedVoice))
                {
                    split[i].VoicePath = mappedVoice;
                }
            }
            ReplaceParts(split);
            ShowSplitFallbackWarningIfNeeded("Re-split");
            UpdateAutoPrepareStatus(prefix: "Re-split complete");
        });
    }

    private async void RegenerateInstructionsButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_allowInstructions)
        {
            return;
        }

        if (_isBusy)
        {
            return;
        }
        if (!await EnsureLocalServerReadyForActionAsync("Regenerate Instructions"))
        {
            return;
        }

        await RunBusyAsync("Regenerating instructions with LLM...", async () =>
        {
            var current = _parts.Select(p => p.ToModel()).ToList();
            var progress = BuildLlmProgressAdapter();
            var regenerated = await _llmClient.GenerateInstructionsAsync(current, default, progress);
            for (var i = 0; i < _parts.Count && i < regenerated.Count; i++)
            {
                if (_parts[i].Locked)
                {
                    continue;
                }
                _parts[i].Instruction = regenerated[i].Instruction;
            }
            UpdateAutoPrepareStatus(prefix: "Instruction regeneration complete");
        });
    }

    private async void LocalServerToggleButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        if (_llmClient.IsLocalServerRunning)
        {
            await RunBusyAsync("Stopping local LLM server...", async () =>
            {
                await _llmClient.StopLocalServerAsync();
                UpdateAutoPrepareStatus(prefix: "Local server stopped");
            });
        }
        else
        {
            await RunBusyAsync("Starting local LLM server...", async () =>
            {
                await _llmClient.StartLocalServerAsync(default);
                UpdateAutoPrepareStatus(prefix: "Local server started");
            });
        }
        RefreshLocalServerButtonText();
    }

    private bool IsLocalLlmProvider()
    {
        return string.Equals((_llmConfig.LlmPrepProvider ?? "local").Trim(), "local", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<bool> EnsureLocalServerReadyForActionAsync(string actionName)
    {
        if (!IsLocalLlmProvider() || _llmClient.IsLocalServerRunning)
        {
            return true;
        }

        var answer = MessageBox.Show(
            this,
            $"{actionName} needs local LLM server, but it is not running.\n\nStart LLM server now?",
            "LLM Server Required",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes)
        {
            ScriptStatusText.Text = $"{actionName} canceled (local LLM server not started).";
            return false;
        }

        await RunBusyAsync("Starting local LLM server...", async () =>
        {
            await _llmClient.StartLocalServerAsync(default);
            UpdateAutoPrepareStatus(prefix: "Local server started");
        });
        RefreshLocalServerButtonText();

        if (!_llmClient.IsLocalServerRunning)
        {
            MessageBox.Show(this,
                "Local LLM server is still not running. Check runtime/model path in Settings.",
                "LLM Server",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        PersistResult();
    }

    private void SaveCloseButton_OnClick(object sender, RoutedEventArgs e)
    {
        PersistResult();
        DialogResult = true;
        Close();
    }

    private void PersistResult()
    {
        var valid = _parts
            .Where(p => !string.IsNullOrWhiteSpace((p.Text ?? string.Empty).Trim()))
            .Select((p, idx) =>
            {
                var model = p.ToModel();
                model.Order = idx;
                model.Text = model.Text.Trim();
                model.Instruction = _allowInstructions ? TrimInstruction(model.Instruction) : string.Empty;
                model.SpeakerTag = string.IsNullOrWhiteSpace(model.SpeakerTag) ? "Narrator" : model.SpeakerTag.Trim();
                model.VoicePath = string.IsNullOrWhiteSpace(model.VoicePath) ? null : model.VoicePath.Trim();
                model.Locked = _allowInstructions && model.Locked;
                model.DisableAutoSfxDetection = _allowManualSfx && p.DisableAutoSfxDetection;
                model.SoundEffectPrompt = _allowManualSfx && p.DisableAutoSfxDetection
                    ? (p.SoundEffectPrompt ?? string.Empty).Trim()
                    : string.Empty;
                return model;
            })
            .ToList();

        Result = new PreparedScriptDocument
        {
            SourcePath = _sourcePath,
            SourceSignature = _sourceSignature,
            UpdatedAt = DateTimeOffset.UtcNow,
            Version = "v1",
            Parts = valid
        };
        ScriptStatusText.Text = $"Saved {valid.Count} part(s).";
    }

    private void ReplaceParts(IReadOnlyList<PreparedScriptPart> source)
    {
        foreach (var existing in _parts.ToList())
        {
            UnwirePart(existing);
        }
        _parts.Clear();
        for (var i = 0; i < source.Count; i++)
        {
            var part = source[i];
            EnsureVoiceChoiceExists(part.VoicePath);
            var view = new PreparedScriptPartView
            {
                Id = string.IsNullOrWhiteSpace(part.Id) ? Guid.NewGuid().ToString("N") : part.Id,
                Order = i,
                SpeakerTag = string.IsNullOrWhiteSpace(part.SpeakerTag) ? "Narrator" : part.SpeakerTag.Trim(),
                VoicePath = part.VoicePath ?? string.Empty,
                Instruction = _allowInstructions ? (part.Instruction ?? string.Empty) : string.Empty,
                Text = part.Text ?? string.Empty,
                Locked = _allowInstructions && part.Locked,
                DisableAutoSfxDetection = _allowManualSfx && part.DisableAutoSfxDetection,
                SoundEffectPrompt = _allowManualSfx ? (part.SoundEffectPrompt ?? string.Empty) : string.Empty
            };
            WirePart(view);
            _parts.Add(view);
        }
        SyncRecentVoiceChoicesFromParts();
        RefreshFooter();
        if (_parts.Count > 0)
        {
            PartsList.SelectedIndex = 0;
        }
    }

    private async Task RunBusyAsync(string status, Func<Task> action)
    {
        SetBusy(true, status);
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Script Prep", ex.Message),
                "Script Prep Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false, string.Empty);
        }
    }

    private void SetBusy(bool isBusy, string status)
    {
        _isBusy = isBusy;
        AutoPrepareButton.IsEnabled = !isBusy;
        ResplitButton.IsEnabled = !isBusy;
        RegenerateInstructionsButton.IsEnabled = !isBusy;
        if (LocalServerToggleButton is not null)
        {
            LocalServerToggleButton.IsEnabled = !isBusy;
        }
        SaveButton.IsEnabled = !isBusy;
        SaveCloseButton.IsEnabled = !isBusy;
        PartsList.IsEnabled = !isBusy;
        if (AddPartButton is not null)
        {
            AddPartButton.IsEnabled = !isBusy;
        }
        if (DeletePartButton is not null)
        {
            DeletePartButton.IsEnabled = !isBusy && PartsList.SelectedItem is PreparedScriptPartView && _parts.Count > 1;
        }
        if (PartEditorBorder is not null)
        {
            PartEditorBorder.IsEnabled = !isBusy;
        }
        if (isBusy)
        {
            _llmBusyStopwatch.Restart();
            if (LlmProgressPanel is not null)
            {
                LlmProgressPanel.Visibility = Visibility.Visible;
            }
            if (LlmProgressBar is not null)
            {
                LlmProgressBar.IsIndeterminate = true;
                LlmProgressBar.Value = 0;
            }
            if (LlmProgressText is not null)
            {
                LlmProgressText.Text = status;
            }
            ScriptStatusText.Text = status;
        }
        else
        {
            _llmBusyStopwatch.Stop();
            if (LlmProgressPanel is not null)
            {
                LlmProgressPanel.Visibility = Visibility.Collapsed;
            }
        }
        UpdatePartButtonsState();
    }

    private void UpdateAutoPrepareStatus(string prefix = "Ready")
    {
        var provider = (_llmConfig.LlmPrepProvider ?? "local").Trim().ToLowerInvariant();
        var splitModel = (_llmConfig.LlmPrepSplitModel ?? string.Empty).Trim();
        var instructModel = (_llmConfig.LlmPrepInstructionModel ?? string.Empty).Trim();
        if (!_allowInstructions)
        {
            ScriptStatusText.Text = $"{prefix}: provider={provider}, split={splitModel} (instruction disabled for this model).";
            return;
        }
        ScriptStatusText.Text = _llmConfig.LlmPrepUseSeparateModels
            ? $"{prefix}: provider={provider}, split={splitModel}, instruction={instructModel}."
            : $"{prefix}: provider={provider}, model={splitModel}.";
    }

    private void ShowSplitFallbackWarningIfNeeded(string actionLabel)
    {
        if (!_llmClient.LastSplitUsedFallback)
        {
            return;
        }

        var reason = string.IsNullOrWhiteSpace(_llmClient.LastSplitFallbackReason)
            ? "LLM split response was invalid."
            : _llmClient.LastSplitFallbackReason;
        ScriptStatusText.Text = $"{actionLabel}: LLM split failed; deterministic fallback used.";
        MessageBox.Show(
            this,
            $"{actionLabel} could not use LLM split output.\n" +
            "Used deterministic fallback splitter instead.\n\n" +
            $"Reason: {reason}",
            "LLM Split Fallback",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private static string TrimInstruction(string value)
    {
        var cleaned = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        if (cleaned.Length <= 220)
        {
            return cleaned;
        }
        return cleaned[..220].Trim();
    }

    private sealed class PreparedScriptPartView : INotifyPropertyChanged
    {
        private static readonly Brush[] SpeakerAccentPalette =
        {
            CreateBrush("#5C846B"),
            CreateBrush("#5A7395"),
            CreateBrush("#9A6B54"),
            CreateBrush("#7A5D96"),
            CreateBrush("#8A6A46"),
            CreateBrush("#4E8A88")
        };
        private string _speakerTag = "Narrator";
        private string _voicePath = string.Empty;
        private string _instruction = string.Empty;
        private string _text = string.Empty;
        private bool _locked;
        private bool _disableAutoSfxDetection;
        private string _soundEffectPrompt = string.Empty;
        private int _order;

        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public int Order
        {
            get => _order;
            set
            {
                if (SetField(ref _order, value))
                {
                    OnPropertyChanged(nameof(DisplayLabel));
                }
            }
        }

        public string SpeakerTag
        {
            get => _speakerTag;
            set
            {
                if (SetField(ref _speakerTag, value))
                {
                    OnPropertyChanged(nameof(DisplayLabel));
                    OnPropertyChanged(nameof(SpeakerAccentBrush));
                }
            }
        }

        public string VoicePath
        {
            get => _voicePath;
            set
            {
                if (SetField(ref _voicePath, value))
                {
                    OnPropertyChanged(nameof(VoiceDisplayLabel));
                }
            }
        }

        public string Instruction
        {
            get => _instruction;
            set => SetField(ref _instruction, value);
        }

        public string Text
        {
            get => _text;
            set
            {
                if (SetField(ref _text, value))
                {
                    OnPropertyChanged(nameof(DisplayLabel));
                }
            }
        }

        public bool Locked
        {
            get => _locked;
            set => SetField(ref _locked, value);
        }

        public bool DisableAutoSfxDetection
        {
            get => _disableAutoSfxDetection;
            set
            {
                if (SetField(ref _disableAutoSfxDetection, value))
                {
                    OnPropertyChanged(nameof(ManualSfxEditorVisibility));
                    if (!value && !string.IsNullOrWhiteSpace(_soundEffectPrompt))
                    {
                        _soundEffectPrompt = string.Empty;
                        OnPropertyChanged(nameof(SoundEffectPrompt));
                    }
                }
            }
        }

        public string SoundEffectPrompt
        {
            get => _soundEffectPrompt;
            set => SetField(ref _soundEffectPrompt, value);
        }

        public ObservableCollection<RecentVoiceChipView> RecentVoiceChoices { get; } = new();
        public Visibility ManualSfxEditorVisibility => DisableAutoSfxDetection ? Visibility.Visible : Visibility.Collapsed;

        public Brush SpeakerAccentBrush
        {
            get
            {
                var key = string.IsNullOrWhiteSpace(_speakerTag) ? "Narrator" : _speakerTag.Trim();
                var index = Math.Abs(StringComparer.OrdinalIgnoreCase.GetHashCode(key)) % SpeakerAccentPalette.Length;
                return SpeakerAccentPalette[index];
            }
        }

        public string DisplayLabel
        {
            get
            {
                var preview = (_text ?? string.Empty).Trim();
                if (preview.Length > 72)
                {
                    preview = preview[..72] + "...";
                }
                var speaker = string.IsNullOrWhiteSpace(_speakerTag) ? "Narrator" : _speakerTag.Trim();
                return $"{Order + 1:00}. [{speaker}] {preview}";
            }
        }

        public string VoiceDisplayLabel
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_voicePath))
                {
                    return "Default Voice";
                }

                return ScriptPrepWindow.GetVoiceDisplayName(_voicePath);
            }
        }

        public PreparedScriptPart ToModel() => new()
        {
            Id = Id,
            Order = Order,
            SpeakerTag = SpeakerTag,
            VoicePath = string.IsNullOrWhiteSpace(VoicePath) ? null : VoicePath,
            Instruction = Instruction,
            Text = Text,
            Locked = Locked,
            DisableAutoSfxDetection = DisableAutoSfxDetection,
            SoundEffectPrompt = DisableAutoSfxDetection ? (SoundEffectPrompt ?? string.Empty) : string.Empty
        };

        public event PropertyChangedEventHandler? PropertyChanged;

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

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static SolidColorBrush CreateBrush(string hex)
        {
            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(hex)!;
            brush.Freeze();
            return brush;
        }
    }

    public sealed record VoiceChoice(string Display, string Path)
    {
        public override string ToString() => Display;
    }

    public sealed record RecentVoiceChipView(string Display, string Path, bool IsSelected);

    private IProgress<ScriptPrepLlmClient.ScriptPrepProgress> BuildLlmProgressAdapter()
    {
        return new Progress<ScriptPrepLlmClient.ScriptPrepProgress>(p =>
        {
            var total = Math.Max(1, p.Total);
            var completed = Math.Clamp(p.Completed, 0, total);
            var percent = Math.Clamp((completed * 100.0) / total, 0.0, 100.0);
            var elapsed = _llmBusyStopwatch.Elapsed;
            var elapsedText = $"{Math.Min(99, (int)elapsed.TotalMinutes):00}:{elapsed.Seconds:00}";

            if (LlmProgressBar is not null)
            {
                LlmProgressBar.IsIndeterminate = false;
                LlmProgressBar.Value = percent;
            }

            if (LlmProgressText is not null)
            {
                var stage = string.IsNullOrWhiteSpace(p.Stage) ? "Processing..." : p.Stage.Trim();
                LlmProgressText.Text = $"{stage}  ({completed}/{total}, {percent:0}%)  elapsed {elapsedText}";
            }
        });
    }

    protected override async void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        if (_llmClient.IsLocalServerRunning)
        {
            try { await _llmClient.StopLocalServerAsync(); } catch { }
        }
    }
}
