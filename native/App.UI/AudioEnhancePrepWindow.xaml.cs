using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Shapes = System.Windows.Shapes;
using Microsoft.Win32;
using NAudio.Wave;
using App.Core.Models;

namespace AudiobookCreator.UI;

public partial class AudioEnhancePrepWindow : Window
{
    private readonly string _sourcePath;
    private readonly ObservableCollection<SegmentRow> _segments = new();
    private readonly List<float> _wavePeaks = new();
    private double _durationSeconds;
    private bool _updatingRows;
    private WaveOutEvent? _waveOut;
    private AudioFileReader? _playbackReader;
    private readonly DispatcherTimer _playbackTimer;
    private DragPreviewState? _dragPreview;

    public List<EnhancePreparedSegment> ResultSegments { get; private set; } = new();

    public AudioEnhancePrepWindow(string sourcePath, IReadOnlyList<EnhancePreparedSegment>? initialSegments)
    {
        _sourcePath = sourcePath;
        InitializeComponent();
        SegmentsGrid.ItemsSource = _segments;
        _segments.CollectionChanged += (_, _) =>
        {
            Reindex();
            RenderWaveform();
        };
        _playbackTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(80) };
        _playbackTimer.Tick += PlaybackTimer_OnTick;
        Closed += (_, _) => StopPlayback();

        SourceFileText.Text = "Prepare Enhancement: " + Path.GetFileName(sourcePath);
        LoadWaveform(sourcePath);
        UpdatePlaybackTimeText();
        LoadInitialSegments(initialSegments);
    }

    private void LoadInitialSegments(IReadOnlyList<EnhancePreparedSegment>? initialSegments)
    {
        var source = (initialSegments ?? Array.Empty<EnhancePreparedSegment>())
            .OrderBy(s => s.Order)
            .ThenBy(s => s.StartSeconds)
            .ToList();

        if (source.Count == 0 && TryFindSubtitleSidecar(_sourcePath, out var subtitlePath))
        {
            source = ParseSubtitleFileToPreparedSegments(subtitlePath);
            StatusText.Text = $"Auto-loaded subtitle: {Path.GetFileName(subtitlePath)}";
        }

        if (source.Count == 0)
        {
            source.Add(new EnhancePreparedSegment
            {
                Id = Guid.NewGuid().ToString("N"),
                Order = 1,
                StartSeconds = 0,
                EndSeconds = Math.Min(Math.Max(_durationSeconds, 10), 6),
                OneShotSeconds = 1.5,
                Intensity = 0.6,
                Enabled = true
            });
        }

        _updatingRows = true;
        try
        {
            _segments.Clear();
            foreach (var seg in source)
            {
                var row = new SegmentRow(seg);
                row.PropertyChanged += Segment_OnPropertyChanged;
                _segments.Add(row);
            }
            Reindex();
            NormalizeRows();
        }
        finally
        {
            _updatingRows = false;
        }

        RenderWaveform();
    }

    private void Segment_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_updatingRows)
        {
            return;
        }
        RenderWaveform();
    }

    private void LoadWaveform(string path)
    {
        _wavePeaks.Clear();
        _durationSeconds = 0;

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return;
        }

        try
        {
            using var reader = new AudioFileReader(path);
            _durationSeconds = Math.Max(0.0, reader.TotalTime.TotalSeconds);
            var channels = Math.Max(1, reader.WaveFormat.Channels);
            var targetBins = 1600;
            var bins = new float[targetBins];
            var totalFrames = Math.Max(1L, reader.Length / Math.Max(1, reader.WaveFormat.BlockAlign));

            var buffer = new float[4096 * channels];
            long frameIndex = 0;
            int read;
            while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (var i = 0; i + channels - 1 < read; i += channels)
                {
                    var mono = 0.0f;
                    for (var ch = 0; ch < channels; ch++)
                    {
                        mono += Math.Abs(buffer[i + ch]);
                    }
                    mono /= channels;
                    var bin = (int)Math.Min(targetBins - 1, (frameIndex * targetBins) / (double)totalFrames);
                    if (mono > bins[bin])
                    {
                        bins[bin] = mono;
                    }
                    frameIndex++;
                }
            }

            var max = bins.Length == 0 ? 0f : bins.Max();
            if (max > 0)
            {
                for (var i = 0; i < bins.Length; i++)
                {
                    bins[i] = Math.Clamp(bins[i] / max, 0f, 1f);
                }
            }
            _wavePeaks.AddRange(bins);
        }
        catch
        {
            // keep empty waveform
        }
    }

    private void RenderWaveform()
    {
        if (WaveformCanvas is null)
        {
            return;
        }

        WaveformCanvas.Children.Clear();
        var duration = Math.Max(_durationSeconds, 0.1);
        var pps = GetPixelsPerSecond();
        var width = Math.Max(900, duration * pps);
        var height = Math.Max(220, WaveformCanvas.Height);
        const double rulerHeight = 28;
        const double laneHeight = 54;
        const double outerPadding = 8;
        var laneTop = height - laneHeight - outerPadding;
        var waveformTop = rulerHeight + 12;
        var waveformBottom = laneTop - 12;
        var waveformCenterY = waveformTop + ((waveformBottom - waveformTop) / 2.0);
        WaveformCanvas.Width = width;

        var rulerBand = new Border
        {
            Width = width,
            Height = rulerHeight,
            Background = new SolidColorBrush(Color.FromRgb(245, 248, 246)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(221, 229, 224)),
            BorderThickness = new Thickness(0, 0, 0, 1)
        };
        Canvas.SetLeft(rulerBand, 0);
        Canvas.SetTop(rulerBand, 0);
        WaveformCanvas.Children.Add(rulerBand);

        var waveformBaseline = new Shapes.Line
        {
            X1 = 0,
            X2 = width,
            Y1 = waveformCenterY,
            Y2 = waveformCenterY,
            Stroke = new SolidColorBrush(Color.FromArgb(120, 164, 180, 171)),
            StrokeThickness = 1
        };
        WaveformCanvas.Children.Add(waveformBaseline);

        var gridStroke = new SolidColorBrush(Color.FromRgb(236, 241, 238));
        var labelEverySeconds = duration > 180 ? 15 : duration > 90 ? 10 : 5;
        for (var sec = 0; sec <= Math.Ceiling(duration); sec++)
        {
            var x = sec * pps;
            var isMajor = sec % labelEverySeconds == 0;
            var line = new System.Windows.Shapes.Line
            {
                X1 = x, X2 = x, Y1 = rulerHeight, Y2 = height,
                Stroke = gridStroke,
                StrokeThickness = isMajor ? 1.2 : 0.6
            };
            WaveformCanvas.Children.Add(line);

            var rulerTick = new Shapes.Line
            {
                X1 = x,
                X2 = x,
                Y1 = rulerHeight - (isMajor ? 11 : 7),
                Y2 = rulerHeight,
                Stroke = new SolidColorBrush(Color.FromRgb(142, 153, 147)),
                StrokeThickness = isMajor ? 1.2 : 1
            };
            WaveformCanvas.Children.Add(rulerTick);

            if (sec < Math.Ceiling(duration) && isMajor)
            {
                var label = new TextBlock
                {
                    Text = FormatTimelineTime(sec),
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(110, 120, 116))
                };
                Canvas.SetLeft(label, x + 2);
                Canvas.SetTop(label, 4);
                WaveformCanvas.Children.Add(label);
            }
        }

        var waveformStroke = new SolidColorBrush(Color.FromRgb(99, 136, 117));
        var bins = _wavePeaks.Count;
        if (bins > 0)
        {
            for (var i = 0; i < bins; i++)
            {
                var x = i * (width / Math.Max(1, bins - 1));
                var amp = Math.Clamp(_wavePeaks[i], 0f, 1f) * (float)((waveformBottom - waveformTop) * 0.48);
                var bar = new System.Windows.Shapes.Line
                {
                    X1 = x, X2 = x,
                    Y1 = waveformCenterY - amp, Y2 = waveformCenterY + amp,
                    Stroke = waveformStroke,
                    StrokeThickness = 1
                };
                WaveformCanvas.Children.Add(bar);
            }
        }

        var laneBackground = new Border
        {
            Width = width,
            Height = laneHeight,
            Background = new SolidColorBrush(Color.FromRgb(244, 247, 250)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(220, 228, 223)),
            BorderThickness = new Thickness(1)
        };
        Canvas.SetLeft(laneBackground, 0);
        Canvas.SetTop(laneBackground, laneTop);
        WaveformCanvas.Children.Add(laneBackground);

        var selected = SegmentsGrid.SelectedItem as SegmentRow;
        var ambienceBrush = new SolidColorBrush(Color.FromArgb(105, 124, 170, 223));
        var selectedBrush = new SolidColorBrush(Color.FromArgb(150, 67, 126, 194));
        foreach (var seg in _segments.OrderBy(s => s.Order))
        {
            if (!seg.Enabled)
            {
                continue;
            }

            var start = Math.Clamp(GetRenderedStart(seg), 0, duration);
            var end = Math.Clamp(GetRenderedEnd(seg), 0, duration);
            if (end <= start)
            {
                continue;
            }

            var x = start * pps;
            var segWidth = Math.Max(6, (end - start) * pps);
            var region = new Border
            {
                Width = segWidth,
                Height = laneHeight - 10,
                Background = ReferenceEquals(seg, selected) ? selectedBrush : ambienceBrush,
                BorderBrush = new SolidColorBrush(Color.FromArgb(160, 39, 92, 148)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Opacity = 0.78
            };
            Canvas.SetLeft(region, x);
            Canvas.SetTop(region, laneTop + 5);
            WaveformCanvas.Children.Add(region);

            var leftGripVisual = new Border
            {
                Width = 8,
                Height = laneHeight - 10,
                Background = new SolidColorBrush(Color.FromArgb(220, 37, 83, 124)),
                BorderBrush = new SolidColorBrush(Color.FromArgb(235, 24, 61, 94)),
                BorderThickness = new Thickness(1, 0, 1, 0)
            };
            Canvas.SetLeft(leftGripVisual, x);
            Canvas.SetTop(leftGripVisual, laneTop + 5);
            WaveformCanvas.Children.Add(leftGripVisual);

            var rightGripVisual = new Border
            {
                Width = 8,
                Height = laneHeight - 10,
                Background = new SolidColorBrush(Color.FromArgb(220, 37, 83, 124)),
                BorderBrush = new SolidColorBrush(Color.FromArgb(235, 24, 61, 94)),
                BorderThickness = new Thickness(1, 0, 1, 0)
            };
            Canvas.SetLeft(rightGripVisual, x + segWidth - 8);
            Canvas.SetTop(rightGripVisual, laneTop + 5);
            WaveformCanvas.Children.Add(rightGripVisual);

            var label = new TextBlock
            {
                Text = $"{seg.Order:00}",
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                FontSize = 11
            };
            Canvas.SetLeft(label, x + 4);
            Canvas.SetTop(label, laneTop + 8);
            WaveformCanvas.Children.Add(label);

            AddMoveThumb(seg, x + 8, laneTop + 5, Math.Max(12, segWidth - 16), laneHeight - 10);
            AddThumb(seg, "left", x - 5, laneTop + 5, laneHeight - 10, Cursors.SizeWE);
            AddThumb(seg, "right", x + segWidth - 5, laneTop + 5, laneHeight - 10, Cursors.SizeWE);

            var markerX = Math.Clamp(GetRenderedMarker(seg), start, end) * pps;
            var cueLine = new System.Windows.Shapes.Line
            {
                X1 = markerX,
                X2 = markerX,
                Y1 = waveformTop,
                Y2 = laneTop + laneHeight - 4,
                Stroke = ReferenceEquals(seg, selected)
                    ? new SolidColorBrush(Color.FromRgb(205, 98, 18))
                    : new SolidColorBrush(Color.FromRgb(216, 116, 33)),
                StrokeThickness = ReferenceEquals(seg, selected) ? 3 : 2
            };
            WaveformCanvas.Children.Add(cueLine);

            var cueHead = new System.Windows.Shapes.Polygon
            {
                Fill = new SolidColorBrush(Color.FromRgb(216, 116, 33)),
                Stroke = new SolidColorBrush(Color.FromRgb(164, 82, 16)),
                StrokeThickness = 1,
                Points = new PointCollection
                {
                    new(markerX, waveformTop),
                    new(markerX - 6, waveformTop - 9),
                    new(markerX + 6, waveformTop - 9)
                }
            };
            WaveformCanvas.Children.Add(cueHead);

            AddCueThumb(seg, markerX - 7, waveformTop - 10, laneTop + laneHeight - waveformTop + 10);
        }
    }

    private static string FormatTimelineTime(int totalSeconds)
    {
        var minutes = totalSeconds / 60;
        var seconds = totalSeconds % 60;
        return $"{minutes:00}:{seconds:00}";
    }

    private void AddThumb(SegmentRow seg, string handle, double left, double top, double height, Cursor cursor)
    {
        var thumb = new Thumb
        {
            Width = 10,
            Height = Math.Max(24, height),
            Cursor = cursor,
            Background = new SolidColorBrush(Color.FromArgb(40, 255, 255, 255)),
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Opacity = 1,
            Tag = new DragTag(seg, handle)
        };
        thumb.DragStarted += Thumb_OnDragStarted;
        thumb.DragDelta += Thumb_OnDragDelta;
        thumb.DragCompleted += Thumb_OnDragCompleted;
        Canvas.SetLeft(thumb, left);
        Canvas.SetTop(thumb, top);
        WaveformCanvas.Children.Add(thumb);
    }

    private void AddMoveThumb(SegmentRow seg, double left, double top, double width, double height)
    {
        var thumb = new Thumb
        {
            Width = Math.Max(12, width),
            Height = Math.Max(24, height),
            Cursor = Cursors.SizeAll,
            Background = Brushes.Transparent,
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Tag = new DragTag(seg, "move")
        };
        thumb.DragStarted += Thumb_OnDragStarted;
        thumb.DragDelta += Thumb_OnDragDelta;
        thumb.DragCompleted += Thumb_OnDragCompleted;
        Canvas.SetLeft(thumb, left);
        Canvas.SetTop(thumb, top);
        WaveformCanvas.Children.Add(thumb);
    }

    private void AddCueThumb(SegmentRow seg, double left, double top, double height)
    {
        var thumb = new Thumb
        {
            Width = 18,
            Height = Math.Max(30, height),
            Cursor = Cursors.Hand,
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Tag = new DragTag(seg, "marker")
        };
        thumb.DragStarted += Thumb_OnDragStarted;
        thumb.DragDelta += Thumb_OnDragDelta;
        thumb.DragCompleted += Thumb_OnDragCompleted;
        Canvas.SetLeft(thumb, left);
        Canvas.SetTop(thumb, top);
        WaveformCanvas.Children.Add(thumb);
    }

    private void Thumb_OnDragStarted(object sender, DragStartedEventArgs e)
    {
        if (sender is not Thumb { Tag: DragTag tag })
        {
            return;
        }

        SegmentsGrid.SelectedItem = tag.Segment;
        _dragPreview = new DragPreviewState(
            tag.Segment,
            tag.Handle,
            tag.Segment.StartSeconds,
            tag.Segment.EndSeconds,
            tag.Segment.OneShotSeconds);
        RenderWaveform();
    }

    private void Thumb_OnDragDelta(object sender, DragDeltaEventArgs e)
    {
        if (sender is not Thumb { Tag: DragTag tag })
        {
            return;
        }

        var pps = GetPixelsPerSecond();
        if (pps <= 0)
        {
            return;
        }

        var deltaSec = e.HorizontalChange / pps;
        var seg = tag.Segment;
        var duration = Math.Max(_durationSeconds, Math.Max(GetRenderedEnd(seg), 0.1));
        var preview = _dragPreview is not null && ReferenceEquals(_dragPreview.Segment, seg) && string.Equals(_dragPreview.Handle, tag.Handle, StringComparison.Ordinal)
            ? _dragPreview
            : new DragPreviewState(seg, tag.Handle, seg.StartSeconds, seg.EndSeconds, seg.OneShotSeconds);

        var start = preview.StartSeconds;
        var end = preview.EndSeconds;
        var marker = preview.OneShotSeconds;

        switch (tag.Handle)
        {
            case "left":
            {
                start = Math.Clamp(start + deltaSec, 0, Math.Max(0, end - 0.05));
                if (marker < start) marker = start;
                break;
            }
            case "right":
            {
                end = Math.Clamp(end + deltaSec, Math.Min(duration, start + 0.05), duration);
                if (marker > end) marker = end;
                break;
            }
            case "move":
            {
                var segLength = Math.Max(0.05, end - start);
                var markerOffset = marker - start;
                start += deltaSec;
                end += deltaSec;

                if (start < 0)
                {
                    end -= start;
                    start = 0;
                }

                if (end > duration)
                {
                    var overshoot = end - duration;
                    start = Math.Max(0, start - overshoot);
                    end = duration;
                }

                if (end - start < segLength)
                {
                    end = Math.Min(duration, start + segLength);
                }

                marker = Math.Clamp(start + markerOffset, start, end);
                break;
            }
            case "marker":
            {
                marker = Math.Clamp(marker + deltaSec, start, end);
                break;
            }
        }

        _dragPreview = preview with
        {
            StartSeconds = start,
            EndSeconds = end,
            OneShotSeconds = marker
        };
        RenderWaveform();
    }

    private void Thumb_OnDragCompleted(object sender, DragCompletedEventArgs e)
    {
        if (_dragPreview is null)
        {
            return;
        }

        var preview = _dragPreview;
        _dragPreview = null;
        _updatingRows = true;
        try
        {
            preview.Segment.StartSeconds = preview.StartSeconds;
            preview.Segment.EndSeconds = preview.EndSeconds;
            preview.Segment.OneShotSeconds = preview.OneShotSeconds;
        }
        finally
        {
            _updatingRows = false;
        }
        RenderWaveform();
    }

    private double GetPixelsPerSecond()
    {
        var zoom = WaveZoomSlider?.Value ?? 1.0;
        return 80.0 * Math.Clamp(zoom, 0.5, 4.0);
    }

    private void WaveZoomSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        RenderWaveform();
    }

    private double GetRenderedStart(SegmentRow seg)
    {
        return _dragPreview is not null && ReferenceEquals(_dragPreview.Segment, seg)
            ? _dragPreview.StartSeconds
            : seg.StartSeconds;
    }

    private double GetRenderedEnd(SegmentRow seg)
    {
        return _dragPreview is not null && ReferenceEquals(_dragPreview.Segment, seg)
            ? _dragPreview.EndSeconds
            : seg.EndSeconds;
    }

    private double GetRenderedMarker(SegmentRow seg)
    {
        return _dragPreview is not null && ReferenceEquals(_dragPreview.Segment, seg)
            ? _dragPreview.OneShotSeconds
            : seg.OneShotSeconds;
    }

    private void PlayButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_waveOut is not null && _playbackReader is not null)
            {
                _waveOut.Play();
                _playbackTimer.Start();
                return;
            }

            if (string.IsNullOrWhiteSpace(_sourcePath) || !File.Exists(_sourcePath))
            {
                return;
            }

            _playbackReader = new AudioFileReader(_sourcePath);
            _waveOut = new WaveOutEvent();
            _waveOut.Init(_playbackReader);
            _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
            _waveOut.Play();
            _playbackTimer.Start();
            UpdatePlaybackTimeText();
        }
        catch (Exception ex)
        {
            StopPlayback();
            MessageBox.Show(this,
                UserMessageFormatter.FormatOperationError("Play source audio", ex.Message),
                "Prepare Enhancement",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void StopButton_OnClick(object sender, RoutedEventArgs e)
    {
        StopPlayback();
    }

    private void WaveOut_PlaybackStopped(object? sender, StoppedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            _playbackTimer.Stop();
            UpdatePlaybackTimeText(resetToZero: _playbackReader is not null && _playbackReader.CurrentTime >= _playbackReader.TotalTime);
        });
    }

    private void PlaybackTimer_OnTick(object? sender, EventArgs e)
    {
        UpdatePlaybackTimeText();
    }

    private void UpdatePlaybackTimeText(bool resetToZero = false)
    {
        if (PlaybackTimeText is null)
        {
            return;
        }

        var current = resetToZero || _playbackReader is null ? TimeSpan.Zero : _playbackReader.CurrentTime;
        var total = _playbackReader?.TotalTime ?? TimeSpan.FromSeconds(Math.Max(0, _durationSeconds));
        PlaybackTimeText.Text = $"{current:mm\\:ss} / {total:mm\\:ss}";
    }

    private void StopPlayback()
    {
        _playbackTimer.Stop();
        if (_waveOut is not null)
        {
            _waveOut.PlaybackStopped -= WaveOut_PlaybackStopped;
            _waveOut.Stop();
            _waveOut.Dispose();
            _waveOut = null;
        }

        if (_playbackReader is not null)
        {
            _playbackReader.Dispose();
            _playbackReader = null;
        }

        UpdatePlaybackTimeText(resetToZero: true);
    }

    private void ImportSubtitleButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Subtitle Files (*.srt;*.ass)|*.srt;*.ass|All Files (*.*)|*.*",
            Multiselect = false,
            InitialDirectory = Path.GetDirectoryName(_sourcePath) ?? AppDomain.CurrentDomain.BaseDirectory
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var parsed = ParseSubtitleFileToPreparedSegments(dialog.FileName);
        if (parsed.Count == 0)
        {
            MessageBox.Show(this, "No subtitle lines were parsed from this file.", "Prepare Enhancement", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _updatingRows = true;
        try
        {
            _segments.Clear();
            foreach (var s in parsed)
            {
                var row = new SegmentRow(s);
                row.PropertyChanged += Segment_OnPropertyChanged;
                _segments.Add(row);
            }
            Reindex();
            NormalizeRows();
        }
        finally
        {
            _updatingRows = false;
        }

        StatusText.Text = $"Imported subtitle: {Path.GetFileName(dialog.FileName)}";
        RenderWaveform();
    }

    private void AddSegmentButton_OnClick(object sender, RoutedEventArgs e)
    {
        var duration = Math.Max(_durationSeconds, 10.0);
        var lastEnd = _segments.Count == 0 ? 0 : _segments.Max(s => s.EndSeconds);
        var start = Math.Clamp(lastEnd, 0, Math.Max(0, duration - 0.5));
        var end = Math.Clamp(start + 4.0, start + 0.2, duration);

        var row = CreateSegmentRow(start, end);
        _segments.Add(row);
        SegmentsGrid.SelectedItem = row;
    }

    private void AddSegmentAfterButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: SegmentRow row })
        {
            return;
        }

        var duration = Math.Max(_durationSeconds, 10.0);
        var insertIndex = Math.Max(0, _segments.IndexOf(row) + 1);
        var start = Math.Clamp(row.EndSeconds, 0, Math.Max(0, duration - 0.5));
        var end = Math.Clamp(start + 4.0, start + 0.2, duration);

        var newRow = CreateSegmentRow(start, end);
        _segments.Insert(insertIndex, newRow);
        SegmentsGrid.SelectedItem = newRow;
    }

    private void DeleteSegmentRowButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: SegmentRow row })
        {
            return;
        }

        _segments.Remove(row);
    }

    private void DeleteSegmentButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (SegmentsGrid.SelectedItem is not SegmentRow row)
        {
            return;
        }
        _segments.Remove(row);
    }

    private void ClearSegmentsButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_segments.Count == 0)
        {
            return;
        }
        var result = MessageBox.Show(
            this,
            "Clear all segments?",
            "Prepare Enhancement",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }
        _segments.Clear();
    }

    private void SegmentsGrid_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        RenderWaveform();
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        NormalizeRows();
        ResultSegments = _segments
            .OrderBy(s => s.Order)
            .ThenBy(s => s.StartSeconds)
            .Select((s, idx) => new EnhancePreparedSegment
            {
                Id = string.IsNullOrWhiteSpace(s.Id) ? Guid.NewGuid().ToString("N") : s.Id,
                Order = idx + 1,
                StartSeconds = Math.Max(0, s.StartSeconds),
                EndSeconds = Math.Max(s.StartSeconds + 0.05, s.EndSeconds),
                Text = s.Text?.Trim() ?? string.Empty,
                AmbiencePrompt = s.AmbiencePrompt?.Trim() ?? string.Empty,
                OneShotPrompt = s.OneShotPrompt?.Trim() ?? string.Empty,
                OneShotSeconds = Math.Clamp(s.OneShotSeconds, s.StartSeconds, s.EndSeconds),
                Intensity = Math.Clamp(s.Intensity <= 0 ? 0.6 : s.Intensity, 0.1, 1.0),
                Enabled = s.Enabled
            })
            .ToList();
        DialogResult = true;
        Close();
    }

    private void CancelButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void Reindex()
    {
        for (var i = 0; i < _segments.Count; i++)
        {
            _segments[i].Order = i + 1;
        }
    }

    private SegmentRow CreateSegmentRow(double start, double end)
    {
        var row = new SegmentRow(new EnhancePreparedSegment
        {
            Id = Guid.NewGuid().ToString("N"),
            StartSeconds = start,
            EndSeconds = end,
            OneShotSeconds = Math.Clamp(start + 1.0, start, end),
            Intensity = 0.6,
            Enabled = true
        });
        row.PropertyChanged += Segment_OnPropertyChanged;
        return row;
    }

    private void NormalizeRows()
    {
        var duration = Math.Max(_durationSeconds, 0.2);
        foreach (var seg in _segments)
        {
            seg.StartSeconds = Math.Clamp(seg.StartSeconds, 0, duration);
            seg.EndSeconds = Math.Clamp(seg.EndSeconds, seg.StartSeconds + 0.05, duration);
            seg.OneShotSeconds = Math.Clamp(seg.OneShotSeconds, seg.StartSeconds, seg.EndSeconds);
            seg.Intensity = Math.Clamp(seg.Intensity <= 0 ? 0.6 : seg.Intensity, 0.1, 1.0);
        }
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
        void Flush()
        {
            if (block.Count == 0) return;
            var timeLine = block.FirstOrDefault(l => l.Contains("-->", StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(timeLine))
            {
                var parts = timeLine.Split(new[] { "-->" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 &&
                    TryParseSrtTimestamp(parts[0].Trim(), out var start) &&
                    TryParseSrtTimestamp(parts[1].Trim(), out var end))
                {
                    var text = string.Join(" ", block.Where(l => !l.Contains("-->", StringComparison.Ordinal) && !Regex.IsMatch(l.Trim(), @"^\d+$"))).Trim();
                    cues.Add(new SubtitleCue(start, end, text));
                }
            }
            block.Clear();
        }

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                Flush();
            }
            else
            {
                block.Add(line.Trim());
            }
        }
        Flush();
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
            if (fields.Length < 10) continue;
            if (!TryParseAssTimestamp(fields[1].Trim(), out var start) || !TryParseAssTimestamp(fields[2].Trim(), out var end))
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
        if (parts.Length < 4) return false;
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
        if (parts.Length < 4) return false;
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

    private readonly record struct SubtitleCue(double StartSeconds, double EndSeconds, string Text);
    private readonly record struct DragTag(SegmentRow Segment, string Handle);
    private sealed record DragPreviewState(SegmentRow Segment, string Handle, double StartSeconds, double EndSeconds, double OneShotSeconds);

    private sealed class SegmentRow : INotifyPropertyChanged
    {
        private int _order;
        private double _startSeconds;
        private double _endSeconds;
        private string _text;
        private string _ambiencePrompt;
        private string _oneShotPrompt;
        private double _oneShotSeconds;
        private double _intensity;
        private bool _enabled;

        public SegmentRow(EnhancePreparedSegment src)
        {
            Id = string.IsNullOrWhiteSpace(src.Id) ? Guid.NewGuid().ToString("N") : src.Id;
            _order = src.Order;
            _startSeconds = src.StartSeconds;
            _endSeconds = src.EndSeconds;
            _text = src.Text ?? string.Empty;
            _ambiencePrompt = src.AmbiencePrompt ?? string.Empty;
            _oneShotPrompt = src.OneShotPrompt ?? string.Empty;
            _oneShotSeconds = src.OneShotSeconds;
            _intensity = src.Intensity <= 0 ? 0.6 : src.Intensity;
            _enabled = src.Enabled;
        }

        public string Id { get; }

        public int Order { get => _order; set => Set(ref _order, value); }
        public double StartSeconds { get => _startSeconds; set => Set(ref _startSeconds, value); }
        public double EndSeconds { get => _endSeconds; set => Set(ref _endSeconds, value); }
        public string Text { get => _text; set => Set(ref _text, value ?? string.Empty); }
        public string AmbiencePrompt { get => _ambiencePrompt; set => Set(ref _ambiencePrompt, value ?? string.Empty); }
        public string OneShotPrompt { get => _oneShotPrompt; set => Set(ref _oneShotPrompt, value ?? string.Empty); }
        public double OneShotSeconds { get => _oneShotSeconds; set => Set(ref _oneShotSeconds, value); }
        public double Intensity { get => _intensity; set => Set(ref _intensity, value); }
        public bool Enabled { get => _enabled; set => Set(ref _enabled, value); }

        private void Set<T>(ref T field, T value, [CallerMemberName] string? property = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return;
            }
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
