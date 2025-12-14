using System;
using System.Diagnostics;
using System.Net.Http;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Serilog;
using ILogger = Serilog.ILogger;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Interop;
using System.Windows.Shell;
using System.Collections.Generic;
using System.ComponentModel;
using WinForms = System.Windows.Forms;
using Microsoft.Win32;
using Application = System.Windows.Application;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using SDColor = System.Drawing.Color;
using Button = System.Windows.Controls.Button;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using System.Text.Json;
using Bobrus.App.Services;
using System.Runtime.InteropServices;
using Serilog.Events;
using System.Text.RegularExpressions;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Globalization;
using System.Timers;
using Timer = System.Timers.Timer;

namespace Bobrus.App;

public partial class MainWindow : Window
{
    private readonly ILogger _logger = Log.ForContext<MainWindow>();
    private readonly HttpClient _httpClient = new();
    private readonly UpdateService _updateService;
    private readonly CleaningService _cleaningService = new();
    private readonly SecurityService _securityService = new();
    private readonly TlsConfigurator _tlsConfigurator = new();
    private readonly PrintSpoolService _printSpoolService = new();
    private readonly TouchDeviceService _touchManager = new();
    private readonly Thickness _defaultResizeBorder = new(8);
    private HwndSource? _hwndSource;
    private readonly bool _startHidden;
    private WinForms.NotifyIcon? _trayIcon;
    private bool _hideToTrayEnabled;
    private bool _autostartEnabled;
    private bool _suppressSettingsToggle;
    private bool _isExiting;
    private bool _updateCheckInProgress;
    private Timer? _autoUpdateTimer;
    private bool _showAllSections;
    private bool _consoleVisible;
    private bool _overlayEnabled;
    private string _overlayCrm = string.Empty;
    private string _overlayCashDesk = string.Empty;
    private string _overlayDisplay = string.Empty;
    private readonly bool _launchSetupWizard;
    private double _restoreLeft;
    private double _restoreTop;
    private double _restoreWidth;
    private double _restoreHeight;
    private Section _currentSection = Section.System;
    private readonly Dictionary<Section, List<Button>> _sectionButtons = new();
    private readonly Dictionary<Button, string> _buttonSearchIndex = new();
    private SetupWizardWindow? _setupWizardWindow;
    private ThemeVariant _currentTheme = ThemeVariant.Light;
    private readonly bool _isFirstRunSettings;
    private readonly Dictionary<string, Color> _lightTheme = new()
    {
        ["BackgroundBrush"] = ColorFromHex("#F5F7FA"),
        ["PanelBrush"] = ColorFromHex("#FFFFFF"),
        ["SurfaceBrush"] = ColorFromHex("#F1F4F9"),
        ["BorderBrushMuted"] = ColorFromHex("#DDE2EB"),
        ["TextPrimaryBrush"] = ColorFromHex("#1E2330"),
        ["TextSecondaryBrush"] = ColorFromHex("#4A5568"),
        ["AccentBrush"] = ColorFromHex("#2F855A"),
        ["AccentBrushHover"] = ColorFromHex("#276749"),
        ["AccentBlueBrush"] = ColorFromHex("#2B6CB0"),
        ["AccentBlueHoverBrush"] = ColorFromHex("#245FA1"),
        ["DangerBrush"] = ColorFromHex("#C53030"),
        ["DangerHoverBrush"] = ColorFromHex("#A11919"),
        ["TopBarBrush"] = ColorFromHex("#EEF1F7"),
        ["ScrollbarTrackBrush"] = ColorFromHex("#E7EBF2"),
        ["ScrollbarThumbBrush"] = ColorFromHex("#C6CCD8"),
        ["ScrollbarThumbHoverBrush"] = ColorFromHex("#B4BAC7"),
        ["ScrollbarThumbPressedBrush"] = ColorFromHex("#A2A8B6"),
        ["IconHoverBrush"] = ColorFromHex("#E6ECF4"),
        ["IconPressedBrush"] = ColorFromHex("#D7DEE9"),
        ["CloseHoverBrush"] = ColorFromHex("#FFE8E8"),
        ["ClosePressedBrush"] = ColorFromHex("#FFD6D6"),
        ["ToggleTrackBrush"] = ColorFromHex("#E2E6EE"),
        ["ComboHighlightBrush"] = ColorFromHex("#E8EDF7"),
        ["ComboToggleBackgroundBrush"] = ColorFromHex("#EDEFF5"),
        ["InputBackgroundBrush"] = ColorFromHex("#FFFFFF"),
        ["InputForegroundBrush"] = ColorFromHex("#1E2330"),
        ["InputButtonBackgroundBrush"] = ColorFromHex("#E6EAF2"),
        ["ConsoleBackgroundBrush"] = ColorFromHex("#F2F4F8"),
        ["ConsoleForegroundBrush"] = ColorFromHex("#1E2330"),
        ["OnAccentBrush"] = ColorFromHex("#FFFFFF"),
        ["NavButtonActiveBrush"] = ColorFromHex("#D8E7FF"),
        ["NavButtonActiveBorderBrush"] = ColorFromHex("#AAC7FF"),
        ["OverlayBackgroundBrush"] = ColorFromHex("#99000000"), 
        ["CardBrush"] = ColorFromHex("#FFFFFF")
    };

    private readonly Dictionary<string, Color> _darkTheme = new()
    {
        ["BackgroundBrush"] = ColorFromHex("#0D0D11"),
        ["PanelBrush"] = ColorFromHex("#13131A"),
        ["SurfaceBrush"] = ColorFromHex("#1A1A23"),
        ["BorderBrushMuted"] = ColorFromHex("#25252F"),
        ["TextPrimaryBrush"] = ColorFromHex("#F0F2F8"),
        ["TextSecondaryBrush"] = ColorFromHex("#B5BAC7"),
        ["AccentBrush"] = ColorFromHex("#3EA16C"),
        ["AccentBrushHover"] = ColorFromHex("#318659"),
        ["AccentBlueBrush"] = ColorFromHex("#2D72D2"),
        ["AccentBlueHoverBrush"] = ColorFromHex("#245EB0"),
        ["DangerBrush"] = ColorFromHex("#C0392B"),
        ["DangerHoverBrush"] = ColorFromHex("#A43126"),
        ["TopBarBrush"] = ColorFromHex("#16171F"),
        ["ScrollbarTrackBrush"] = ColorFromHex("#0F0F14"),
        ["ScrollbarThumbBrush"] = ColorFromHex("#2B2B35"),
        ["ScrollbarThumbHoverBrush"] = ColorFromHex("#3A3A46"),
        ["ScrollbarThumbPressedBrush"] = ColorFromHex("#4A4A58"),
        ["IconHoverBrush"] = ColorFromHex("#1E1E26"),
        ["IconPressedBrush"] = ColorFromHex("#25252F"),
        ["CloseHoverBrush"] = ColorFromHex("#3D1D1D"),
        ["ClosePressedBrush"] = ColorFromHex("#5A2626"),
        ["ToggleTrackBrush"] = ColorFromHex("#3A3A45"),
        ["ComboHighlightBrush"] = ColorFromHex("#262633"),
        ["ComboToggleBackgroundBrush"] = ColorFromHex("#1F1F2A"),
        ["InputBackgroundBrush"] = ColorFromHex("#161620"),
        ["InputForegroundBrush"] = ColorFromHex("#EAEAEA"),
        ["InputButtonBackgroundBrush"] = ColorFromHex("#1F1F2A"),
        ["ConsoleBackgroundBrush"] = ColorFromHex("#0C0C10"),
        ["ConsoleForegroundBrush"] = ColorFromHex("#EAEAEA"),
        ["OnAccentBrush"] = ColorFromHex("#F0F2F8"),
        ["NavButtonActiveBrush"] = ColorFromHex("#233552"),
        ["NavButtonActiveBorderBrush"] = ColorFromHex("#34517A"),
        ["OverlayBackgroundBrush"] = ColorFromHex("#99000000"), 
        ["CardBrush"] = ColorFromHex("#13131A") 
    };
    private const string StartHiddenArg = "--start-hidden";
    private const string SetupWizardArg = "--setup-wizard";
    private const double ConsoleColumnMinWidth = 200;
    private const double ConsoleColumnMaxWidth = 320;
    private string SettingsFilePath => Path.Combine(AppPaths.AppDataRoot, "settings.json");
    private const string IikoFrontExePath = @"C:\Program Files\iiko\iikoRMS\Front.Net\iikoFront.Net.exe";
    private const string IikoCardUrl = "https://iiko.biz/ru-RU/About/DownloadPosInstaller?useRc=False";
    private readonly string _cashServerBase = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "iiko",
        "CashServer");
    private Action? _pendingConfirmAction;
    private bool? _isTouchEnabled;
    private List<Button> _actionButtons = new();
    private Brush? _cachedTextSecondaryBrush;
    private Brush? _cachedDangerBrush;
    private Brush? _cachedAccentBrush;
    private Brush? _cachedTextPrimaryBrush;
    private static readonly SolidColorBrush AccentOrangeBrush = new(Color.FromRgb(0xE7, 0xA0, 0x4F));
    private Style? _cachedDangerButtonStyle;
    private Style? _cachedPrimaryButtonStyle;
    private Style? _cachedBaseButtonStyle;
    private CancellationTokenSource? _searchDebounceToken;
    private const int MaxLogEntries = 200;

    public MainWindow()
    {
        var settings = LoadAppSettings();
        _hideToTrayEnabled = settings?.HideToTray ?? true;
        _autostartEnabled = settings?.Autostart ?? true;
        _showAllSections = settings?.ShowAllSections ?? false;
        _consoleVisible = settings?.ShowConsole ?? true;
        _overlayEnabled = settings?.OverlayEnabled ?? false;
        _overlayCrm = settings?.OverlayCrm?.Trim() ?? string.Empty;
        _overlayCashDesk = settings?.OverlayCashDesk?.Trim() ?? string.Empty;
        _overlayDisplay = settings?.OverlayDisplay ?? string.Empty;
        _currentTheme = ParseTheme(settings?.Theme);
        _isFirstRunSettings = settings is null;

        ApplyTheme(_currentTheme);

        InitializeComponent();
        CacheResources();
        CaptureRestoreBounds();
        LocationChanged += OnWindowLocationChanged;
        SizeChanged += OnWindowSizeChanged;
        _updateService = new UpdateService(_httpClient);
        VersionText.Text = $"v{_updateService.CurrentVersion.ToString(3)}";
        _logger.Information("Bobrus запущен. Текущая версия {Version}", _updateService.CurrentVersion);
        _startHidden = Environment.GetCommandLineArgs().Any(a => string.Equals(a, StartHiddenArg, StringComparison.OrdinalIgnoreCase));
        _launchSetupWizard = Environment.GetCommandLineArgs().Any(a => string.Equals(a, SetupWizardArg, StringComparison.OrdinalIgnoreCase));
        ApplyInitialSettingsToUi();
        Loaded += OnLoaded;
    }

    private void CacheResources()
    {
        _cachedTextSecondaryBrush = FindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray;
        _cachedDangerBrush = FindResource("DangerBrush") as Brush ?? Brushes.Red;
        _cachedAccentBrush = FindResource("AccentBrush") as Brush ?? Brushes.LimeGreen;
        _cachedTextPrimaryBrush = FindResource("TextPrimaryBrush") as Brush ?? Brushes.White;
        _cachedDangerButtonStyle = FindResource("DangerButton") as Style;
        _cachedPrimaryButtonStyle = FindResource("PrimaryButton") as Style;
        _cachedBaseButtonStyle = FindResource("BaseButton") as Style;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        AdjustToWorkArea();
        AdjustResizeBorder();

        _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
        _hwndSource?.AddHook(WndProc);
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);
        if (WindowState == WindowState.Minimized)
        {
            HideAllDropdowns();
        }
        AdjustResizeBorder();
        AdjustToWorkArea();
    }

    protected override void OnDeactivated(EventArgs e)
    {
        base.OnDeactivated(e);
        HideAllDropdowns();
    }

    private void OnRebootClicked(object sender, RoutedEventArgs e)
    {
        ShowConfirm("Перезагрузка",
            "Вы уверены, что хотите перезагрузить компьютер?",
            StartReboot);
    }

    private async void OnCheckUpdatesClicked(object sender, RoutedEventArgs e)
    {
        OnSettingsOverlayCloseClicked(this, new RoutedEventArgs());
        await RunUpdateCheckAsync(isAuto: false);
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        _httpClient.Dispose();
        UiLogBuffer.OnLog -= OnUiLog;
        TryKillNetworkProcess();
        TryDisposeNetworkProcess();
        if (_hwndSource is not null)
        {
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }
        if (_trayIcon is not null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }
        _setupWizardWindow?.Close();
        _autoUpdateTimer?.Stop();
        _autoUpdateTimer?.Dispose();
        if (_overlayEnabled && HasOverlayValues())
        {
            SetOverlayAutostart(true);
            StartOverlayRunner();
        }
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        var args = Environment.GetCommandLineArgs();
        var isResume = args.Contains("--resume");
        if ((_launchSetupWizard || isResume) && IsAdministrator())
        {
            ShowSetupWizardWindow(autoResume: isResume);
            return; 
        }

        UiLogBuffer.OnLog += OnUiLog;
        _actionButtons = ActionsPanel.Children.OfType<Button>()
            .Concat(IikoActionsPanel.Children.OfType<Button>())
            .Concat(ProgramsActionsPanel.Children.OfType<Button>())
            .Concat(NetworkActionsPanel.Children.OfType<Button>())
            .Concat(LogsActionsPanel.Children.OfType<Button>())
            .Concat(FoldersActionsPanel.Children.OfType<Button>())
            .ToList();
        BuildSectionButtons();
        RebuildSearchIndex();
        await RefreshTouchStateAsync();
        StartAutoUpdateTimer();
        ApplyStartHidden();
    }

    private void StartAutoUpdateTimer()
    {
        _autoUpdateTimer?.Stop();
        _autoUpdateTimer?.Dispose();
        _autoUpdateTimer = new Timer(TimeSpan.FromHours(2).TotalMilliseconds)
        {
            AutoReset = true,
            Enabled = true
        };
        _autoUpdateTimer.Elapsed += async (_, _) =>
        {
            await Dispatcher.InvokeAsync(async () => await RunUpdateCheckAsync(isAuto: true));
        };
    }

    private void StartReboot()
    {
        try
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "shutdown",
                Arguments = "/r /t 5 /c \"Перезагрузка инициирована Bobrus\"",
                UseShellExecute = true,
                CreateNoWindow = true
            };

            Process.Start(processStartInfo);
            _logger.Information("Перезагрузка запущена пользователем.");
            ShowNotification("Перезагрузка запущена", NotificationType.Success);
            Close();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка при попытке перезагрузки системы.");
            ShowNotification($"Не удалось запустить перезагрузку: {ex.Message}", NotificationType.Error);
        }
    }

    private async Task RefreshTouchStateAsync()
    {
        TouchToggleButton.IsEnabled = false;
        RestartTouchButton.IsEnabled = false;
        TouchToggleButton.Content = "Сенсор: проверка...";
        TouchToggleButton.Style = _cachedBaseButtonStyle;

        var devices = await _touchManager.GetTouchDevicesAsync();
        LogTouchDevices(devices, "initial");
        if (devices.Count == 0)
        {
            TouchToggleButton.Content = "Сенсор не найден";
            TouchToggleButton.IsEnabled = false;
            RestartTouchButton.IsEnabled = false;
            _isTouchEnabled = null;
            return;
        }

        _isTouchEnabled = devices.Any(d => d.IsEnabled);
        UpdateTouchButtonVisual();
        TouchToggleButton.IsEnabled = true;
        RestartTouchButton.IsEnabled = true;
    }

    private async void OnTouchToggleClicked(object sender, RoutedEventArgs e)
    {
        TouchToggleButton.IsEnabled = false;
        TouchToggleButton.Content = "Применение...";
        try
        {
            var targetState = !(_isTouchEnabled ?? true);
            _logger.Information("Touch toggle requested: target={Target}", targetState ? "on" : "off");
            var ok = await _touchManager.SetTouchEnabledAsync(targetState);
            if (!ok)
            {
                ShowNotification("Сенсор не найден или не удалось применить", NotificationType.Error);
            }
            else
            {
                _isTouchEnabled = targetState;
                UpdateTouchButtonVisual();
                ShowNotification(targetState ? "Сенсор включён" : "Сенсор отключён", NotificationType.Success);
                var devices = await _touchManager.GetTouchDevicesAsync();
                LogTouchDevices(devices, "after-toggle");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось переключить сенсор");
            ShowNotification($"Ошибка переключения сенсора: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            TouchToggleButton.IsEnabled = true;
        }
    }

    private void OnClearLogClicked(object sender, RoutedEventArgs e)
    {
        LogRichTextBox.Document.Blocks.Clear();
        LogRichTextBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph());
    }

    private static readonly Regex LogHighlightPattern = new(@"\d[\d\s]*байт|начало", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Thickness LogParagraphMargin = new(0, 0, 0, 4);

    private void OnUiLog(LogEvent logEvent)
    {
        if (!_consoleVisible) return;

        Dispatcher.InvokeAsync(() =>
        {
            var blocks = LogRichTextBox.Document.Blocks;
            while (blocks.Count >= MaxLogEntries)
            {
                blocks.Remove(blocks.FirstBlock);
            }

            var paragraph = new System.Windows.Documents.Paragraph { Margin = LogParagraphMargin };
            var timeRun = new System.Windows.Documents.Run($"[{logEvent.Timestamp:HH:mm:ss}] ")
            {
                Foreground = _cachedTextSecondaryBrush
            };

            paragraph.Inlines.Add(timeRun);
            foreach (var inline in BuildMessageInlines(logEvent))
            {
                paragraph.Inlines.Add(inline);
            }
            blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        }, DispatcherPriority.Background);
    }

    private IEnumerable<System.Windows.Documents.Inline> BuildMessageInlines(LogEvent logEvent)
    {
        var message = logEvent.RenderMessage();
        if (logEvent.Level == LogEventLevel.Error || logEvent.Level == LogEventLevel.Fatal)
        {
            yield return new System.Windows.Documents.Run(message)
            {
                Foreground = _cachedDangerBrush,
                FontWeight = FontWeights.SemiBold
            };
            yield break;
        }

        var lastIndex = 0;
        foreach (Match match in LogHighlightPattern.Matches(message))
        {
            if (match.Index > lastIndex)
            {
                var text = message.Substring(lastIndex, match.Index - lastIndex);
                yield return new System.Windows.Documents.Run(text)
                {
                    Foreground = AccentOrangeBrush
                };
            }

            var highlightText = match.Value;
            yield return new System.Windows.Documents.Run(highlightText)
            {
                Foreground = _cachedAccentBrush,
                FontWeight = FontWeights.SemiBold
            };

            lastIndex = match.Index + match.Length;
        }

        if (lastIndex < message.Length)
        {
            var tail = message.Substring(lastIndex);
            yield return new System.Windows.Documents.Run(tail)
            {
                Foreground = AccentOrangeBrush
            };
        }
    }

    private async void OnRestartTouchClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Перезапуск...";
        var toggleWasEnabled = TouchToggleButton.IsEnabled;
        TouchToggleButton.IsEnabled = false;
        try
        {
            _logger.Information("Touch restart requested");
            var ok = await _touchManager.RestartTouchAsync();
            if (!ok)
            {
                ShowNotification("Сенсор не найден для перезапуска", NotificationType.Error);
            }
            else
            {
                ShowNotification("Сенсор перезапущен", NotificationType.Success);
                _isTouchEnabled = true;
                UpdateTouchButtonVisual();
                var devices = await _touchManager.GetTouchDevicesAsync();
                LogTouchDevices(devices, "after-restart");
                await RefreshTouchStateAsync();
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось перезапустить сенсор");
            ShowNotification($"Ошибка перезапуска сенсора: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Перезапуск сенсора";
            button.IsEnabled = true;
            TouchToggleButton.IsEnabled = toggleWasEnabled;
        }
    }

    private void LogTouchDevices(IReadOnlyList<Services.TouchDevice> devices, string stage)
    {
        var info = devices.Select(d => $"{d.FriendlyName} [{d.InstanceId}] {(d.IsEnabled ? "ON" : "OFF")}");
        _logger.Information("Touch devices ({Stage}): {List}", stage, string.Join("; ", info));
    }

    private async void OnCleanupClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Очистка...";
        try
        {
            _logger.Information("Очистка: старт");
            var results = await _cleaningService.RunCleanupAsync(progress =>
            {
                if (progress.IsStart)
                {
                    _logger.Information("Очистка {Name}: начало", progress.Name);
                }
                else
                {
                    _logger.Information("Очистка {Name}: освобождено {Bytes} байт", progress.Name, progress.BytesFreed);
                }
            });
            var freed = results.Sum(r => r.BytesFreed);
            _logger.Information("Очистка: итог освобождено {Bytes} байт", freed);
            ShowNotification($"Очистка завершена. Освобождено {FormatBytes(freed)}", NotificationType.Success);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка очистки");
            ShowNotification($"Ошибка очистки: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Очистить мусор";
            button.IsEnabled = true;
        }
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes < 1024) return $"{bytes} Б";
        double kb = bytes / 1024.0;
        if (kb < 1024) return $"{kb:F1} КБ";
        double mb = kb / 1024.0;
        if (mb < 1024) return $"{mb:F1} МБ";
        double gb = mb / 1024.0;
        return $"{gb:F1} ГБ";
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        _searchDebounceToken?.Cancel();
        _searchDebounceToken = new CancellationTokenSource();
        var token = _searchDebounceToken.Token;
        var searchText = SearchBox.Text;

        Task.Delay(150, token).ContinueWith(_ =>
        {
            if (!token.IsCancellationRequested)
            {
                Dispatcher.Invoke(() => ApplySearch(searchText));
            }
        }, token, TaskContinuationOptions.OnlyOnRanToCompletion, TaskScheduler.Default);
    }

    private void OnSearchKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            SearchBox.Text = string.Empty;
            ApplySearch(string.Empty);
            e.Handled = true;
        }
    }

    private void ApplySearch(string query)
    {
        var q = (query ?? string.Empty).Trim().ToLowerInvariant();
        var targetButtons = _actionButtons;

        if (q.Length == 0)
        {
            foreach (var btn in targetButtons)
            {
                SetVisibilityIfChanged(btn, Visibility.Visible);
            }

            if (_showAllSections)
            {
                SetVisibilityIfChanged(SystemCard, Visibility.Visible);
                SetVisibilityIfChanged(IikoCard, Visibility.Visible);
                SetVisibilityIfChanged(ProgramsCard, Visibility.Visible);
                SetVisibilityIfChanged(NetworkCard, Visibility.Visible);
                SetVisibilityIfChanged(LogsCard, Visibility.Visible);
                SetVisibilityIfChanged(FoldersCard, Visibility.Visible);
                SetVisibilityIfChanged(PluginsCard, Visibility.Visible);
            }
            else
            {
                SetSectionVisibility(_currentSection);
            }

            return;
        }

        foreach (var btn in targetButtons)
        {
            var text = GetSearchIndex(btn);
            var newVisibility = text.Contains(q) ? Visibility.Visible : Visibility.Collapsed;
            SetVisibilityIfChanged(btn, newVisibility);
        }

        var systemVisible = HasVisibleButtonsFast(ActionsPanel);
        var iikoVisible = HasVisibleButtonsFast(IikoActionsPanel);
        var programsVisible = HasVisibleButtonsFast(ProgramsActionsPanel);
        var networkVisible = HasVisibleButtonsFast(NetworkActionsPanel);
        var logsVisible = HasVisibleButtonsFast(LogsActionsPanel);
        var foldersVisible = HasVisibleButtonsFast(FoldersActionsPanel);

        SetVisibilityIfChanged(SystemCard, systemVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(IikoCard, iikoVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(ProgramsCard, programsVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(NetworkCard, networkVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(LogsCard, logsVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(FoldersCard, foldersVisible ? Visibility.Visible : Visibility.Collapsed);
        SetVisibilityIfChanged(PluginsCard, Visibility.Visible);
    }

    private static void SetVisibilityIfChanged(UIElement element, Visibility newVisibility)
    {
        if (element.Visibility != newVisibility)
            element.Visibility = newVisibility;
    }

    private static bool HasVisibleButtonsFast(System.Windows.Controls.Panel panel)
    {
        foreach (var child in panel.Children)
        {
            if (child is Button { Visibility: Visibility.Visible })
                return true;
        }
        return false;
    }

    private void RebuildSearchIndex()
    {
        _buttonSearchIndex.Clear();
        foreach (var button in _actionButtons)
        {
            _buttonSearchIndex[button] = BuildSearchIndex(button);
        }
    }

    private string GetSearchIndex(Button button)
    {
        if (_buttonSearchIndex.TryGetValue(button, out var cached))
        {
            return cached;
        }

        var built = BuildSearchIndex(button);
        _buttonSearchIndex[button] = built;
        return built;
    }

    private string BuildSearchIndex(Button button)
    {
        var baseText = $"{button.Content} {button.Tag}".ToLowerInvariant();
        var compact = RemoveSeparators(baseText);
        var extras = GetExtraKeywords(button);
        return string.Join(' ', new[] { baseText, compact, extras }).Trim();
    }

    private static string RemoveSeparators(string text)
    {
        var chars = text.Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_' && c != '.').ToArray();
        return new string(chars);
    }

    private string GetExtraKeywords(Button button)
    {
        var content = (button.Content?.ToString() ?? string.Empty).ToLowerInvariant();
        var tag = (button.Tag?.ToString() ?? string.Empty).ToLowerInvariant();
        var builder = new List<string>();

        if (content.Contains("iiko") || tag.Contains("iiko") || content.Contains("front"))
        {
            builder.Add("айко айка айкофронт айкафронт фронт iikofront iikofront frontnet иикофронт айко фронт айка фронт");
        }

        if (content.Contains("плаги") || tag.Contains("plugin"))
        {
            builder.Add("plugin plugins плагин плагины модуль моды расширения экстеншн");
        }

        if (content.Contains("лог") || tag.Contains("log"))
        {
            builder.Add("журнал лог файллога logs логфайл log file");
        }

        if (content.Contains("очист") || tag.Contains("cleanup"))
        {
            builder.Add("clean clear почистить удалить мусор temp кеш cache tempfiles");
        }

        if (content.Contains("сеть") || tag.Contains("network"))
        {
            builder.Add("интернет ip сеть сетевой сетка вайфай wifi ethernet lan wan");
        }

        if (content.Contains("обнов") || tag.Contains("update"))
        {
            builder.Add("обновление update апдейт апд");
        }

        return string.Join(' ', builder);
    }

    private async void OnRestartPrintSpoolerClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Перезапуск...";
        try
        {
            _logger.Information("Диспетчер печати: старт перезапуска");
            await _printSpoolService.RestartAndCleanAsync(msg => _logger.Information(msg));
            ShowNotification("Диспетчер печати перезапущен и очередь очищена", NotificationType.Success);
            _logger.Information("Диспетчер печати: завершено");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка перезапуска диспетчера печати");
            ShowNotification($"Ошибка перезапуска печати: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Перезапуск и очистка дисп. печати";
            button.IsEnabled = true;
        }
    }

    private async void OnCloseIikoFrontClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Закрываем...";
        try
        {
            _logger.Information("Закрытие iikoFront: старт");
            var killed = await KillIikoFrontAsync();
            if (killed > 0)
            {
                ShowNotification($"iikoFront закрыт ({killed} процессов)", NotificationType.Success);
                _logger.Information("Закрытие iikoFront: завершено, убито {Count}", killed);
            }
            else
            {
                ShowNotification("iikoFront не запущен", NotificationType.Info);
                _logger.Information("Закрытие iikoFront: процессов не найдено");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка при закрытии iikoFront");
            ShowNotification($"Ошибка закрытия iikoFront: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Закрыть iikoFront";
            button.IsEnabled = true;
        }
    }

    private async void OnRestartIikoFrontClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Перезапуск...";
        try
        {
            _logger.Information("Перезапуск iikoFront: старт");
            await KillIikoFrontAsync();
            await Task.Delay(800);

            if (!File.Exists(IikoFrontExePath))
            {
                ShowNotification("Не найден iikoFront.Net.exe по стандартному пути", NotificationType.Error);
                _logger.Error("Перезапуск iikoFront: файл не найден {Path}", IikoFrontExePath);
                return;
            }

            var psi = new ProcessStartInfo
            {
                FileName = IikoFrontExePath,
                UseShellExecute = true
            };

            var proc = Process.Start(psi);
            if (proc == null)
            {
                ShowNotification("Не удалось запустить iikoFront", NotificationType.Error);
                _logger.Error("Перезапуск iikoFront: Process.Start вернул null");
                return;
            }

            ShowNotification("iikoFront перезапускается", NotificationType.Success);
            _logger.Information("Перезапуск iikoFront: запущен PID {Pid}", proc.Id);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка перезапуска iikoFront");
            ShowNotification($"Ошибка перезапуска iikoFront: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Перезапустить iikoFront";
            button.IsEnabled = true;
        }
    }

    private async void OnUpdateIikoCardClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Загрузка...";
        var tempPath = Path.Combine(Path.GetTempPath(), "iikoCardInstaller.exe");

        try
        {
            _logger.Information("Обновление iikoCard: загрузка из {Url}", IikoCardUrl);
            var progress = new Progress<int>(p =>
            {
                button.Content = $"Загрузка {p}%";
                _logger.Information("Обновление iikoCard: загрузка {Percent}% завершена", p);
            });

            await DownloadFileWithProgressAsync(IikoCardUrl, tempPath, progress, CancellationToken.None);

            button.Content = "Установка...";
            _logger.Information("Обновление iikoCard: загрузка завершена, запускаем установку {Path}", tempPath);

            var installerSize = new FileInfo(tempPath).Length;
            _logger.Information("Обновление iikoCard: размер файла {Size} байт", installerSize);

            var psi = new ProcessStartInfo
            {
                FileName = tempPath,
                Arguments = "/S",
                UseShellExecute = true,
                Verb = "runas"
            };

            var proc = Process.Start(psi);
            if (proc == null)
            {
                ShowNotification("Не удалось запустить установщик iikoCard", NotificationType.Error);
                _logger.Error("Обновление iikoCard: Process.Start вернул null");
                return;
            }

            ShowNotification("Установка iikoCard запущена", NotificationType.Success);
            _logger.Information("Обновление iikoCard: процесс установки PID {Pid}", proc.Id);

            await Task.Run(() => proc.WaitForExit());
            _logger.Information("Обновление iikoCard: установка завершена с кодом {Code}", proc.ExitCode);
            if (proc.ExitCode == 0)
            {
                ShowNotification("iikoCard установлена", NotificationType.Success);
            }
            else
            {
                ShowNotification($"Установка iikoCard завершилась с кодом {proc.ExitCode}", NotificationType.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка обновления iikoCard");
            ShowNotification($"Ошибка обновления iikoCard: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Обновить iikoCard";
            button.IsEnabled = true;
        }
    }

    private async void OnOpenConfigLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "config.xml"), description: "config.xml");
    }

    private async void OnOpenCashServerLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "cash-server.log"), description: "cash-server.log");
    }

    private async void OnOpenOnlineMarkingLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenPatternLogAsync("Resto.Front.Api.OnlineMarkingVerificationPlugin", "OnlineMarking");
    }

    private async void OnOpenDualConnectorLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenPatternLogAsync("Resto.Front.Api.PaymentSystem.DualConnector", "dualConnector");
    }

    private async void OnOpenAlcoholLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenPatternLogAsync("Resto.Front.Api.AlcoholMarkingPlugin", "AlcoholMarking");
    }

    private async void OnOpenSberbankLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenPatternLogAsync("Resto.Front.Api.SberbankPlugin", "Sberbank");
    }

    private async void OnOpenVirtualPrinterLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "virtual-printer.log"), description: "virtual-printer.log");
    }

    private async void OnOpenErrorLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "error.log"), description: "error.log");
    }

    private async void OnOpenTransportLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "plugin-Resto.Front.Api.Transport.V9Preview5.log"), description: "Transport.log");
    }

    private async void OnOpenDeliveryLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "plugin-Resto.Front.Api.Delivery.log"), description: "Delivery.log");
    }

    private async void OnOpenMessagesLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(Path.Combine(_cashServerBase, "Logs", "messages.log"), description: "messages.log");
    }

    private async void OnOpenFolderLogsClicked(object sender, RoutedEventArgs e)
    {
        await OpenFolderAsync(Path.Combine(_cashServerBase, "Logs"), "Logs");
    }

    private async void OnOpenFolderEntitiesClicked(object sender, RoutedEventArgs e)
    {
        await OpenFolderAsync(Path.Combine(_cashServerBase, "EntitiesStorage"), "EntitiesStorage");
    }

    private async void OnOpenFolderCashServerClicked(object sender, RoutedEventArgs e)
    {
        await OpenFolderAsync(_cashServerBase, "CashServer");
    }

    private async void OnOpenFolderPluginsClicked(object sender, RoutedEventArgs e)
    {
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "iiko", "iikoRMS", "Front.Net", "Plugins");
        await OpenFolderAsync(path, "Plugins");
    }

    private async void OnOpenFolderUtmClicked(object sender, RoutedEventArgs e)
    {
        await OpenFolderAsync(@"C:\UTM", "UTM");
    }

    private async void OnOpenFolderPluginConfigsClicked(object sender, RoutedEventArgs e)
    {
        await OpenFolderAsync(Path.Combine(_cashServerBase, "PluginConfigs"), "PluginConfigs");
    }

    private async void OnOpenUtmTransportLogClicked(object sender, RoutedEventArgs e)
    {
        await OpenLogAsync(@"C:\UTM\transporter\l\transport_info.log", description: "UTM transport_info.log");
    }

    private async Task OpenPatternLogAsync(string pattern, string friendlyName)
    {
        var dir = Path.Combine(_cashServerBase, "Logs");
        if (!Directory.Exists(dir))
        {
            ShowNotification($"Папка логов не найдена ({dir})", NotificationType.Error);
            return;
        }

        var today = DateTime.Today;
        var files = Directory.EnumerateFiles(dir, "*.log", SearchOption.TopDirectoryOnly)
            .Where(f => Path.GetFileName(f).Contains(pattern, StringComparison.OrdinalIgnoreCase))
            .Where(f => File.GetLastWriteTime(f).Date == today)
            .OrderByDescending(File.GetLastWriteTime)
            .ToList();

        var file = files.FirstOrDefault();
        if (file == null)
        {
            ShowNotification($"{friendlyName}: файл за сегодня не найден", NotificationType.Warning);
            return;
        }

        await OpenInEditorAsync(file, friendlyName);
    }

    private async Task OpenLogAsync(string path, string description)
    {
        var today = DateTime.Today;
        if (!File.Exists(path) || File.GetLastWriteTime(path).Date != today)
        {
            ShowNotification($"{description}: файл за сегодня не найден", NotificationType.Warning);
            return;
        }

        await OpenInEditorAsync(path, description);
    }

    private Task OpenInEditorAsync(string path, string name)
    {
        return Task.Run(() =>
        {
            var editorPath = GetEditorPath();
            var psi = new ProcessStartInfo
            {
                FileName = editorPath,
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            };

            try
            {
                Process.Start(psi);
                _logger.Information("Открыт лог {Name}: {Path}", name, path);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Не удалось открыть файл {Path}", path);
                Dispatcher.Invoke(() =>
                    ShowNotification($"Не удалось открыть {name}: {ex.Message}", NotificationType.Error));
            }
        });
    }

    private string GetEditorPath()
    {
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Notepad++", "notepad++.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Notepad++", "notepad++.exe")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return "notepad.exe";
    }

    private Task OpenFolderAsync(string path, string name)
    {
        return Task.Run(() =>
        {
            if (!Directory.Exists(path))
            {
                Dispatcher.Invoke(() => ShowNotification($"{name}: папка не найдена ({path})", NotificationType.Warning));
                _logger.Warning("{Name}: папка не найдена {Path}", name, path);
                return;
            }

            var psi = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            };

            try
            {
                Process.Start(psi);
                _logger.Information("Открыта папка {Name}: {Path}", name, path);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Не удалось открыть папку {Path}", path);
                Dispatcher.Invoke(() => ShowNotification($"Не удалось открыть {name}: {ex.Message}", NotificationType.Error));
            }
        });
    }

    private async Task DownloadFileWithProgressAsync(string url, string destinationPath, IProgress<int> progress, CancellationToken token)
    {
        using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token);
        response.EnsureSuccessStatusCode();

        var total = response.Content.Headers.ContentLength ?? -1L;
        var canReport = total > 0 && progress != null;
        const int bufferSize = 81920;

        await using var contentStream = await response.Content.ReadAsStreamAsync(token);
        await using var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, useAsync: true);

        var buffer = new byte[bufferSize];
        long totalRead = 0;
        int lastPercent = 0;

        while (true)
        {
            var read = await contentStream.ReadAsync(buffer.AsMemory(0, buffer.Length), token);
            if (read == 0)
                break;

            await fileStream.WriteAsync(buffer.AsMemory(0, read), token);
            totalRead += read;

            if (canReport)
            {
                var percent = (int)Math.Clamp((totalRead * 100L) / total, 0, 100);
                if (percent != lastPercent && (percent % 5 == 0 || percent == 100))
                {
                    lastPercent = percent;
                    progress?.Report(percent);
                }
            }
        }

        progress?.Report(100);
    }

    private Task<int> KillIikoFrontAsync()
    {
        return Task.Run(() =>
        {
            var processes = Process.GetProcessesByName("iikoFront.Net");
            var killed = 0;
            foreach (var proc in processes)
            {
                try
                {
                    var pid = proc.Id;
                    proc.Kill(true);
                    proc.WaitForExit(5000);
                    killed++;
                    _logger.Information("Процесс iikoFront.Net {Pid} завершён", pid);
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Не удалось завершить iikoFront.Net PID {Pid}", proc.Id);
                }
            }
            return killed;
        });
    }

    private async void OnConfigureTlsClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Настройка...";
        try
        {
            _logger.Information("Настройка TLS 1.2: старт");
            var result = await _tlsConfigurator.EnableTls12Async();
            if (result.Ok)
            {
                ShowNotification("TLS 1.2 включён для клиента и сервера", NotificationType.Success);
                _logger.Information("Настройка TLS 1.2: завершено успешно");
            }
            else
            {
                ShowNotification($"TLS 1.2: не удалось применить ({result.Message})", NotificationType.Error);
                _logger.Warning("Настройка TLS 1.2: ошибка применения. Детали: {Detail}", result.Message);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка настройки TLS 1.2");
            ShowNotification($"Ошибка TLS 1.2: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Настройка TLS 1.2";
            button.IsEnabled = true;
        }
    }

    private void OnDisableSecurityClicked(object sender, RoutedEventArgs e)
    {
        ShowConfirm(
            "Отключение защитника и брандмауэра",
            "Вы уверены, что хотите отключить Windows Defender и брандмауэр? (требуются админ-права)",
            async () =>
            {
                try
                {
                    using var cts = new CancellationTokenSource();
                    var ct = cts.Token;
                    ShowDefenderTip(true);

                    ShowNotification("Откройте 'Защита от вирусов' и отключите защиту", NotificationType.Info);
                    _logger.Information("Запуск окна Defender...");
                    try
                    {
                         Process.Start(new ProcessStartInfo
                         {
                             FileName = "windowsdefender://threatsettings",
                             UseShellExecute = true
                         });
                    }
                    catch 
                    {
                        Process.Start(new ProcessStartInfo { FileName = "ms-settings:windowsdefender", UseShellExecute = true });
                    }
                    await Task.Run(async () =>
                    {
                         var processNames = new[] { "SecHealthUI", "WindowsSecurityHealthUI", "WindowsSecurity" };
                         var windowTitles = new[] { "Безопасность Windows", "Windows Security", "Центр безопасности защитника Windows" };
                         await Task.Delay(1500, ct);
                         await Task.Delay(3000, ct);
                         _logger.Information("Ожидание открытия окна Defender...");
                         while (!HasVisibleWindow(processNames, windowTitles))
                         {
                             ct.ThrowIfCancellationRequested();
                             await Task.Delay(2000, ct); 
                         }

                         _logger.Information("Окно Defender открыто. Ожидание закрытия...");
                         while (HasVisibleWindow(processNames, windowTitles))
                         {
                             ct.ThrowIfCancellationRequested();
                             await Task.Delay(1000, ct);
                         }
                         
                         _logger.Information("Окно Defender закрыто.");
                    }, ct);

                    ShowDefenderTip(false);

                    ShowNotification("Вырезаем защитник...", NotificationType.Warning);
                    var def = await _securityService.DisableDefenderAsync();
                    _logger.Information("Отключение защитника: {Result}. Вывод: {Output}", def.Ok ? "успех" : "ошибка", def.Output?.Trim());

                    ShowNotification("Вырезаем брандмауэр...", NotificationType.Warning);
                    var fw = await _securityService.DisableFirewallAsync();
                    _logger.Information("Отключение брандмауэра: {Result}. Вывод: {Output}", fw.Ok ? "успех" : "ошибка", fw.Output?.Trim());
                }
                catch (OperationCanceledException)
                {
                    ShowDefenderTip(false);
                    ShowNotification("Операция отменена", NotificationType.Warning);
                }
                catch (Exception ex)
                {
                    ShowDefenderTip(false);
                    _logger.Error(ex, "Ошибка при отключении защитника/брандмауэра");
                    ShowNotification($"Ошибка при отключении защиты: {ex.Message}", NotificationType.Error);
                }
            });
    }

    private DefenderTipWindow? _defenderTipWindow;

    private void ShowDefenderTip(bool show)
    {
        Dispatcher.Invoke(() =>
        {
            if (show)
            {
                if (_defenderTipWindow == null || !_defenderTipWindow.IsLoaded)
                {
                    _defenderTipWindow = new DefenderTipWindow
                    {
                        Owner = this,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner
                    };
                }
                _defenderTipWindow.Show();
            }
            else
            {
                _defenderTipWindow?.Close();
                _defenderTipWindow = null;
            }
        });
    }
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private static bool HasVisibleWindow(string[] processNames, string[] titles)
    {
        var pids = Process.GetProcesses()
            .Where(p => processNames.Contains(p.ProcessName, StringComparer.OrdinalIgnoreCase))
            .Select(p => p.Id)
            .ToHashSet();

        var found = false;
        EnumWindows((hwnd, lParam) =>
        {
            if (!IsWindowVisible(hwnd)) return true;
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pids.Contains((int)pid))
            {
                found = true;
                return false;
            }
            var length = GetWindowTextLength(hwnd);
            if (length > 0)
            {
                var sb = new StringBuilder(length + 1);
                GetWindowText(hwnd, sb, sb.Capacity);
                var windowTitle = sb.ToString();
                if (titles.Any(t => string.Equals(windowTitle, t, StringComparison.OrdinalIgnoreCase)))
                {
                    found = true;
                    return false;
                }
            }

            return true;
        }, IntPtr.Zero);

        return found;
    }

    private void UpdateTouchButtonVisual()
    {
        if (_isTouchEnabled == true)
        {
            TouchToggleButton.Content = "Отключить сенсор";
            TouchToggleButton.Style = _cachedDangerButtonStyle;
        }
        else if (_isTouchEnabled == false)
        {
            TouchToggleButton.Content = "Включить сенсор";
            TouchToggleButton.Style = _cachedPrimaryButtonStyle;
        }
        else
        {
            TouchToggleButton.Content = "Сенсор неизвестен";
            TouchToggleButton.Style = _cachedBaseButtonStyle;
            TouchToggleButton.IsEnabled = false;
        }
    }

    private void OnTopBarMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left)
        {
            if (IsClickOnButton(e.OriginalSource as DependencyObject))
            {
                return;
            }

            if (e.ClickCount == 2)
            {
                ToggleWindowState();
                return;
            }

            try
            {
                if (WindowState == WindowState.Maximized)
                {
                    RestoreAndDragFromMaximized(e);
                    return;
                }

                DragMove();
            }
            catch
            {
            }
        }
    }

    private void RestoreAndDragFromMaximized(MouseButtonEventArgs e)
    {
        var mousePos = e.GetPosition(this);
        var percentX = mousePos.X / ActualWidth;
        var percentY = mousePos.Y / ActualHeight;
        var screenPos = PointToScreen(mousePos);

        WindowState = WindowState.Normal;

        var targetWidth = _restoreWidth > 0 ? _restoreWidth : RestoreBounds.Width;
        var targetHeight = _restoreHeight > 0 ? _restoreHeight : RestoreBounds.Height;

        if (targetWidth > 0 && targetHeight > 0)
        {
            Width = targetWidth;
            Height = targetHeight;
        }

        Left = screenPos.X - (targetWidth * percentX);
        Top = screenPos.Y - (targetHeight * percentY);

        Dispatcher.BeginInvoke(new Action(() =>
        {
            try
            {
                DragMove();
            }
            catch
            {
            }
        }), DispatcherPriority.Input);
    }

    private static bool IsClickOnButton(DependencyObject? source)
    {
        var current = source;
        while (current != null)
        {
            if (current is Button)
            {
                return true;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private void OnWindowLocationChanged(object? sender, EventArgs e) => CaptureRestoreBounds();

    private void OnWindowSizeChanged(object sender, SizeChangedEventArgs e) => CaptureRestoreBounds();

    private void CaptureRestoreBounds()
    {
        if (WindowState != WindowState.Normal)
        {
            return;
        }

        _restoreWidth = ActualWidth > 0 ? ActualWidth : Width;
        _restoreHeight = ActualHeight > 0 ? ActualHeight : Height;
        _restoreLeft = double.IsNaN(Left) ? RestoreBounds.Left : Left;
        _restoreTop = double.IsNaN(Top) ? RestoreBounds.Top : Top;
    }

    private void OnMinimizeClicked(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void OnToggleMaximizeClicked(object sender, RoutedEventArgs e) => ToggleWindowState();

    private void OnCloseClicked(object sender, RoutedEventArgs e) => Close();

    private void ToggleWindowState()
    {
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private void AdjustToWorkArea()
    {
        var workArea = GetMonitorWorkArea();

        MaxHeight = workArea.Height;
        MaxWidth = workArea.Width;

        if (WindowState == WindowState.Maximized)
        {
            Left = workArea.Left;
            Top = workArea.Top;
            Width = workArea.Width;
            Height = workArea.Height;
            return;
        }

        Width = Math.Min(Width, workArea.Width);
        Height = Math.Min(Height, workArea.Height);

        if (Left < workArea.Left) Left = workArea.Left;
        if (Top < workArea.Top) Top = workArea.Top;
        if (Left + Width > workArea.Right) Left = workArea.Right - Width;
        if (Top + Height > workArea.Bottom) Top = workArea.Bottom - Height;
    }

    private Rect GetMonitorWorkArea()
    {
        var handle = new WindowInteropHelper(this).Handle;
        var monitor = MonitorFromWindow(handle, MonitorDefaultToNearest);

        var info = new MONITORINFO
        {
            cbSize = Marshal.SizeOf<MONITORINFO>()
        };

        if (!GetMonitorInfo(monitor, ref info))
        {
            var area = SystemParameters.WorkArea;
            return new Rect(area.Left, area.Top, area.Width, area.Height);
        }

        var (scaleX, scaleY) = GetDpiScale();
        var work = info.rcWork;
        var left = work.Left / scaleX;
        var top = work.Top / scaleY;
        var width = (work.Right - work.Left) / scaleX;
        var height = (work.Bottom - work.Top) / scaleY;
        return new Rect(left, top, width, height);
    }

    private (double scaleX, double scaleY) GetDpiScale()
    {
        var source = PresentationSource.FromVisual(this);
        if (source?.CompositionTarget is { } target)
        {
            return (target.TransformToDevice.M11, target.TransformToDevice.M22);
        }

        return (1.0, 1.0);
    }

    private const int MonitorDefaultToNearest = 0x00000002;
    private const int WmGetMinMaxInfo = 0x0024;

    [DllImport("User32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int dwFlags);

    [DllImport("User32.dll", SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmGetMinMaxInfo)
        {
            HandleMinMaxInfo(hwnd, lParam);
            handled = false;
        }

        return IntPtr.Zero;
    }

    private void HandleMinMaxInfo(IntPtr hwnd, IntPtr lParam)
    {
        var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
        var monitor = MonitorFromWindow(hwnd, MonitorDefaultToNearest);

        if (monitor != IntPtr.Zero)
        {
            var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
            if (GetMonitorInfo(monitor, ref info))
            {
                var work = info.rcWork;
                var monitorArea = info.rcMonitor;

                mmi.ptMaxPosition.X = work.Left - monitorArea.Left;
                mmi.ptMaxPosition.Y = work.Top - monitorArea.Top;
                mmi.ptMaxSize.X = work.Right - work.Left;
                mmi.ptMaxSize.Y = work.Bottom - work.Top;
            }
        }

        Marshal.StructureToPtr(mmi, lParam, true);
    }

    private void AdjustResizeBorder()
    {
        var chrome = WindowChrome.GetWindowChrome(this);
        if (chrome is null)
        {
            return;
        }

        chrome.ResizeBorderThickness = WindowState == WindowState.Maximized
            ? new Thickness(0)
            : _defaultResizeBorder;
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_hideToTrayEnabled && !_isExiting)
        {
            e.Cancel = true;
            EnsureTrayIcon();
            Hide();
            ShowInTaskbar = false;
            WindowState = WindowState.Minimized;
            return;
        }

        base.OnClosing(e);
    }

    private void EnsureTrayIcon()
    {
        if (_trayIcon != null)
        {
            _trayIcon.Visible = true;
            return;
        }

        try
        {
            var icon = LoadTrayIcon();
            var menu = new WinForms.ContextMenuStrip
            {
                ShowImageMargin = false,
                BackColor = ToDrawingColor(FindResource("PanelBrush") as Brush, SDColor.FromArgb(28, 28, 34)),
                ForeColor = ToDrawingColor(FindResource("TextPrimaryBrush") as Brush, SDColor.White),
                Font = new System.Drawing.Font("Segoe UI", 9f)
            };
            menu.Renderer = new TrayMenuRenderer(
                ToDrawingColor(FindResource("PanelBrush") as Brush, SDColor.FromArgb(28, 28, 34)),
                ToDrawingColor(FindResource("BorderBrushMuted") as Brush, SDColor.FromArgb(64, 64, 72)),
                ToDrawingColor(FindResource("AccentBrush") as Brush, SDColor.FromArgb(56, 166, 118)),
                ToDrawingColor(FindResource("TextPrimaryBrush") as Brush, SDColor.White));

            menu.Items.Add("Открыть", null, (_, _) => ShowFromTray());
            menu.Items.Add("Закрыть", null, (_, _) => ExitFromTray());

            _trayIcon = new WinForms.NotifyIcon
            {
                Icon = icon,
                Visible = true,
                Text = "Bobrus",
                ContextMenuStrip = menu
            };
            _trayIcon.DoubleClick += (_, _) => ShowFromTray();
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось создать иконку в трее");
        }
    }

    private System.Drawing.Icon LoadTrayIcon()
    {
        try
        {
            var trayIconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            if (File.Exists(trayIconPath))
            {
                return new System.Drawing.Icon(trayIconPath);
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось загрузить пользовательскую иконку трея");
        }

        return System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty)
               ?? System.Drawing.SystemIcons.Application;
    }

    private static SDColor ToDrawingColor(Brush? brush, SDColor fallback)
    {
        if (brush is SolidColorBrush solid)
        {
            return SDColor.FromArgb(solid.Color.A, solid.Color.R, solid.Color.G, solid.Color.B);
        }
        return fallback;
    }

    private sealed class TrayMenuRenderer : WinForms.ToolStripProfessionalRenderer
    {
        private readonly SDColor _background;
        private readonly SDColor _border;
        private readonly SDColor _highlight;
        private readonly SDColor _text;

        public TrayMenuRenderer(SDColor background, SDColor border, SDColor highlight, SDColor text)
        {
            _background = background;
            _border = border;
            _highlight = highlight;
            _text = text;
        }

        protected override void OnRenderToolStripBorder(System.Windows.Forms.ToolStripRenderEventArgs e)
        {
            using var pen = new System.Drawing.Pen(_border);
            e.Graphics.DrawRectangle(pen, 0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
        }

        protected override void OnRenderMenuItemBackground(System.Windows.Forms.ToolStripItemRenderEventArgs e)
        {
            var rect = new System.Drawing.Rectangle(System.Drawing.Point.Empty, e.Item.Size);
            var fill = e.Item.Selected ? _highlight : _background;
            using var brush = new System.Drawing.SolidBrush(fill);
            e.Graphics.FillRectangle(brush, rect);
        }

        protected override void OnRenderItemText(System.Windows.Forms.ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = _text;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderSeparator(System.Windows.Forms.ToolStripSeparatorRenderEventArgs e)
        {
            using var brush = new System.Drawing.SolidBrush(_border);
            var y = e.Item.Height / 2;
            e.Graphics.FillRectangle(brush, 4, y, e.Item.Width - 8, 1);
        }
    }

    private void ShowFromTray()
    {
        Dispatcher.Invoke(() =>
        {
            Show();
            ShowInTaskbar = true;
            WindowState = WindowState.Normal;
            Activate();
        });
    }

    private void ExitFromTray()
    {
        _isExiting = true;
        _trayIcon?.Dispose();
        _trayIcon = null;
        Dispatcher.Invoke(Close);
    }

    private void ShowConfirm(string title, string message, Action onConfirm)
    {
        ConfirmTitle.Text = title;
        ConfirmMessage.Text = message;
        _pendingConfirmAction = onConfirm;
        ConfirmOverlay.Visibility = Visibility.Visible;
    }

    private void HideConfirm()
    {
        ConfirmOverlay.Visibility = Visibility.Collapsed;
        _pendingConfirmAction = null;
    }

    private void OnConfirmCancelClicked(object sender, RoutedEventArgs e) => HideConfirm();

    private void OnConfirmAcceptClicked(object sender, RoutedEventArgs e)
    {
        var action = _pendingConfirmAction;
        HideConfirm();
        action?.Invoke();
    }

    private async Task RunUpdateCheckAsync(bool isAuto)
    {
        if (_updateCheckInProgress)
        {
            return;
        }

        _updateCheckInProgress = true;
        var willShutdown = false;
        if (!isAuto)
        {
            CheckUpdatesButton.IsEnabled = false;
            _logger.Information("Запрос на проверку обновлений.");
            ShowNotification("Проверяем обновления...", NotificationType.Info);
            SetGlobalProgress("Проверяем обновления...", null);
        }

        try
        {
            var checkResult = await _updateService.CheckForUpdatesAsync(CancellationToken.None);
            if (!checkResult.IsUpdateAvailable || checkResult.Update is null)
            {
                _logger.Information("Обновлений нет: {Message}", checkResult.Message);
                if (!isAuto)
                {
                    ShowNotification(checkResult.Message, NotificationType.Info);
                    SetGlobalProgress(null, null);
                }
                return;
            }

            if (!isAuto)
            {
                ShowNotification($"Найдена версия {checkResult.Update.LatestVersion}. Скачиваем...", NotificationType.Info);
                SetGlobalProgress($"Загрузка {checkResult.Update.LatestVersion}", 0);
            }
            else
            {
                _logger.Information("Автообновление: найдена версия {Version}", checkResult.Update.LatestVersion);
            }

            var packagePath = _updateService.GetPackageCachePath(checkResult.Update.LatestVersion, checkResult.Update.Asset.Name);
            IProgress<double>? progress = null;
            if (!isAuto)
            {
                progress = new Progress<double>(p =>
                {
                    var percent = Math.Clamp((int)(p * 100), 0, 100);
                    UpdateStatusText.Text = string.Empty;
                    SetGlobalProgress($"Загрузка {checkResult.Update.LatestVersion}", percent / 100.0);
                });
            }

            await _updateService.DownloadAssetAsync(checkResult.Update.Asset, packagePath, progress, CancellationToken.None);

            if (!isAuto)
            {
                SetGlobalProgress("Распаковка...", null);
            }
            var extractedFolder = _updateService.ExtractPackage(packagePath, checkResult.Update.LatestVersion);
            if (!isAuto)
            {
                SetGlobalProgress(null, null);
            }

            var process = _updateService.StartApplyUpdate(extractedFolder, Process.GetCurrentProcess().Id);
            if (process is null)
            {
                _logger.Error("Не найден исполняемый файл в пакете обновления. Ожидалось {Expected}.", "Bobrus.exe");
                if (!isAuto)
                {
                    ShowNotification("Не удалось запустить обновление: нет исполняемого файла.", NotificationType.Error);
                    SetGlobalProgress(null, null);
                }
                return;
            }

            _logger.Information("Установка обновления запущена из {Folder}. Процесс {Pid}.", extractedFolder, process.Id);
            if (!isAuto)
            {
                ShowNotification("Устанавливаем обновление...", NotificationType.Info);
                SetGlobalProgress("Устанавливаем обновление...", null);
            }
            willShutdown = true;
            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка при выполнении обновления.");
            if (!isAuto)
            {
                ShowNotification($"Ошибка обновления: {ex.Message}", NotificationType.Error);
                SetGlobalProgress(null, null);
            }
        }
        finally
        {
            if (!willShutdown && !isAuto)
            {
                CheckUpdatesButton.IsEnabled = true;
                UpdateStatusText.Text = string.Empty;
                SetGlobalProgress(null, null);
            }

            _updateCheckInProgress = false;
        }
    }

    private void SetGlobalProgress(string? text, double? fraction)
    {
        if (GlobalProgressPanel == null || GlobalProgressBar == null || GlobalProgressText == null)
        {
            return;
        }

        if (text is null && fraction is null)
        {
            GlobalProgressPanel.Visibility = Visibility.Collapsed;
            GlobalProgressBar.IsIndeterminate = false;
            GlobalProgressBar.Value = 0;
            GlobalProgressText.Text = string.Empty;
            return;
        }

        GlobalProgressPanel.Visibility = Visibility.Visible;
        GlobalProgressText.Text = text ?? string.Empty;

        if (fraction.HasValue)
        {
            GlobalProgressBar.IsIndeterminate = false;
            GlobalProgressBar.Value = Math.Clamp(fraction.Value * 100, 0, 100);
        }
        else
        {
            GlobalProgressBar.IsIndeterminate = true;
        }
    }

    private enum NotificationType
    {
        Info,
        Success,
        Error,
        Warning
    }

    private void ShowNotification(string message, NotificationType type)
    {
        var accent = type switch
        {
            NotificationType.Success => (Brush)FindResource("AccentBrush"),
            NotificationType.Error => (Brush)FindResource("DangerBrush"),
            NotificationType.Warning => (Brush)FindResource("AccentBrush"),
            _ => (Brush)FindResource("AccentBlueBrush")
        };

        var container = new Border
        {
            Background = FindResource("PanelBrush") as Brush,
            CornerRadius = new CornerRadius(8),
            BorderBrush = FindResource("BorderBrushMuted") as Brush,
            BorderThickness = new Thickness(1),
            Padding = new Thickness(0),
            Margin = new Thickness(0, 0, 0, 8),
            Opacity = 0.97
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.Margin = new Thickness(10, 8, 10, 8);

        var stripe = new Border
        {
            Background = accent,
            CornerRadius = new CornerRadius(8, 0, 0, 8)
        };

        Grid.SetColumn(stripe, 0);
        grid.Children.Add(stripe);

        var textBlock = new TextBlock
        {
            Text = message,
            Foreground = FindResource("TextPrimaryBrush") as Brush,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(10, 0, 0, 0)
        };

        Grid.SetColumn(textBlock, 1);
        grid.Children.Add(textBlock);

        container.Child = grid;
        NotificationStack.Children.Insert(0, container);

        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(4) };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            NotificationStack.Children.Remove(container);
        };
        timer.Start();
    }

    private bool IsAutostartEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", writable: false);
            if (key == null) return false;

            var value = key.GetValue("Bobrus") as string;
            var exePath = Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
            return string.Equals(value?.Trim('"'), exePath, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Простая сериализация настроек, типы известны и не меняются.")]
    private AppSettings? LoadAppSettings()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                return JsonSerializer.Deserialize<AppSettings>(json);
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось прочитать настройки");
        }

        return null;
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Простая сериализация настроек, типы известны и не меняются.")]
    private void SaveAppSettings()
    {
        try
        {
            var settings = new AppSettings(
                _hideToTrayEnabled,
                _autostartEnabled,
                GetThemeName(),
                _showAllSections,
                _consoleVisible,
                _overlayEnabled,
                _overlayCrm,
                _overlayCashDesk,
                _overlayDisplay);
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось сохранить настройки");
        }
    }
    private void ReloadOverlaySettingsFromFile()
    {
        try
        {
            var settings = LoadAppSettings();
            if (settings == null) return;
            
            _overlayEnabled = settings.OverlayEnabled;
            _overlayCrm = settings.OverlayCrm?.Trim() ?? string.Empty;
            _overlayCashDesk = settings.OverlayCashDesk?.Trim() ?? string.Empty;
            _overlayDisplay = settings.OverlayDisplay ?? string.Empty;
            _suppressSettingsToggle = true;
            if (OverlayToggle != null)
            {
                OverlayToggle.IsChecked = _overlayEnabled;
            }
            if (OverlayCrmInput != null)
            {
                OverlayCrmInput.Text = _overlayCrm;
            }
            if (OverlayCashInput != null)
            {
                OverlayCashInput.Text = _overlayCashDesk;
            }
            UpdateOverlayDisplayLabel();
            ApplyOverlaySettingsVisibility();
            ApplyOverlayState();
            _suppressSettingsToggle = false;
            
            _logger.Information("Настройки оверлея перезагружены из файла: Enabled={Enabled}, CRM={Crm}, Cash={Cash}", 
                _overlayEnabled, _overlayCrm, _overlayCashDesk);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось перезагрузить настройки оверлея");
        }
    }

    private void ApplyOverlayState()
    {
        if (_overlayEnabled && HasOverlayValues())
        {
            StartOverlayRunner();
            SetOverlayAutostart(true);
        }
        else
        {
            StopOverlayRunner();
            SetOverlayAutostart(false);
        }
    }

    private void UpdateOverlayValues(string? crm, string? cashDesk)
    {
        _overlayCrm = crm?.Trim() ?? string.Empty;
        _overlayCashDesk = cashDesk?.Trim() ?? string.Empty;
    }

    private bool HasOverlayValues()
    {
        return !string.IsNullOrWhiteSpace(_overlayCrm) && !string.IsNullOrWhiteSpace(_overlayCashDesk);
    }

    private void ApplyOverlaySettingsVisibility()
    {
        var visibility = _overlayEnabled ? Visibility.Visible : Visibility.Collapsed;
        if (OverlaySettingsPanel != null)
        {
            OverlaySettingsPanel.Visibility = visibility;
        }

        if (OverlayDisplayPanel != null)
        {
            OverlayDisplayPanel.Visibility = visibility;
        }
    }

    private void UpdateOverlayDisplayLabel()
    {
        if (OverlayDisplayLabel == null) return;

        var screens = WinForms.Screen.AllScreens;
        var screen = screens.FirstOrDefault(s => string.Equals(s.DeviceName, _overlayDisplay, StringComparison.OrdinalIgnoreCase))
                     ?? WinForms.Screen.PrimaryScreen
                     ?? screens.FirstOrDefault();

        if (screen is null)
        {
            OverlayDisplayLabel.Text = "Экран не найден";
            return;
        }

        var index = Array.IndexOf(screens, screen);
        var name = screen.Primary ? "Основной" : "Экран";
        OverlayDisplayLabel.Text = $"{index + 1}: {name} {screen.Bounds.Width}x{screen.Bounds.Height}";
    }

    private void StartOverlayRunner()
    {
        try
        {
            if (!IsOverlayRunnerActive())
            {
                var exePath = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(exePath)) return;

                var psi = BuildOverlayStartInfo(exePath);
                Process.Start(psi);
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось запустить оверлей-процесс");
        }
    }

    private void StopOverlayRunner()
    {
    }

    private static bool IsOverlayRunnerActive()
    {
        try
        {
            return Mutex.TryOpenExisting("Global\\BobrusOverlayRunner", out _);
        }
        catch
        {
            return false;
        }
    }

    private void SetOverlayAutostart(bool enable)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", writable: true)
                             ?? Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
            if (key == null) return;

            if (enable)
            {
                var exePath = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(exePath)) return;
                key.SetValue("BobrusOverlay", $"\"{exePath}\" --overlay-runner");
            }
            else
            {
                key.DeleteValue("BobrusOverlay", false);
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось обновить автозапуск оверлея");
        }
    }

    private static ProcessStartInfo BuildOverlayStartInfo(string launcherPath)
    {
        var isDll = string.Equals(Path.GetExtension(launcherPath), ".dll", StringComparison.OrdinalIgnoreCase);
        if (isDll)
        {
            return new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{launcherPath}\" --overlay-runner",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }

        return new ProcessStartInfo
        {
            FileName = launcherPath,
            Arguments = "--overlay-runner",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
    }

    private void ApplySectionLayout()
    {
        if (SectionNav != null)
        {
            SectionNav.Visibility = _showAllSections ? Visibility.Collapsed : Visibility.Visible;
        }

        if (SectionNavContainer != null && SectionNavColumn != null)
        {
            if (_showAllSections)
            {
                SectionNavContainer.Visibility = Visibility.Collapsed;
                SectionNavColumn.Width = new GridLength(0);
                SectionNavColumn.MinWidth = 0;
            }
            else
            {
                SectionNavContainer.Visibility = Visibility.Visible;
                SectionNavColumn.Width = GridLength.Auto;
                SectionNavColumn.MinWidth = 130;
            }
        }

        if (_showAllSections)
        {
            SystemCard.Visibility = Visibility.Visible;
            IikoCard.Visibility = Visibility.Visible;
            ProgramsCard.Visibility = Visibility.Visible;
            NetworkCard.Visibility = Visibility.Visible;
            LogsCard.Visibility = Visibility.Visible;
            FoldersCard.Visibility = Visibility.Visible;
            PluginsCard.Visibility = Visibility.Visible;
        }
        else
        {
            SetSectionVisibility(_currentSection);
        }

        ApplySearch(SearchBox.Text);
    }

    private void ApplyConsoleVisibility()
    {
        if (ConsolePanel != null)
        {
            ConsolePanel.Visibility = _consoleVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        if (ConsoleColumn != null)
        {
            if (_consoleVisible)
            {
                ConsoleColumn.Width = new GridLength(1, GridUnitType.Star);
                ConsoleColumn.MinWidth = ConsoleColumnMinWidth;
                ConsoleColumn.MaxWidth = ConsoleColumnMaxWidth;
            }
            else
            {
                ConsoleColumn.Width = new GridLength(0);
                ConsoleColumn.MinWidth = 0;
                ConsoleColumn.MaxWidth = 0;
            }
        }
    }

    private void ApplyInitialSettingsToUi()
    {
        _suppressSettingsToggle = true;
        if (AutostartToggle != null)
        {
            AutostartToggle.IsChecked = _autostartEnabled;
        }

        if (HideToTrayToggle != null)
        {
            HideToTrayToggle.IsChecked = _hideToTrayEnabled;
        }

        if (ShowAllSectionsToggle != null)
        {
            ShowAllSectionsToggle.IsChecked = _showAllSections;
        }

        if (ConsoleToggle != null)
        {
            ConsoleToggle.IsChecked = _consoleVisible;
        }

        if (OverlayToggle != null)
        {
            OverlayToggle.IsChecked = _overlayEnabled;
        }

        if (OverlayCrmInput != null)
        {
            OverlayCrmInput.Text = _overlayCrm;
        }

        if (OverlayCashInput != null)
        {
            OverlayCashInput.Text = _overlayCashDesk;
        }

        UpdateOverlayDisplayLabel();
        ApplyOverlaySettingsVisibility();

        if (ThemeToggle != null)
        {
            ThemeToggle.IsChecked = _currentTheme == ThemeVariant.Dark;
        }

        ApplySectionLayout();
        ApplyConsoleVisibility();
        ApplyOverlayState();

        if (_isFirstRunSettings)
        {
            EnsureTrayIcon();
            SetAutostart(true);
            SaveAppSettings();
        }
        else if (_hideToTrayEnabled)
        {
            EnsureTrayIcon();
        }

        _suppressSettingsToggle = false;
    }

    private void SetSectionVisibility(Section section)
    {
        _currentSection = section;
        SystemCard.Visibility = section == Section.System ? Visibility.Visible : Visibility.Collapsed;

        IikoCard.Visibility = section == Section.Iiko ? Visibility.Visible : Visibility.Collapsed;
        ProgramsCard.Visibility = section == Section.Programs ? Visibility.Visible : Visibility.Collapsed;
        NetworkCard.Visibility = section == Section.Network ? Visibility.Visible : Visibility.Collapsed;
        LogsCard.Visibility = section == Section.Logs ? Visibility.Visible : Visibility.Collapsed;
        FoldersCard.Visibility = section == Section.Folders ? Visibility.Visible : Visibility.Collapsed;
        PluginsCard.Visibility = section == Section.Plugins ? Visibility.Visible : Visibility.Collapsed;
        UpdateSectionNavState(section);
    }

    private void UpdateSectionNavState(Section section)
    {
        SystemNavToggle.IsChecked = section == Section.System;

        NetworkNavToggle.IsChecked = section == Section.Network;
        IikoNavToggle.IsChecked = section == Section.Iiko;
        ProgramsNavToggle.IsChecked = section == Section.Programs;
        LogsNavToggle.IsChecked = section == Section.Logs;
        FoldersNavToggle.IsChecked = section == Section.Folders;
        PluginsNavToggle.IsChecked = section == Section.Plugins;
    }

    private void BuildSectionButtons()
    {
        _sectionButtons[Section.System] = ActionsPanel.Children.OfType<Button>().ToList();

        _sectionButtons[Section.Network] = NetworkActionsPanel.Children.OfType<Button>().ToList();
        _sectionButtons[Section.Iiko] = IikoActionsPanel.Children.OfType<Button>().ToList();
        _sectionButtons[Section.Programs] = ProgramsActionsPanel.Children.OfType<Button>().ToList();
        _sectionButtons[Section.Logs] = LogsActionsPanel.Children.OfType<Button>().ToList();
        _sectionButtons[Section.Folders] = FoldersActionsPanel.Children.OfType<Button>().ToList();
        _sectionButtons[Section.Plugins] = new List<Button> { InstallPluginButton };
    }

    private List<Button> GetButtonsForSection(Section section)
    {
        if (_sectionButtons.TryGetValue(section, out var buttons))
        {
            return buttons;
        }
        return new List<Button>();
    }

    private void OnSectionNavClicked(object sender, RoutedEventArgs e)
    {
        if (_showAllSections) return;
        if (sender is ToggleButton toggle &&
            Enum.TryParse<Section>(toggle.Tag?.ToString(), out var section))
        {
            SetSectionVisibility(section);
            ApplySearch(SearchBox.Text);
        }
    }

    private void OnOpenSetupWizardClicked(object sender, RoutedEventArgs e)
    {
        if (!IsAdministrator())
        {
            RelaunchAsAdmin(launchSetupWizard: true);
            return;
        }

        ShowSetupWizardWindow();
    }

    private void ShowSetupWizardWindow(bool autoResume = false)
    {
        if (_setupWizardWindow is { IsLoaded: true })
        {
            _setupWizardWindow.WindowState = WindowState.Normal;
            _setupWizardWindow.Show();
            _setupWizardWindow.Activate();
            return;
        }

        _setupWizardWindow = new SetupWizardWindow();
        if (IsLoaded)
        {
            _setupWizardWindow.Owner = this;
        }

        _setupWizardWindow.Closed += (_, _) => 
        {
            _setupWizardWindow = null;
            ReloadOverlaySettingsFromFile();
        };
        _setupWizardWindow.Show();
        _setupWizardWindow.Activate();

        if (autoResume)
        {
            _setupWizardWindow.ResumeSetup();
        }
    }

    private void ApplyTheme(ThemeVariant theme)
    {
        _currentTheme = theme;
        var palette = theme == ThemeVariant.Dark ? _darkTheme : _lightTheme;
        foreach (var pair in palette)
        {
            var brush = new SolidColorBrush(pair.Value);
            Application.Current.Resources[pair.Key] = brush;
        }

        if (ThemeToggle != null)
        {
            ThemeToggle.IsChecked = theme == ThemeVariant.Dark;
        }
    }

    private ThemeVariant ParseTheme(string? themeValue)
    {
        return string.Equals(themeValue, "dark", StringComparison.OrdinalIgnoreCase)
            ? ThemeVariant.Dark
            : ThemeVariant.Light;
    }

    private string GetThemeName() => _currentTheme == ThemeVariant.Dark ? "dark" : "light";

    private static Color ColorFromHex(string hex) => (Color)ColorConverter.ConvertFromString(hex)!;

    internal static bool IsAdministratorStatic() => IsAdministrator();

    private static bool IsAdministrator()
    {
        try
        {
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    internal static void RelaunchAsAdminStatic(bool launchSetupWizard = false) => RelaunchAsAdmin(launchSetupWizard);

    private static void RelaunchAsAdmin(bool launchSetupWizard = false)
    {
        try
        {
            var exePath = Process.GetCurrentProcess().MainModule?.FileName;
            if (string.IsNullOrWhiteSpace(exePath))
            {
                Log.Warning("RelaunchAsAdmin: не удалось получить путь к exe");
                return;
            }

            Log.Information("RelaunchAsAdmin: запуск с правами администратора, путь: {ExePath}, setupWizard: {SetupWizard}", exePath, launchSetupWizard);

            var psi = new ProcessStartInfo(exePath)
            {
                UseShellExecute = true,
                Verb = "runas",
                Arguments = launchSetupWizard ? SetupWizardArg : string.Empty
            };

            try 
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    Application.Current.Shutdown();
                });
                System.Threading.Thread.Sleep(300);
                
                var started = Process.Start(psi);
                if (started != null)
                {
                    Log.Information("RelaunchAsAdmin: процесс запущен, PID: {PID}", started.Id);
                }
                else
                {
                    Log.Warning("RelaunchAsAdmin: Process.Start вернул null");
                }
            }
            catch (System.ComponentModel.Win32Exception w32ex) when (w32ex.NativeErrorCode == 1223)
            {
                Log.Warning("RelaunchAsAdmin: пользователь отменил UAC запрос");
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "RelaunchAsAdmin: ошибка при перезапуске");
            Debug.WriteLine(ex);
        }
    }

    private void SetAutostart(bool enable)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", writable: true);
            if (key == null)
            {
                ShowNotification("Не удалось открыть раздел автозагрузки", NotificationType.Error);
                AutostartToggle.IsChecked = _autostartEnabled;
                return;
            }

            var exePath = Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
            if (enable)
            {
                key.SetValue("Bobrus", $"\"{exePath}\" {StartHiddenArg}");
            }
            else
            {
                key.DeleteValue("Bobrus", throwOnMissingValue: false);
            }

            _autostartEnabled = enable;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось изменить автозапуск");
            ShowNotification($"Не удалось изменить автозапуск: {ex.Message}", NotificationType.Error);
            AutostartToggle.IsChecked = _autostartEnabled;
        }
    }

    private void ApplyStartHidden()
    {
        if (!_startHidden)
        {
            return;
        }

        if (_hideToTrayEnabled)
        {
            EnsureTrayIcon();
            Hide();
            ShowInTaskbar = false;
            WindowState = WindowState.Minimized;
            return;
        }

        WindowState = WindowState.Minimized;
    }

    private enum Section
    {
        System,
        SetupWizard,
        Network,
        Iiko,
        Programs,
        Logs,
        Folders,
        Plugins
    }

    private enum ThemeVariant
    {
        Light,
        Dark
    }

}
