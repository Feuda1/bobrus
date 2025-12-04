using System;
using System.Diagnostics;
using System.Net.Http;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using Serilog;
using ILogger = Serilog.ILogger;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Interop;
using Bobrus.App.Services;
using Serilog.Events;

namespace Bobrus.App;

public partial class MainWindow : Window
{
    private readonly ILogger _logger = Log.ForContext<MainWindow>();
    private readonly HttpClient _httpClient = new();
    private readonly UpdateService _updateService;
    private readonly TouchDeviceManager _touchManager = new();
    private readonly CleaningService _cleaningService = new();
    private Action? _pendingConfirmAction;
    private bool? _isTouchEnabled;

    public MainWindow()
    {
        InitializeComponent();
        _updateService = new UpdateService(_httpClient);
        VersionText.Text = $"v{_updateService.CurrentVersion}";
        _logger.Information("Bobrus запущен. Текущая версия {Version}", _updateService.CurrentVersion);
        Loaded += OnLoaded;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        AdjustToWorkArea();
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);
        AdjustToWorkArea();
    }

    private void OnRebootClicked(object sender, RoutedEventArgs e)
    {
        ShowConfirm("Перезагрузка",
            "Перезагрузить компьютер сейчас? Все несохраненные данные будут потеряны.",
            StartReboot);
    }

    private async void OnCheckUpdatesClicked(object sender, RoutedEventArgs e)
    {
        var willShutdown = false;
        CheckUpdatesButton.IsEnabled = false;
        _logger.Information("Запрос на проверку обновлений.");
        ShowNotification("Проверяем обновления...", NotificationType.Info);
        AppPaths.CleanupOldUpdates();

        try
        {
            var checkResult = await _updateService.CheckForUpdatesAsync(CancellationToken.None);
            if (!checkResult.IsUpdateAvailable || checkResult.Update is null)
            {
                _logger.Information("Обновлений нет: {Message}", checkResult.Message);
                ShowNotification(checkResult.Message, NotificationType.Info);
                return;
            }

            ShowNotification($"Найдена версия {checkResult.Update.LatestVersion}. Скачиваем...", NotificationType.Info);
            _logger.Information("Найдено обновление {Version}. Начинаем загрузку {Asset}.", checkResult.Update.LatestVersion, checkResult.Update.Asset.Name);
            var packagePath = _updateService.GetPackageCachePath(checkResult.Update.LatestVersion, checkResult.Update.Asset.Name);
            var progress = new Progress<double>(p =>
            {
                var percent = Math.Clamp((int)(p * 100), 0, 100);
                UpdateStatusText.Text = string.Empty;
                if (percent % 10 == 0)
                {
                    ShowNotification($"Скачивание {checkResult.Update.LatestVersion}: {percent}%", NotificationType.Info);
                }
            });

            await _updateService.DownloadAssetAsync(checkResult.Update.Asset, packagePath, progress, CancellationToken.None);

            ShowNotification("Распаковка пакета...", NotificationType.Info);
            var extractedFolder = _updateService.ExtractPackage(packagePath, checkResult.Update.LatestVersion);

            var process = _updateService.StartApplyUpdate(extractedFolder, Process.GetCurrentProcess().Id);
            if (process is null)
            {
                _logger.Error("Не найден исполняемый файл в пакете обновления. Ожидалось {Expected}.", "Bobrus.exe");
                ShowNotification("Не удалось запустить обновление: нет исполняемого файла.", NotificationType.Error);
                return;
            }

            _logger.Information("Установка обновления запущена из {Folder}. Процесс {Pid}.", extractedFolder, process.Id);
            ShowNotification("Устанавливаем обновление...", NotificationType.Info);
            willShutdown = true;
            Application.Current.Shutdown();
            return;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка при выполнении обновления.");
            ShowNotification($"Ошибка обновления: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            if (!willShutdown)
            {
                CheckUpdatesButton.IsEnabled = true;
                UpdateStatusText.Text = string.Empty;
            }
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        _httpClient.Dispose();
        UiLogBuffer.OnLog -= OnUiLog;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        UiLogBuffer.OnLog += OnUiLog;
        await RefreshTouchStateAsync();
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
        TouchToggleButton.Content = "Сенсор: проверка...";
        TouchToggleButton.Style = FindResource("BaseButton") as Style;

        var devices = await _touchManager.GetTouchDevicesAsync();
        if (devices.Count == 0)
        {
            TouchToggleButton.Content = "Сенсор не найден";
            TouchToggleButton.IsEnabled = false;
            _isTouchEnabled = null;
            return;
        }

        _isTouchEnabled = devices.Any(d => d.IsEnabled);
        UpdateTouchButtonVisual();
        TouchToggleButton.IsEnabled = true;
    }

    private async void OnTouchToggleClicked(object sender, RoutedEventArgs e)
    {
        TouchToggleButton.IsEnabled = false;
        TouchToggleButton.Content = "Применение...";
        try
        {
            var targetState = !(_isTouchEnabled ?? true);
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

    private void OnUiLog(LogEvent logEvent)
    {
        Dispatcher.InvokeAsync(() =>
        {
            var paragraph = new System.Windows.Documents.Paragraph { Margin = new Thickness(0, 0, 0, 4) };
            var run = new System.Windows.Documents.Run($"[{logEvent.Timestamp:HH:mm:ss}] {logEvent.RenderMessage()}");
            run.Foreground = GetBrushForLevel(logEvent.Level);
            paragraph.Inlines.Add(run);
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        }, DispatcherPriority.Background);
    }

    private Brush GetBrushForLevel(LogEventLevel level) =>
        level switch
        {
            LogEventLevel.Error or LogEventLevel.Fatal => FindResource("DangerBrush") as Brush ?? Brushes.Red,
            LogEventLevel.Warning => FindResource("AccentBlueBrush") as Brush ?? Brushes.Orange,
            LogEventLevel.Debug => FindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray,
            _ => FindResource("TextPrimaryBrush") as Brush ?? Brushes.White
        };

    private async void OnRestartTouchClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Перезапуск...";
        try
        {
            var ok = await _touchManager.RestartTouchAsync();
            if (!ok)
            {
                ShowNotification("Сенсор не найден для перезапуска", NotificationType.Error);
            }
            else
            {
                ShowNotification("Сенсор перезапущен", NotificationType.Success);
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
        }
    }

    private async void OnCleanupClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Очистка...";
        try
        {
            var results = await _cleaningService.RunCleanupAsync();
            var freed = results.Sum(r => r.BytesFreed);
            foreach (var r in results)
            {
                _logger.Information("Очистка {Name}: освобождено {Bytes} байт", r.Name, r.BytesFreed);
            }
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

    private void UpdateTouchButtonVisual()
    {
        if (_isTouchEnabled == true)
        {
            TouchToggleButton.Content = "Отключить сенсор";
            TouchToggleButton.Style = FindResource("DangerButton") as Style;
        }
        else if (_isTouchEnabled == false)
        {
            TouchToggleButton.Content = "Включить сенсор";
            TouchToggleButton.Style = FindResource("PrimaryButton") as Style;
        }
        else
        {
            TouchToggleButton.Content = "Сенсор неизвестен";
            TouchToggleButton.Style = FindResource("BaseButton") as Style;
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
                    var mousePos = e.GetPosition(this);
                    var percentX = mousePos.X / ActualWidth;
                    var percentY = mousePos.Y / ActualHeight;
                    var screenPos = PointToScreen(mousePos);

                    WindowState = WindowState.Normal;

                    Left = screenPos.X - (RestoreBounds.Width * percentX);
                    Top = screenPos.Y - (RestoreBounds.Height * percentY);
                }

                DragMove();
            }
            catch
            {
                // ignore
            }
        }
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

    private void OnMinimizeClicked(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void OnToggleMaximizeClicked(object sender, RoutedEventArgs e) => ToggleWindowState();

    private void OnCloseClicked(object sender, RoutedEventArgs e) => Close();

    private void ToggleWindowState()
    {
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private void AdjustToWorkArea()
    {
        var area = SystemParameters.WorkArea;

        MaxHeight = area.Height;
        MaxWidth = area.Width;

        if (WindowState == WindowState.Maximized)
        {
            Left = area.Left;
            Top = area.Top;
            Width = area.Width;
            Height = area.Height;
        }
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

    private enum NotificationType
    {
        Info,
        Success,
        Error
    }

    private void ShowNotification(string message, NotificationType type)
    {
        var accent = type switch
        {
            NotificationType.Success => (Brush)FindResource("AccentBrush"),
            NotificationType.Error => (Brush)FindResource("DangerBrush"),
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
}
