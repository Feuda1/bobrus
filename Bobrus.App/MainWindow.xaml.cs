using System;
using System.Diagnostics;
using System.Net.Http;
using System.IO;
using System.Linq;
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
using System.Text.RegularExpressions;

namespace Bobrus.App;

public partial class MainWindow : Window
{
    private readonly ILogger _logger = Log.ForContext<MainWindow>();
    private readonly HttpClient _httpClient = new();
    private readonly UpdateService _updateService;
    private readonly TouchDeviceManager _touchManager = new();
    private readonly CleaningService _cleaningService = new();
    private readonly ComPortManager _comPortManager = new();
    private readonly SecurityService _securityService = new();
    private readonly TlsConfigurator _tlsConfigurator = new();
    private readonly PrintSpoolService _printSpoolService = new();
    private const string IikoFrontExePath = @"C:\Program Files\iiko\iikoRMS\Front.Net\iikoFront.Net.exe";
    private const string IikoCardUrl = "https://iiko.biz/ru-RU/About/DownloadPosInstaller?useRc=False";
    private Action? _pendingConfirmAction;
    private bool? _isTouchEnabled;
    private List<Button> _actionButtons = new();

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
            "Вы уверены, что хотите перезагрузить компьютер?",
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
        _actionButtons = ActionsPanel.Children.OfType<Button>()
            .Concat(IikoActionsPanel.Children.OfType<Button>())
            .ToList();
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
            var timeRun = new System.Windows.Documents.Run($"[{logEvent.Timestamp:HH:mm:ss}] ")
            {
                Foreground = FindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray
            };

            paragraph.Inlines.Add(timeRun);
            foreach (var inline in BuildMessageInlines(logEvent))
            {
                paragraph.Inlines.Add(inline);
            }
            LogRichTextBox.Document.Blocks.Add(paragraph);
            LogRichTextBox.ScrollToEnd();
        }, DispatcherPriority.Background);
    }

    private IEnumerable<System.Windows.Documents.Inline> BuildMessageInlines(LogEvent logEvent)
    {
        var danger = FindResource("DangerBrush") as Brush ?? Brushes.Red;
        var accentOrange = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E7A04F"));
        var accentGreen = FindResource("AccentBrush") as Brush ?? Brushes.LimeGreen;
        var info = FindResource("TextPrimaryBrush") as Brush ?? Brushes.White;
        var secondary = FindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray;

        var message = logEvent.RenderMessage();
        if (logEvent.Level == LogEventLevel.Error || logEvent.Level == LogEventLevel.Fatal)
        {
            yield return new System.Windows.Documents.Run(message)
            {
                Foreground = danger,
                FontWeight = FontWeights.SemiBold
            };
            yield break;
        }

        var baseBrush = accentOrange;
        var pattern = new Regex(@"\d[\d\s]*байт|начало", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var lastIndex = 0;
        foreach (Match match in pattern.Matches(message))
        {
            if (match.Index > lastIndex)
            {
                var text = message.Substring(lastIndex, match.Index - lastIndex);
                yield return new System.Windows.Documents.Run(text)
                {
                    Foreground = baseBrush
                };
            }

            var highlightText = match.Value;
            yield return new System.Windows.Documents.Run(highlightText)
            {
                Foreground = accentGreen,
                FontWeight = FontWeights.SemiBold
            };

            lastIndex = match.Index + match.Length;
        }

        if (lastIndex < message.Length)
        {
            var tail = message.Substring(lastIndex);
            yield return new System.Windows.Documents.Run(tail)
            {
                Foreground = baseBrush
            };
        }
    }

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

    private async void OnRestartComClicked(object sender, RoutedEventArgs e)
    {
        var button = (Button)sender;
        button.IsEnabled = false;
        button.Content = "Перезапуск...";
        try
        {
            _logger.Information("Перезапуск COM: старт");
            var ok = await _comPortManager.RestartPortsAsync();
            if (!ok)
            {
                ShowNotification("COM-порты не найдены", NotificationType.Error);
                _logger.Warning("Перезапуск COM: устройства не найдены");
            }
            else
            {
                ShowNotification("COM-порты перезапущены", NotificationType.Success);
                _logger.Information("Перезапуск COM: завершено");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка перезапуска COM-портов");
            ShowNotification($"Ошибка перезапуска COM: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = "Перезапуск COM-портов";
            button.IsEnabled = true;
        }
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        ApplySearch(SearchBox.Text);
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
        if (q.Length == 0)
        {
            foreach (var btn in _actionButtons)
            {
                btn.Visibility = Visibility.Visible;
            }
            SystemCard.Visibility = Visibility.Visible;
            IikoCard.Visibility = Visibility.Visible;
            return;
        }

        foreach (var btn in _actionButtons)
        {
            var text = $"{btn.Content} {btn.Tag}".ToString().ToLowerInvariant();
            btn.Visibility = text.Contains(q) ? Visibility.Visible : Visibility.Collapsed;
        }

        var systemVisible = ActionsPanel.Children.OfType<Button>().Any(b => b.Visibility == Visibility.Visible);
        var iikoVisible = IikoActionsPanel.Children.OfType<Button>().Any(b => b.Visibility == Visibility.Visible);

        SystemCard.Visibility = systemVisible ? Visibility.Visible : Visibility.Collapsed;
        IikoCard.Visibility = iikoVisible ? Visibility.Visible : Visibility.Collapsed;
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

            var installerType = DetectInstallerType(tempPath);
            if (installerType == InstallerType.Html || installerType == InstallerType.Unknown)
            {
                ShowNotification("Скачан некорректный установщик iikoCard (ожидался .exe/.msi)", NotificationType.Error);
                _logger.Error("Обновление iikoCard: неизвестный формат файла, установка прервана");
                return;
            }

            var psi = BuildInstallerStartInfo(tempPath, installerType);

            var proc = Process.Start(psi);
            if (proc == null)
            {
                ShowNotification("Не удалось запустить установщик iikoCard", NotificationType.Error);
                _logger.Error("Обновление iikoCard: Process.Start вернул null");
                return;
            }

            ShowNotification("Установка iikoCard запущена", NotificationType.Success);
            _logger.Information("Обновление iikoCard: процесс установки PID {Pid}", proc.Id);

            var exitCode = await WaitForInstallerExitAsync(proc, tempPath);
            _logger.Information("Обновление iikoCard: установка завершена с кодом {Code}", exitCode);
            if (exitCode == 0)
            {
                ShowNotification("iikoCard установлена", NotificationType.Success);
            }
            else
            {
                ShowNotification($"Установка iikoCard завершилась с кодом {exitCode}", NotificationType.Warning);
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

    private enum InstallerType
    {
        Unknown,
        Exe,
        Msi,
        Html
    }

    private InstallerType DetectInstallerType(string path)
    {
        try
        {
            var bytes = File.ReadAllBytes(path);
            if (bytes.Length >= 4 && bytes[0] == 'M' && bytes[1] == 'Z')
            {
                // EXE
                return InstallerType.Exe;
            }

            if (bytes.Length >= 8 &&
                bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0)
            {
                // MSI (OLE structured storage)
                return InstallerType.Msi;
            }

            var textSample = System.Text.Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 500)).ToLowerInvariant();
            if (textSample.Contains("<html") || textSample.Contains("<!doctype html"))
            {
                return InstallerType.Html;
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось определить тип файла установщика");
        }

        return InstallerType.Unknown;
    }

    private ProcessStartInfo BuildInstallerStartInfo(string path, InstallerType type)
    {
        if (type == InstallerType.Msi)
        {
            return new ProcessStartInfo
            {
                FileName = "msiexec.exe",
                Arguments = $"/i \"{path}\" /quiet /norestart",
                UseShellExecute = true,
                Verb = "runas"
            };
        }

        // exe
        return new ProcessStartInfo
        {
            FileName = path,
            Arguments = "/S",
            UseShellExecute = true,
            Verb = "runas"
        };
    }

    private async Task<int> WaitForInstallerExitAsync(Process process, string installerPath)
    {
        var watch = Stopwatch.StartNew();
        var max = TimeSpan.FromMinutes(10);
        var name = Path.GetFileNameWithoutExtension(installerPath);
        var startTime = DateTime.Now;

        try
        {
            if (!process.HasExited)
            {
                await Task.Run(() => process.WaitForExit((int)max.TotalMilliseconds));
            }
        }
        catch
        {
            // ignore
        }

        if (process.HasExited && watch.Elapsed < TimeSpan.FromSeconds(3))
        {
            // Возможно, основной процесс разорвало, ищем дочерние с тем же именем.
            while (watch.Elapsed < max)
            {
                var others = Process.GetProcessesByName(name)
                    .Where(p => p.StartTime >= startTime && !p.HasExited)
                    .ToList();

                if (others.Count == 0)
                {
                    break;
                }

                foreach (var p in others)
                {
                    try
                    {
                        _logger.Information("Обновление iikoCard: ждём дочерний процесс PID {Pid}", p.Id);
                        await Task.Run(() => p.WaitForExit((int)max.Subtract(watch.Elapsed).TotalMilliseconds));
                        if (p.HasExited)
                        {
                            return p.ExitCode;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                await Task.Delay(1000);
            }
        }

        return process.HasExited ? process.ExitCode : -1;
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
                    ShowNotification("Отключаем защитник...", NotificationType.Warning);
                    var def = await _securityService.DisableDefenderAsync();
                    _logger.Information("Отключение защитника: {Result}. Вывод: {Output}", def.Ok ? "успех" : "ошибка", def.Output?.Trim());

                    ShowNotification("Отключаем брандмауэр...", NotificationType.Warning);
                    var fw = await _securityService.DisableFirewallAsync();
                    _logger.Information("Отключение брандмауэра: {Result}. Вывод: {Output}", fw.Ok ? "успех" : "ошибка", fw.Output?.Trim());

                    if (def.Ok && fw.Ok)
                    {
                        ShowNotification("Защитник и брандмауэр отключены", NotificationType.Success);
                    }
                    else
                    {
                        var detail = $"{(def.Ok ? "" : "Defender: " + def.Output)} {(fw.Ok ? "" : "Firewall: " + fw.Output)}".Trim();
                        ShowNotification(string.IsNullOrWhiteSpace(detail) ? "Часть действий не выполнена" : detail, NotificationType.Error);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Ошибка при отключении защитника/брандмауэра");
                    ShowNotification($"Ошибка при отключении защиты: {ex.Message}", NotificationType.Error);
                }
            });
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
}
