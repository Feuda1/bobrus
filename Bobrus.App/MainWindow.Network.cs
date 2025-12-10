using System;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace Bobrus.App;

public partial class MainWindow
{
    private Process? _networkProcess;

    private void OnNetworkDiagnosticsClicked(object sender, RoutedEventArgs e)
    {
        ShowNetworkOverlay(runIpconfig: true);
    }

    private void OnNetworkOverlayCloseClicked(object sender, RoutedEventArgs e)
    {
        HideNetworkOverlay();
    }

    private void OnOpenSettingsClicked(object sender, RoutedEventArgs e)
    {
        SettingsOverlay.Visibility = Visibility.Visible;
    }

    private void OnSettingsOverlayCloseClicked(object sender, RoutedEventArgs e)
    {
        SettingsOverlay.Visibility = Visibility.Collapsed;
    }

    private void OnHideToTrayToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsToggle) return;
        _hideToTrayEnabled = HideToTrayToggle.IsChecked == true;
        EnsureTrayIcon();
        ShowNotification(_hideToTrayEnabled ? "При закрытии приложение будет скрываться в трей" : "При закрытии приложение завершится", NotificationType.Info);
        SaveAppSettings();
    }

    private void OnAutostartToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsToggle) return;
        var enable = AutostartToggle.IsChecked == true;
        SetAutostart(enable);
        ShowNotification(enable ? "Автозапуск включён" : "Автозапуск выключен", NotificationType.Info);
        SaveAppSettings();
    }

    private void OnShowAllSectionsToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsToggle) return;
        _showAllSections = ShowAllSectionsToggle.IsChecked == true;
        ApplySectionLayout();
        ShowNotification(_showAllSections ? "Общий список включён" : "Разделы разделены по вкладкам", NotificationType.Info);
        SaveAppSettings();
    }

    private void OnConsoleToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsToggle) return;
        _consoleVisible = ConsoleToggle.IsChecked != false;
        ApplyConsoleVisibility();
        ShowNotification(_consoleVisible ? "Журнал включён" : "Журнал скрыт", NotificationType.Info);
        SaveAppSettings();
    }

    private void OnThemeToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsToggle) return;
        var theme = ThemeToggle.IsChecked == true ? ThemeVariant.Dark : ThemeVariant.Light;
        ApplyTheme(theme);
        ShowNotification(theme == ThemeVariant.Dark ? "Тёмная тема включена" : "Светлая тема включена", NotificationType.Info);
        SaveAppSettings();
    }

    private void OnNetworkRunPingClicked(object sender, RoutedEventArgs e)
    {
        ClearNetworkOutput();
        var target = (NetworkPingInput.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(target))
        {
            ShowNotification("Введите адрес для ping", NotificationType.Warning);
            FocusNetworkInput();
            return;
        }

        _ = RunNetworkCommandAsync("ping", $"-n 4 {target}", $"ping {target}", TimeSpan.FromSeconds(45), useOemEncoding: true);
    }

    private void OnNetworkClearOutputClicked(object sender, RoutedEventArgs e)
    {
        ClearNetworkOutput();
        FocusNetworkInput();
    }

    private async Task RunFullNetworkDiagnosticsAsync()
    {
        if (_networkProcess != null)
        {
            AppendNetworkOutput("Уже выполняется сетевой процесс. Дождитесь завершения.");
            return;
        }

        SetNetworkBusy(true);

        try
        {
            var server = NormalizeServer(TryGetDefaultServerFromConfig());
            if (string.IsNullOrWhiteSpace(server))
            {
                server = NormalizeServer(NetworkPingInput.Text ?? string.Empty);
            }
            if (string.IsNullOrWhiteSpace(server))
            {
                AppendNetworkOutput("Не удалось определить сервер для диагностики (config.xml и поле пустые).");
                return;
            }

            var domain = server.EndsWith(".iiko.it", StringComparison.OrdinalIgnoreCase)
                ? server
                : $"{server}.iiko.it";

            var tempDir = Path.Combine(Path.GetTempPath(), "bobrus_netdiag_run_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var tempScript = Path.Combine(tempDir, "netdiag.bat");
            var tempLog = Path.Combine(tempDir, "532_full_diagnostic.log");
            File.WriteAllText(tempScript, GetEmbeddedNetDiagScript(domain));

            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c \"\"{tempScript}\"\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = tempDir
            };
            var consoleEncoding = GetOemConsoleEncoding();
            psi.StandardOutputEncoding = consoleEncoding;
            psi.StandardErrorEncoding = consoleEncoding;

            using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
            _networkProcess = process;
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(10));
            var sb = new StringBuilder();

            process.OutputDataReceived += (_, ev) =>
            {
                if (ev.Data == null) return;
                sb.AppendLine(ev.Data);
                if (ev.Data.StartsWith("[STEP", StringComparison.OrdinalIgnoreCase))
                {
                    Dispatcher.Invoke(() => AppendNetworkOutput(ev.Data));
                }
            };
            process.ErrorDataReceived += (_, ev) =>
            {
                if (ev.Data == null) return;
                sb.AppendLine("[ERR] " + ev.Data);
            };

            if (!process.Start())
            {
                AppendNetworkOutput("Не удалось запустить диагностику сети");
                return;
            }

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
                process.WaitForExit();
                sb.AppendLine("Диагностика сети прервана по таймауту");
                AppendNetworkOutput("Диагностика сети прервана по таймауту");
            }

            var targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BobrusNetDiag");
            Directory.CreateDirectory(targetDir);
            var targetLog = Path.Combine(targetDir, $"netdiag_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            if (File.Exists(tempLog))
            {
                File.Copy(tempLog, targetLog, overwrite: true);
            }
            else
            {
                File.WriteAllText(targetLog, sb.ToString());
            }

            AppendNetworkOutput($"> Диагностика сети завершена, лог: {targetLog}");
            ShowNotification($"Лог диагностики сети сохранён: {targetLog}", NotificationType.Success);
            _logger.Information("Сетевая диагностика из раздела 'Сеть' завершена. Лог: {Log}", targetLog);

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{targetLog}\"",
                    UseShellExecute = true
                });
            }
            catch { }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Ошибка при выполнении диагностики сети (кнопка Диагностика сети)");
            ShowNotification($"Ошибка диагностики сети: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            _networkProcess?.Dispose();
            _networkProcess = null;
            SetNetworkBusy(false);
        }
    }

    private void OnNetworkRunIpconfigClicked(object sender, RoutedEventArgs e)
    {
        ClearNetworkOutput();
        _ = RunNetworkCommandAsync("ipconfig", "/all", "ipconfig /all", TimeSpan.FromSeconds(30), useOemEncoding: true);
    }

    private async void OnNetworkRunFullDiagClicked(object sender, RoutedEventArgs e)
    {
        await RunFullNetworkDiagnosticsAsync();
    }

    private void ShowNetworkOverlay(bool runIpconfig)
    {
        NetworkOverlay.Visibility = Visibility.Visible;
        FocusNetworkInput();
        if (runIpconfig)
        {
            ClearNetworkOutput();
            _ = RunNetworkCommandAsync("ipconfig", "/all", "ipconfig /all", TimeSpan.FromSeconds(30), useOemEncoding: true);
        }
    }

    private void HideNetworkOverlay()
    {
        NetworkOverlay.Visibility = Visibility.Collapsed;
        TryKillNetworkProcess();
        TryDisposeNetworkProcess();
        SetNetworkBusy(false);
    }

    private async Task RunNetworkCommandAsync(string fileName, string arguments, string displayName, TimeSpan timeout, bool useOemEncoding = false)
    {
        if (_networkProcess != null)
        {
            AppendNetworkOutput("Уже выполняется другая команда, дождитесь завершения.");
            return;
        }

        SetNetworkBusy(true);
        try
        {
            AppendNetworkOutput($"> {displayName}");

            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            if (useOemEncoding)
            {
                var consoleEncoding = GetOemConsoleEncoding();
                psi.StandardOutputEncoding = consoleEncoding;
                psi.StandardErrorEncoding = consoleEncoding;
            }

            var process = new Process
            {
                StartInfo = psi,
                EnableRaisingEvents = true
            };

            var tcs = new TaskCompletionSource<int>();
            process.OutputDataReceived += (_, e) => { if (e.Data != null) AppendNetworkOutput(e.Data); };
            process.ErrorDataReceived += (_, e) => { if (e.Data != null) AppendNetworkOutput(e.Data); };
            process.Exited += (_, _) => tcs.TrySetResult(process.ExitCode);

            _networkProcess = process;

            if (!process.Start())
            {
                AppendNetworkOutput("Не удалось запустить процесс.");
                TryDisposeNetworkProcess();
                return;
            }

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            var finishedTask = await Task.WhenAny(tcs.Task, Task.Delay(timeout));
            if (finishedTask != tcs.Task)
            {
                AppendNetworkOutput("Превышен таймаут, завершаем процесс...");
                TryKillNetworkProcess();
                await tcs.Task.ContinueWith(_ => Task.CompletedTask);
            }

            await Task.Run(() => process.WaitForExit());

            var exitCode = _networkProcess?.ExitCode ?? -1;
            AppendNetworkOutput($"Готово (код {exitCode}).");
        }
        catch (Exception ex)
        {
            AppendNetworkOutput($"Ошибка: {ex.Message}");
        }
        finally
        {
            TryDisposeNetworkProcess();
            SetNetworkBusy(false);
        }
    }

    private void AppendNetworkOutput(string line)
    {
        Dispatcher.Invoke(() =>
        {
            NetworkOutputBox.AppendText(line + Environment.NewLine);
            NetworkOutputBox.ScrollToEnd();
        });
    }

    private void ClearNetworkOutput()
    {
        NetworkOutputBox.Clear();
        NetworkOutputBox.ScrollToHome();
    }

    private void SetNetworkBusy(bool busy)
    {
        Dispatcher.Invoke(() =>
        {
            NetworkRunPingButton.IsEnabled = !busy;
            NetworkOverlayCloseButton.IsEnabled = true;
            if (NetworkRunFullDiagButton != null)
                NetworkRunFullDiagButton.IsEnabled = !busy;
            if (NetworkRunIpconfigButton != null)
                NetworkRunIpconfigButton.IsEnabled = !busy;
        });
    }

    private void FocusNetworkInput()
    {
        Dispatcher.InvokeAsync(() =>
        {
            NetworkPingInput.Focus();
            NetworkPingInput.CaretIndex = NetworkPingInput.Text.Length;
        }, DispatcherPriority.Background);
    }

    private void TryKillNetworkProcess()
    {
        try
        {
            if (_networkProcess != null && !_networkProcess.HasExited)
            {
                _networkProcess.Kill(entireProcessTree: true);
            }
        }
        catch
        {
        }
    }

    private void TryDisposeNetworkProcess()
    {
        try
        {
            _networkProcess?.Dispose();
        }
        catch
        {
        }
        finally
        {
            _networkProcess = null;
        }
    }

    private static Encoding GetOemConsoleEncoding()
    {
        try
        {
            var codePage = CultureInfo.CurrentCulture.TextInfo.OEMCodePage;
            return Encoding.GetEncoding(codePage);
        }
        catch
        {
            try { return Encoding.GetEncoding(866); } catch { }
            try { return Encoding.GetEncoding(1251); } catch { }
            return Encoding.UTF8;
        }
    }
}
