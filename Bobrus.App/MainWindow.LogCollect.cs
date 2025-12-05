using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Bobrus.App;

public partial class MainWindow
{
    private List<DateTime> _availableLogDates = new();
    private bool _suppressListSelection;
    private bool _handlingEntityOptions;

    private record struct EntitiesSelection(
        bool CollectAll,
        bool Front,
        bool Events,
        bool Plugins,
        bool Search,
        bool Sync);

    private record struct NetDiagOptions(
        bool Enabled,
        string Server);

    private static readonly string[] NetDiagStepLabels =
    {
        "subinterfaces",
        "route print",
        "ipconfig /all",
        "ping -n 10",
        "ping -f -l 1472",
        "ping -f -l 1462",
        "ping -f -l 1452",
        "ping -f -l 1372",
        "ping -l 65500",
        "tracert",
        "pathping"
    };

    private CancellationTokenSource? _collectCts;
    private bool _isCollecting;

    private void OnCollectLogsClicked(object sender, RoutedEventArgs e)
    {
        LoadAvailableLogDates();
        CollectIncludeCash.IsChecked = false;
        CollectIncludeEntities.IsChecked = false;
        CollectIncludeMsinfo.IsChecked = false;
        CollectIncludeNetDiag.IsChecked = false;
        ResetEntitiesOptions();
        ResetNetDiagOptions();
        UpdateStartItems(string.Empty);
        UpdateEndItems(string.Empty);
        if (_availableLogDates.Count > 0)
        {
            var earliest = _availableLogDates.Last();
            var latest = _availableLogDates.First();
            SetStartSelection(earliest);
            SetEndSelection(latest);
        }
        ClampEndDate(forceToLatest: true);
        HideAllDropdowns();
        LogCollectOverlay.Visibility = Visibility.Visible;
    }

    private void OnLogCollectCancelClicked(object sender, RoutedEventArgs e)
    {
        if (_isCollecting && _collectCts != null && !_collectCts.IsCancellationRequested)
        {
            _logger.Information("Отмена сбора логов по запросу пользователя");
            CollectCancelButton.IsEnabled = false;
            CollectCancelButton.Content = "Отмена...";
            _collectCts.Cancel();
            return;
        }

        HideAllDropdowns();
        LogCollectOverlay.Visibility = Visibility.Collapsed;
    }

    private void OnLogCollectOverlayCloseClicked(object sender, RoutedEventArgs e)
    {
        // Закрываем окно без отмены процесса
        HideAllDropdowns();
        LogCollectOverlay.Visibility = Visibility.Collapsed;
    }

    private async void OnLogCollectAcceptClicked(object sender, RoutedEventArgs e)
    {
        var start = ParseDateFromInput(CollectStartInput, DateTime.Today);
        var end = ParseDateFromInput(CollectEndInput, DateTime.Today);
        if (end < start) end = start;
        var includeCash = CollectIncludeCash.IsChecked == true;
        var includeMsinfo = CollectIncludeMsinfo.IsChecked == true;
        var netDiagOptions = BuildNetDiagOptions();
        var entitySelection = BuildEntitiesSelection();
        var totalStages = 1 + (includeMsinfo ? 1 : 0) + (netDiagOptions.Enabled ? NetDiagStepLabels.Length : 0);

        if (_isCollecting)
        {
            _logger.Warning("Сбор логов уже выполняется, повторный запуск отклонён");
            return;
        }

        try
        {
            _isCollecting = true;
            _collectCts = new CancellationTokenSource();
            CollectAcceptButton.IsEnabled = false;
            CollectCancelButton.IsEnabled = true;
            CollectCancelButton.Content = "Отмена";

            var zipName = $"Logs_{DateTime.Now:ddMMyyyy_HHmmss}.zip";
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var targetDir = Path.Combine(documents, "BobrusLogs");
            Directory.CreateDirectory(targetDir);
            var zipPath = Path.Combine(targetDir, zipName);

            UpdateCollectProgress("Сбор логов...", 0, totalStages);
            _logger.Information("Сбор логов запущен. Cash: {Cash}, Entities: {Entities}, Msinfo: {Msinfo}, NetDiag: {NetDiag}", includeCash, HasEntitiesSelection(entitySelection), includeMsinfo, netDiagOptions.Enabled);

            await Task.Run(() => CreateLogsArchive(zipPath, includeCash, includeMsinfo, netDiagOptions, entitySelection, start, end, totalStages, _collectCts!.Token), _collectCts.Token);

            _collectCts.Token.ThrowIfCancellationRequested();

            ShowNotification($"Архив логов готов: {zipPath}", NotificationType.Success);
            _logger.Information("Сбор логов завершён: {Path}", zipPath);

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{zipPath}\"",
                    UseShellExecute = true
                });
            }
            catch
            {
                // ignore
            }
        }
        catch (OperationCanceledException)
        {
            _logger.Information("Сбор логов отменён");
            ShowNotification("Сбор логов отменён", NotificationType.Warning);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка сбора логов");
            ShowNotification($"Ошибка сбора логов: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            UpdateCollectProgress(string.Empty, 0, 0, hideOnly: true);
            CollectAcceptButton.IsEnabled = true;
            CollectCancelButton.IsEnabled = true;
            CollectCancelButton.Content = "Отмена";
            _collectCts?.Dispose();
            _collectCts = null;
            _isCollecting = false;
            HideAllDropdowns();
            LogCollectOverlay.Visibility = Visibility.Collapsed;
        }
    }

    private void OnCollectCashToggleChanged(object sender, RoutedEventArgs e)
    {
        var cash = CollectIncludeCash.IsChecked == true;
        if (cash)
        {
            CollectIncludeEntities.IsChecked = false;
            CollectIncludeEntities.IsEnabled = false;
            CollectIncludeEntities.Opacity = 0.6;
            ResetEntitiesOptions();
        }
        else
        {
            CollectIncludeEntities.IsEnabled = true;
            CollectIncludeEntities.Opacity = 1.0;
        }
    }

    private void CreateLogsArchive(string zipPath, bool includeCashServer, bool includeMsinfo, NetDiagOptions netDiag, EntitiesSelection entitySelection, DateTime startDate, DateTime endDate, int totalStages, CancellationToken token)
    {
        var logsDir = Path.Combine(_cashServerBase, "Logs");
        var cashRoot = _cashServerBase;
        var entitiesDir = Path.Combine(_cashServerBase, "EntitiesStorage");
        var start = startDate.Date;
        var end = endDate.Date.AddDays(1).AddTicks(-1);
        var currentStage = 0;

        token.ThrowIfCancellationRequested();

        if (File.Exists(zipPath))
        {
            File.Delete(zipPath);
        }

        using var archive = ZipFile.Open(zipPath, ZipArchiveMode.Create);

        currentStage++;
        UpdateCollectProgress("Сбор файлов логов", currentStage, totalStages);
        token.ThrowIfCancellationRequested();
        // *.log и архивы логов из Logs
        if (Directory.Exists(logsDir))
        {
            var logFiles = Directory.EnumerateFiles(logsDir, "*.log", SearchOption.TopDirectoryOnly)
                .Where(f =>
                {
                    var ts = File.GetLastWriteTime(f);
                    return ts >= start && ts <= end;
                });
            foreach (var file in logFiles)
            {
                var entryRoot = includeCashServer ? Path.Combine("CashServer", "Logs") : "Logs";
                AddFileToArchive(archive, file, Path.Combine(entryRoot, Path.GetFileName(file)));
            }

            var archiveFiles = Directory.EnumerateFiles(logsDir, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".7z", StringComparison.OrdinalIgnoreCase))
                .Where(f =>
                {
                    var ts = File.GetLastWriteTime(f);
                    return ts >= start && ts <= end;
                });
            foreach (var file in archiveFiles)
            {
                var entryRoot = includeCashServer ? Path.Combine("CashServer", "Logs") : "Logs";
                AddFileToArchive(archive, file, Path.Combine(entryRoot, Path.GetFileName(file)));
            }
        }

        // *.xml из корня CashServer (пропускаем, если забрали весь CashServer)
        if (!includeCashServer && Directory.Exists(cashRoot))
        {
            var xmlFiles = Directory.EnumerateFiles(cashRoot, "*.xml", SearchOption.TopDirectoryOnly);
            foreach (var file in xmlFiles)
            {
                AddFileToArchive(archive, file, Path.Combine("Config", Path.GetFileName(file)));
            }
        }

        // CashServer (опционально) — кладём всё, кроме Logs (логи уже добавили отфильтрованными)
        if (includeCashServer && Directory.Exists(cashRoot))
        {
            AddDirectoryToArchive(archive, cashRoot, "CashServer", relative =>
            {
                // пропускаем папку Logs
                var firstSegment = relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).FirstOrDefault();
                return !string.Equals(firstSegment, "Logs", StringComparison.OrdinalIgnoreCase);
            });
        }

        // EntitiesStorage (опционально)
        AddEntitiesStorage(archive, entitiesDir, entitySelection);
        token.ThrowIfCancellationRequested();

        if (includeMsinfo)
        {
            currentStage++;
            UpdateCollectProgress("Сбор msinfo32...", currentStage, totalStages);
            AddMsinfoReport(archive);
        }

        if (netDiag.Enabled)
        {
            AddNetworkDiagnostics(archive, netDiag, ref currentStage, totalStages, token);
        }
    }

    private void AddDirectoryToArchive(ZipArchive archive, string dirPath, string entryRoot, Func<string, bool>? fileFilter = null)
    {
        IEnumerable<string>? files;
        try
        {
            files = Directory.EnumerateFiles(dirPath, "*.*", SearchOption.AllDirectories);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось перечислить файлы каталога {Dir}", dirPath);
            return;
        }

        foreach (var file in files)
        {
            var relative = Path.GetRelativePath(dirPath, file);
            if (fileFilter != null && !fileFilter(relative))
            {
                continue;
            }
            var entryName = Path.Combine(entryRoot, relative).Replace('\\', '/');
            AddFileToArchive(archive, file, entryName);
        }
    }

    private void AddFileToArchive(ZipArchive archive, string filePath, string entryName)
    {
        try
        {
            archive.CreateEntryFromFile(filePath, entryName);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось добавить файл {File} в архив", filePath);
        }
    }

    private void AddEntitiesStorage(ZipArchive archive, string entitiesDir, EntitiesSelection selection)
    {
        if (!Directory.Exists(entitiesDir))
        {
            return;
        }

        if (!HasEntitiesSelection(selection))
        {
            return;
        }

        if (selection.CollectAll)
        {
            AddDirectoryToArchive(archive, entitiesDir, "EntitiesStorage");
            return;
        }

        void AddIf(bool flag, string subfolder)
        {
            if (!flag) return;
            var dir = Path.Combine(entitiesDir, subfolder);
            try
            {
                if (Directory.Exists(dir))
                {
                    AddDirectoryToArchive(archive, dir, Path.Combine("EntitiesStorage", subfolder));
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Не удалось добавить раздел {Section} из {Dir}", subfolder, dir);
            }
        }

        AddIf(selection.Front, "Entities");
        AddIf(selection.Events, "Events");
        AddIf(selection.Plugins, "Plugins");
        AddIf(selection.Search, "Search");
        AddIf(selection.Sync, "Synchronization");
    }

    private NetDiagOptions BuildNetDiagOptions()
    {
        var enabled = CollectIncludeNetDiag.IsChecked == true;
        var server = (CollectNetDiagServerInput.Text ?? string.Empty).Trim();
        return new NetDiagOptions(enabled, server);
    }

    private void ResetNetDiagOptions()
    {
        CollectNetDiagOptions.Visibility = Visibility.Collapsed;
        CollectNetDiagServerInput.Text = TryGetDefaultServerFromConfig();
    }

    private void OnCollectNetDiagToggleChanged(object sender, RoutedEventArgs e)
    {
        var enabled = CollectIncludeNetDiag.IsChecked == true;
        CollectNetDiagOptions.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        if (enabled && string.IsNullOrWhiteSpace(CollectNetDiagServerInput.Text))
        {
            CollectNetDiagServerInput.Text = TryGetDefaultServerFromConfig();
        }
    }

    private string TryGetDefaultServerFromConfig()
    {
        try
        {
            var configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "iiko", "CashServer", "config.xml");
            if (!File.Exists(configPath))
            {
                return string.Empty;
            }

            var xml = File.ReadAllText(configPath);
            var marker = "<serverUrl>";
            var endMarker = "</serverUrl>";
            var startIdx = xml.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (startIdx < 0) return string.Empty;
            startIdx += marker.Length;
            var endIdx = xml.IndexOf(endMarker, startIdx, StringComparison.OrdinalIgnoreCase);
            if (endIdx < 0) return string.Empty;
            var url = xml.Substring(startIdx, endIdx - startIdx).Trim();
            return NormalizeServer(url);
        }
        catch
        {
            return string.Empty;
        }
    }

    private string NormalizeServer(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var v = value.Trim();
        if (v.StartsWith("http://", StringComparison.OrdinalIgnoreCase)) v = v.Substring(7);
        if (v.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) v = v.Substring(8);
        var slashIdx = v.IndexOf('/');
        if (slashIdx >= 0)
        {
            v = v[..slashIdx];
        }
        var colonIdx = v.IndexOf(':');
        if (colonIdx >= 0)
        {
            v = v[..colonIdx];
        }
        if (v.EndsWith(".iiko.it", StringComparison.OrdinalIgnoreCase))
        {
            v = v[..^".iiko.it".Length];
        }
        return v;
    }

    private void AddNetworkDiagnostics(ZipArchive archive, NetDiagOptions options, ref int currentStage, int totalStages, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var server = NormalizeServer(options.Server);
        if (string.IsNullOrWhiteSpace(server))
        {
            server = TryGetDefaultServerFromConfig();
        }

        if (string.IsNullOrWhiteSpace(server))
        {
            _logger.Warning("Сетевой сервер для диагностики не определён");
            return;
        }

        var domain = server.EndsWith(".iiko.it", StringComparison.OrdinalIgnoreCase)
            ? server
            : $"{server}.iiko.it";

        _logger.Information("Сетевая диагностика: старт для {Server}", domain);
        var tempDir = Path.Combine(Path.GetTempPath(), "bobrus_netdiag_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var tempScript = Path.Combine(tempDir, "netdiag.bat");
        var logPath = Path.Combine(tempDir, "532_full_diagnostic.log");

        try
        {
            var scriptContent = GetEmbeddedNetDiagScript(domain);
            File.WriteAllText(tempScript, scriptContent);

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

            var diagStepsCompleted = 0;
            using var process = Process.Start(psi);
            if (process == null)
            {
                _logger.Warning("Не удалось запустить сетевую диагностику");
                return;
            }

            var stageRef = currentStage;
            var maxWait = TimeSpan.FromMinutes(10);
            using var reg = token.Register(() =>
            {
                try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
            });
            process.OutputDataReceived += (_, e) =>
            {
                if (e.Data == null) return;
                if (e.Data.StartsWith("[STEP", StringComparison.OrdinalIgnoreCase))
                {
                    var msg = e.Data.Trim();
                    diagStepsCompleted = Math.Min(diagStepsCompleted + 1, NetDiagStepLabels.Length);
                    stageRef = Math.Min(stageRef + 1, totalStages);
                    UpdateCollectProgress($"Сетевая диагностика: {msg}", stageRef, totalStages);
                    _logger.Information("Сетевая диагностика: {Msg}", msg);
                }
            };
            process.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    _logger.Warning("NetDiag stderr: {Msg}", e.Data);
                }
            };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            var started = DateTime.UtcNow;
            while (!process.HasExited)
            {
                if (token.WaitHandle.WaitOne(500))
                {
                    try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
                    process.WaitForExit(2000);
                    if (File.Exists(logPath))
                    {
                        try { File.AppendAllLines(logPath, new[] { "Диагностика отменена пользователем" }); } catch { }
                    }
                    throw new OperationCanceledException(token);
                }

                if (DateTime.UtcNow - started > maxWait)
                {
                    try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
                    process.WaitForExit(2000);
                    if (File.Exists(logPath))
                    {
                        try { File.AppendAllLines(logPath, new[] { "Превышен таймаут выполнения скрипта" }); } catch { }
                    }
                    _logger.Warning("Сетевая диагностика: таймаут (>{Minutes} мин)", maxWait.TotalMinutes);
                    break;
                }
            }
            if (process.HasExited)
            {
                if (File.Exists(logPath))
                {
                    try { File.AppendAllLines(logPath, new[] { $"Завершено с кодом {process.ExitCode}" }); } catch { }
                }
            }

            if (!File.Exists(logPath))
            {
                _logger.Warning("Лог диагностики не создан: {Path}", logPath);
                return;
            }

            AddFileToArchive(archive, logPath, Path.Combine("NetworkDiagnostics", "532_full_diagnostic.log"));
            _logger.Information("Сетевая диагностика завершена, лог добавлен");
            currentStage = stageRef;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось выполнить сетевую диагностику");
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
        }
    }

    private string GetEmbeddedNetDiagScript(string domain)
    {
        var lines = new[]
        {
            @"@echo off",
            @"setlocal enableextensions",
            @"set SERVER=" + domain,
            @"set LOG=532_full_diagnostic.log",
            @"echo Старт диагностики для %SERVER% > %LOG%",
            @"echo [STEP 1/11] subinterfaces",
            @"netsh interface ipv4 show subinterfaces >> %LOG%",
            @"echo [STEP 2/11] route print",
            @"route print >> %LOG%",
            @"echo [STEP 3/11] ipconfig /all",
            @"ipconfig /all >> %LOG%",
            @"echo [STEP 4/11] ping -n 10",
            @"ping %SERVER% -n 10 >> %LOG%",
            @"echo [STEP 5/11] ping -f -l 1472",
            @"ping %SERVER% -f -l 1472 >> %LOG%",
            @"echo [STEP 6/11] ping -f -l 1462",
            @"ping %SERVER% -f -l 1462 >> %LOG%",
            @"echo [STEP 7/11] ping -f -l 1452",
            @"ping %SERVER% -f -l 1452 >> %LOG%",
            @"echo [STEP 8/11] ping -f -l 1372",
            @"ping %SERVER% -f -l 1372 >> %LOG%",
            @"echo [STEP 9/11] ping -l 65500",
            @"ping %SERVER% -l 65500 >> %LOG%",
            @"echo [STEP 10/11] tracert",
            @"tracert %SERVER% >> %LOG%",
            @"echo [STEP 11/11] pathping (может занять до 10 минут)",
            @"pathping %SERVER% >> %LOG%",
            @"echo [STEP DONE] done",
            @"endlocal"
        };

        var template = string.Join(Environment.NewLine, lines);
        return template.Replace("xxx.iiko.it", domain, StringComparison.OrdinalIgnoreCase);
    }
    private void AddMsinfoReport(ZipArchive archive)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"msinfo_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.txt");
        var fallbackPath = Path.Combine(Path.GetTempPath(), $"systeminfo_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.txt");
        var entryName = "msinfo32_report.txt";
        var fallbackEntryName = "systeminfo_fallback.txt";
        try
        {
            if (TryRunMsinfo(tempPath, out var reason))
            {
                AddFileToArchive(archive, tempPath, entryName);
                return;
            }

            _logger.Warning("Не удалось получить сведения msinfo32: {Reason}", reason);

            if (TryRunSystemInfo(fallbackPath, out var fallbackReason))
            {
                AddFileToArchive(archive, fallbackPath, fallbackEntryName);
                return;
            }

            _logger.Warning("Не удалось собрать сведения через systeminfo: {Reason}", fallbackReason);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось получить сведения системы через msinfo32");
        }
        finally
        {
            TryDeleteFile(tempPath);
            TryDeleteFile(fallbackPath);
        }
    }

    private bool TryRunMsinfo(string outputPath, out string reason)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "msinfo32.exe",
                Arguments = $"/report \"{outputPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                reason = "msinfo32.exe не запустилась";
                return false;
            }

            if (!process.WaitForExit(120000))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // ignore
                }

                reason = "msinfo32.exe превысила таймаут";
                return false;
            }

            if (process.ExitCode != 0)
            {
                reason = $"msinfo32.exe завершилась с кодом {process.ExitCode}";
                return false;
            }

            if (!File.Exists(outputPath))
            {
                reason = "msinfo32.exe не создала файл отчёта";
                return false;
            }

            reason = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            return false;
        }
    }

    private bool TryRunSystemInfo(string outputPath, out string reason)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c systeminfo > \"{outputPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                reason = "systeminfo не запустилась";
                return false;
            }

            if (!process.WaitForExit(60000))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // ignore
                }

                reason = "systeminfo превысила таймаут";
                return false;
            }

            if (process.ExitCode != 0)
            {
                reason = $"systeminfo завершилась с кодом {process.ExitCode}";
                return false;
            }

            if (!File.Exists(outputPath))
            {
                reason = "systeminfo не создала файл отчёта";
                return false;
            }

            reason = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            return false;
        }
    }

    private void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // ignore
        }
    }

    private void LoadAvailableLogDates()
    {
        var logsDir = Path.Combine(_cashServerBase, "Logs");
        var dates = new HashSet<DateTime>();

        if (Directory.Exists(logsDir))
        {
            var files = Directory.EnumerateFiles(logsDir, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".log", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".7z", StringComparison.OrdinalIgnoreCase));
            foreach (var file in files)
            {
                var dt = File.GetLastWriteTime(file).Date;
                dates.Add(dt);
            }
        }

        if (dates.Count == 0)
        {
            dates.Add(DateTime.Today);
        }

        var ordered = dates.OrderByDescending(d => d).ToList();
        _availableLogDates = ordered;
    }

    private DateTime ParseDateFromInput(TextBox input, DateTime fallback)
    {
        return ParseDate(input.Text, fallback);
    }

    private void ClampEndDate(bool forceToLatest = false)
    {
        var start = ParseDateFromInput(CollectStartInput, DateTime.Today);
        var end = ParseDateFromInput(CollectEndInput, DateTime.Today);
        var available = CollectEndList.Items
            .OfType<string>()
            .Select(s => DateTime.TryParse(s, out var d) ? d : (DateTime?)null)
            .Where(d => d.HasValue && d.Value >= start)
            .Select(d => d!.Value)
            .OrderByDescending(d => d)
            .ToList();

        if (available.Count == 0)
        {
            CollectEndInput.Text = start.ToShortDateString();
            return;
        }

        var target = end;
        if (forceToLatest || end < start || !available.Contains(end))
        {
            target = available.First();
        }

        SetEndSelection(target);
    }

    private string GetSearchText(TextBox input)
    {
        var text = input.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(text)) return string.Empty;

        var formattedMatches = _availableLogDates
            .Select(FormatDate)
            .Any(d => string.Equals(d, text, StringComparison.OrdinalIgnoreCase));

        return formattedMatches ? string.Empty : text;
    }

    private EntitiesSelection BuildEntitiesSelection()
    {
        if (CollectIncludeEntities.IsChecked != true)
        {
            return new EntitiesSelection(false, false, false, false, false, false);
        }

        var all = CollectEntitiesAll.IsChecked == true;
        var front = CollectEntitiesFront.IsChecked == true;
        var events = CollectEntitiesEvents.IsChecked == true;
        var plugins = CollectEntitiesPlugins.IsChecked == true;
        var search = CollectEntitiesSearch.IsChecked == true;
        var sync = CollectEntitiesSync.IsChecked == true;

        return new EntitiesSelection(all, front, events, plugins, search, sync);
    }

    private bool HasEntitiesSelection(EntitiesSelection selection)
    {
        return selection.CollectAll || selection.Front || selection.Events || selection.Plugins || selection.Search || selection.Sync;
    }

    private void UpdateStartItems(string search)
    {
        var current = ParseDateFromInput(CollectStartInput, DateTime.MinValue);
        CollectStartList.Items.Clear();
        var filtered = _availableLogDates
            .Where(d => string.IsNullOrWhiteSpace(search) || FormatDate(d).Contains(search, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(d => d);

        _suppressListSelection = true;
        foreach (var d in filtered)
        {
            CollectStartList.Items.Add(FormatDate(d));
        }

        SetStartSelection(current);
        _suppressListSelection = false;
    }

    private void UpdateEndItems(string search)
    {
        var start = ParseDateFromInput(CollectStartInput, DateTime.MinValue);
        var current = ParseDateFromInput(CollectEndInput, DateTime.MinValue);
        CollectEndList.Items.Clear();

        var filtered = _availableLogDates
            .Where(d => d >= start)
            .Where(d => string.IsNullOrWhiteSpace(search) || FormatDate(d).Contains(search, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(d => d);

        _suppressListSelection = true;
        foreach (var d in filtered)
        {
            CollectEndList.Items.Add(FormatDate(d));
        }

        SetEndSelection(current);
        _suppressListSelection = false;
    }

    private void OnCollectStartToggleClicked(object sender, RoutedEventArgs e)
    {
        var wasOpen = CollectStartPopup.IsOpen;
        CollectStartSearch.Text = string.Empty;
        UpdateStartItems(string.Empty);
        TogglePopup(CollectStartPopup, CollectEndPopup);
        if (!CollectStartPopup.IsOpen)
            return;
        if (!wasOpen)
            FocusSearchBox(CollectStartSearch);
    }

    private void OnCollectEndToggleClicked(object sender, RoutedEventArgs e)
    {
        var wasOpen = CollectEndPopup.IsOpen;
        CollectEndSearch.Text = string.Empty;
        UpdateEndItems(string.Empty);
        TogglePopup(CollectEndPopup, CollectStartPopup);
        if (!CollectEndPopup.IsOpen)
            return;
        if (!wasOpen)
            FocusSearchBox(CollectEndSearch);
    }

    private void TogglePopup(Popup target, Popup other)
    {
        var newState = !target.IsOpen;
        target.IsOpen = newState;
        other.IsOpen = false;
    }

    private void HideAllDropdowns()
    {
        CollectStartPopup.IsOpen = false;
        CollectEndPopup.IsOpen = false;
    }

    private void ShowCollectProgress(string text)
    {
        Dispatcher.Invoke(() =>
        {
            CollectProgressPanel.Visibility = Visibility.Visible;
            CollectProgressText.Text = text;
            GlobalProgressPanel.Visibility = Visibility.Visible;
            GlobalProgressText.Text = text;
        });
    }

    private void HideCollectProgress()
    {
        Dispatcher.Invoke(() =>
        {
            CollectProgressPanel.Visibility = Visibility.Collapsed;
            CollectProgressText.Text = string.Empty;
            GlobalProgressPanel.Visibility = Visibility.Collapsed;
            GlobalProgressText.Text = string.Empty;
            GlobalProgressBar.Value = 0;
            CollectProgressBar.Value = 0;
        });
    }

    private void UpdateCollectProgress(string text, int currentStage, int totalStages, bool hideOnly = false)
    {
        if (hideOnly)
        {
            HideCollectProgress();
            return;
        }

        var percent = totalStages <= 0 ? 0 : Math.Clamp((int)Math.Round((double)currentStage / Math.Max(totalStages, 1) * 100), 0, 100);
        Dispatcher.Invoke(() =>
        {
            CollectProgressPanel.Visibility = Visibility.Visible;
            CollectProgressText.Text = text;
            CollectProgressBar.Value = percent;

            GlobalProgressPanel.Visibility = Visibility.Visible;
            GlobalProgressText.Text = text;
            GlobalProgressBar.Value = percent;
        });
    }

    private void OnCollectStartListSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressListSelection) return;
        if (CollectStartList.SelectedItem is string s && DateTime.TryParse(s, out var dt))
        {
            SetInputText(CollectStartInput, dt.ToShortDateString());
            HideAllDropdowns();
            UpdateEndItems(GetSearchText(CollectEndInput));
            ClampEndDate(forceToLatest: true);
        }
    }

    private void OnCollectEndListSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressListSelection) return;
        if (CollectEndList.SelectedItem is string s && DateTime.TryParse(s, out var dt))
        {
            SetInputText(CollectEndInput, dt.ToShortDateString());
            HideAllDropdowns();
            ClampEndDate();
        }
    }

    private void OnCollectEntitiesToggleChanged(object sender, RoutedEventArgs e)
    {
        var enabled = CollectIncludeEntities.IsChecked == true;
        CollectEntitiesOptions.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;

        if (!enabled)
        {
            CollectEntitiesAll.IsChecked = false;
            DisableEntityOptions();
            return;
        }

        if (CollectEntitiesAll.IsChecked != true && !EntityOptionChecks().Any(cb => cb.IsChecked == true))
        {
            CollectEntitiesAll.IsChecked = true;
        }

        ApplyEntitiesAllState();
    }

    private void OnCollectEntitiesAllChanged(object sender, RoutedEventArgs e)
    {
        ApplyEntitiesAllState();
    }

    private void OnCollectEntitiesOptionChanged(object sender, RoutedEventArgs e)
    {
        if (_handlingEntityOptions) return;
        _handlingEntityOptions = true;
        CollectEntitiesAll.IsChecked = false;
        EnsureAtLeastOneEntityOption();
        _handlingEntityOptions = false;
    }

    private void ApplyEntitiesAllState()
    {
        var all = CollectEntitiesAll.IsChecked == true;
        _handlingEntityOptions = true;
        if (all)
        {
            EnableEntityOptions(setChecked: true);
        }
        else
        {
            EnableEntityOptions(setChecked: null);
            EnsureAtLeastOneEntityOption();
        }
        _handlingEntityOptions = false;
    }

    private void EnsureAtLeastOneEntityOption()
    {
        if (!EntityOptionChecks().Any(cb => cb.IsChecked == true))
        {
            CollectEntitiesFront.IsChecked = true;
        }
    }

    private void EnableEntityOptions(bool? setChecked = null)
    {
        foreach (var cb in EntityOptionChecks())
        {
            cb.IsEnabled = CollectIncludeEntities.IsChecked == true;
            if (setChecked.HasValue)
            {
                cb.IsChecked = setChecked.Value;
            }
        }
    }

    private void DisableEntityOptions()
    {
        foreach (var cb in EntityOptionChecks())
        {
            cb.IsEnabled = false;
            cb.IsChecked = false;
        }
    }

    private IEnumerable<CheckBox> EntityOptionChecks()
    {
        yield return CollectEntitiesFront;
        yield return CollectEntitiesEvents;
        yield return CollectEntitiesPlugins;
        yield return CollectEntitiesSearch;
        yield return CollectEntitiesSync;
    }

    private void ResetEntitiesOptions()
    {
        CollectEntitiesOptions.Visibility = Visibility.Collapsed;
        CollectEntitiesAll.IsChecked = false;
        foreach (var cb in EntityOptionChecks())
        {
            cb.IsChecked = false;
            cb.IsEnabled = false;
        }
    }

    private void OnCollectStartSearchChanged(object sender, TextChangedEventArgs e)
    {
        EnsurePopupOpen(CollectStartPopup);
        UpdateStartItems(GetSearchText(CollectStartSearch));
        FocusSearchBox(CollectStartSearch);
    }

    private void OnCollectEndSearchChanged(object sender, TextChangedEventArgs e)
    {
        EnsurePopupOpen(CollectEndPopup);
        UpdateEndItems(GetSearchText(CollectEndSearch));
        FocusSearchBox(CollectEndSearch);
    }

    private void OnCollectStartSearchPreviewKeyDown(object sender, KeyEventArgs e)
    {
        EnsurePopupOpen(CollectStartPopup);
        FocusSearchBox(CollectStartSearch);
    }

    private void OnCollectEndSearchPreviewKeyDown(object sender, KeyEventArgs e)
    {
        EnsurePopupOpen(CollectEndPopup);
        FocusSearchBox(CollectEndSearch);
    }

    private void OnCollectStartSearchGotFocus(object sender, RoutedEventArgs e)
    {
        EnsurePopupOpen(CollectStartPopup);
        FocusSearchBox(CollectStartSearch);
    }

    private void OnCollectEndSearchGotFocus(object sender, RoutedEventArgs e)
    {
        EnsurePopupOpen(CollectEndPopup);
        FocusSearchBox(CollectEndSearch);
    }

    private void OnCollectStartListClick(object sender, MouseButtonEventArgs e)
    {
        if (TryApplyListClick(CollectStartList, CollectStartInput, isStart: true, e))
        {
            e.Handled = true;
        }
    }

    private void OnCollectEndListClick(object sender, MouseButtonEventArgs e)
    {
        if (TryApplyListClick(CollectEndList, CollectEndInput, isStart: false, e))
        {
            e.Handled = true;
        }
    }

    private bool TryApplyListClick(ListBox list, TextBox input, bool isStart, MouseButtonEventArgs e)
    {
        var element = e.OriginalSource as DependencyObject;
        while (element != null && element is not ListBoxItem)
        {
            element = VisualTreeHelper.GetParent(element);
        }

        if (element is ListBoxItem item && item.Content is string s && DateTime.TryParse(s, out var dt))
        {
            SetInputText(input, dt.ToShortDateString());
            HideAllDropdowns();
            if (isStart)
            {
                UpdateEndItems(GetSearchText(CollectEndInput));
                ClampEndDate(forceToLatest: true);
            }
            else
            {
                ClampEndDate();
            }
            return true;
        }

        return false;
    }

    private void EnsurePopupOpen(Popup popup)
    {
        if (popup.IsOpen) return;
        Dispatcher.InvokeAsync(() => popup.IsOpen = true, DispatcherPriority.Background);
    }

    private void FocusSearchBox(TextBox box)
    {
        if (box.IsKeyboardFocusWithin) return;
        box.Focus();
        box.CaretIndex = box.Text.Length;
    }

    private void SetStartSelection(DateTime target)
    {
        if (CollectStartList.Items.Count == 0) return;
        var idx = CollectStartList.Items.IndexOf(FormatDate(target));
        if (idx < 0) idx = 0;
        CollectStartList.SelectedIndex = idx;
        var value = CollectStartList.SelectedItem as string ?? CollectStartInput.Text;
        SetInputText(CollectStartInput, value);
    }

    private void SetEndSelection(DateTime target)
    {
        if (CollectEndList.Items.Count == 0) return;
        var idx = CollectEndList.Items.IndexOf(FormatDate(target));
        if (idx < 0) idx = 0;
        CollectEndList.SelectedIndex = idx;
        var value = CollectEndList.SelectedItem as string ?? CollectEndInput.Text;
        SetInputText(CollectEndInput, value);
    }

    private void SetInputText(TextBox input, string value)
    {
        input.Text = value;
    }

    private string FormatDate(DateTime date) => date.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

    private DateTime ParseDate(string? text, DateTime fallback)
    {
        if (string.IsNullOrWhiteSpace(text)) return fallback;

        var formats = new[]
        {
            "dd.MM.yyyy",
            "d.M.yyyy",
            "yyyy-MM-dd",
            CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        };

        foreach (var fmt in formats.Distinct())
        {
            if (DateTime.TryParseExact(text, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exact))
                return exact.Date;
            if (DateTime.TryParseExact(text, fmt, new CultureInfo("ru-RU"), DateTimeStyles.None, out exact))
                return exact.Date;
        }

        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return dt.Date;
        if (DateTime.TryParse(text, new CultureInfo("ru-RU"), DateTimeStyles.None, out dt))
            return dt.Date;

        return fallback;
    }
}

