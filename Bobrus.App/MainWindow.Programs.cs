using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Button = System.Windows.Controls.Button;

namespace Bobrus.App;

public partial class MainWindow
{
    private readonly string _downloadDirectory = AppPaths.DownloadsDirectory;

    static MainWindow()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private const string SevenZipUrl = "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/7z2500-x64.exe";
    private const string AdvancedIpScannerUrl = "https://download.advanced-ip-scanner.com/download/files/Advanced_IP_Scanner_2.5.4594.1.exe";
    private const string AnyDeskUrl = "https://download.anydesk.com/AnyDesk.exe";
    private const string AssistantUrl = "https://мойассистент.рф/%D1%81%D0%BA%D0%B0%D1%87%D0%B0%D1%82%D1%8C/Download/1369";
    private const string ComPortCheckerUrl = "https://atol-kassa.ru/wp-content/nfiles/files/ATOL/soft/zx/comportchecker/ComPortChecker%201.1.zip";
    private const string DatabaseNetUrl = "https://fishcodelib.com/files/DatabaseNet4.zip";
    private const string Notepad64Url = "https://release-assets.githubusercontent.com/github-production-release-asset/33014811/71731d1b-bd3e-4ad8-9835-cdda378f5599?sp=r&sv=2018-11-09&sr=b&spr=https&se=2025-12-05T13%3A45%3A24Z&rscd=attachment%3B+filename%3Dnpp.8.8.8.Installer.x64.exe&rsct=application%2Foctet-stream&skoid=96c2d410-5711-43a1-aedd-ab1947aa7ab0&sktid=398a6654-997b-47e9-b12b-9515b896b4de&skt=2025-12-05T12%3A44%3A36Z&ske=2025-12-05T13%3A45%3A24Z&sks=b&skv=2018-11-09&sig=SkNxkKmXjHFgF7C3dJgYiypnLAS5qGhvzAIk2Y8KdOQ%3D&jwt=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmVsZWFzZS1hc3NldHMuZ2l0aHVidXNlcmNvbnRlbnQuY29tIiwia2V5Ijoia2V5MSIsImV4cCI6MTc2NDkzODk3NiwibmJmIjoxNzY0OTM4Njc2LCJwYXRoIjoicmVsZWFzZWFzc2V0cHJvZHVjdGlvbi5ibG9iLmNvcmUud2luZG93cy5uZXQifQ._rVBza1bwdykfIJn9PlAJKnHmX9-PudsuEm6WnWHcPQ&response-content-disposition=attachment%3B%20filename%3Dnpp.8.8.8.Installer.x64.exe&response-content-type=application%2Foctet-stream";
    private const string Notepad32Url = "https://release-assets.githubusercontent.com/github-production-release-asset/33014811/c7fc9bf4-e8ea-4031-9e7d-af0356f0a2a3?sp=r&sv=2018-11-09&sr=b&spr=https&se=2025-12-05T13%3A38%3A07Z&rscd=attachment%3B+filename%3Dnpp.8.8.8.Installer.exe&rsct=application%2Foctet-stream&skoid=96c2d410-5711-43a1-aedd-ab1947aa7ab0&sktid=398a6654-997b-47e9-b12b-9515b896b4de&skt=2025-12-05T12%3A37%3A27Z&ske=2025-12-05T13%3A38%3A07Z&sks=b&skv=2018-11-09&sig=kwBJInyIdoNPYH635oZJ%2BjPIerZhGfWI0lfxmfO%2Bayw%3D&jwt=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmVsZWFzZS1hc3NldHMuZ2l0aHVidXNlcmNvbnRlbnQuY29tIiwia2V5Ijoia2V5MSIsImV4cCI6MTc2NDkzODk3OCwibmJmIjoxNzY0OTM4Njc4LCJwYXRoIjoicmVsZWFzZWFzc2V0cHJvZHVjdGlvbi5ibG9iLmNvcmUud2luZG93cy5uZXQifQ.CT129pSJFIytToDoZuSA77cgUvLnXWmNRS7IremhSLg&response-content-disposition=attachment%3B%20filename%3Dnpp.8.8.8.Installer.exe&response-content-type=application%2Foctet-stream";
    private const string PrinterTestUrl = "https://725920.selcdn.ru/upload_portkkm/iblock/929/9291b0d0c76c7c5756b81732df929086/Printer-TEST-V3.1C.zip";
    private const string RhelperUrl = "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/remote-access-setup.exe";
    private const string OrderCheckUrl = "https://clearbat.iiko.online/downloads/OrderCheck.exe";
    private const string ClearBatUrl = "https://clearbat.iiko.online/downloads/CLEAR.bat.exe";
    private const string FrontToolsUrl = "https://fronttools.iiko.it/FrontTools.exe";
    private const string FrontToolsSqliteUrl = "https://fronttools.iiko.it/fronttools_sqlite.zip";

    private void OnSevenZipClicked(object sender, RoutedEventArgs e)
    {
        _ = DownloadProgramAsync("7-Zip", SevenZipUrl, SevenZipButton, handling: ProgramHandling.RunInstallerThenReveal);
    }

    private async void OnAdvancedIpScannerClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Advanced IP Scanner", AdvancedIpScannerUrl, AdvancedIpScannerButton, handling: ProgramHandling.RunInstallerThenReveal);

    private async void OnAnyDeskClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("AnyDesk", AnyDeskUrl, AnyDeskButton, handling: ProgramHandling.RevealOnly);

    private async void OnAssistantClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Ассистент", AssistantUrl, AssistantButton, "Assistant.exe", handling: ProgramHandling.RunInstallerThenReveal);

    private async void OnComPortCheckerClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Com Port Checker", ComPortCheckerUrl, ComPortCheckerButton, handling: ProgramHandling.ExtractAndRevealFolder);

    private async void OnDatabaseNetClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Database Net", DatabaseNetUrl, DatabaseNetButton, handling: ProgramHandling.ExtractAndRevealFolder);

    private void OnNotepadClicked(object sender, RoutedEventArgs e)
    {
        ShowProgramOptions(NotepadButton,
            new ProgramOption("Notepad++ 64-bit", () => DownloadProgramAsync("Notepad++ 64-bit", Notepad64Url, NotepadButton, handling: ProgramHandling.RunInstallerThenReveal)),
            new ProgramOption("Notepad++ 32-bit", () => DownloadProgramAsync("Notepad++ 32-bit", Notepad32Url, NotepadButton, handling: ProgramHandling.RunInstallerThenReveal)));
    }

    private async void OnNotepad64Clicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Notepad++ 64-bit", Notepad64Url, NotepadButton, handling: ProgramHandling.RunInstallerThenReveal);

    private async void OnNotepad32Clicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Notepad++ 32-bit", Notepad32Url, NotepadButton, handling: ProgramHandling.RunInstallerThenReveal);

    private async void OnPrinterTestClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("Printer Test V3.1C", PrinterTestUrl, PrinterTestButton, handling: ProgramHandling.ExtractAndRevealFolder);

    private async void OnRhelperClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync(
            "Rhelper",
            RhelperUrl,
            RhelperButton,
            suggestedFileName: null,
            customFolderName: "Rhelper",
            afterDownload: CreateRhelperPasswordFile);

    private async void OnOrderCheckClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("OrderCheck", OrderCheckUrl, OrderCheckButton, handling: ProgramHandling.RevealOnly);

    private async void OnClearBatClicked(object sender, RoutedEventArgs e) =>
        await DownloadProgramAsync("CLEAR.bat", ClearBatUrl, ClearBatButton, handling: ProgramHandling.RunInstallerThenReveal);

    private void OnFrontToolsClicked(object sender, RoutedEventArgs e)
    {
        ShowProgramOptions(FrontToolsButton,
            new ProgramOption("FrontTools (iikoTools)", () => DownloadProgramAsync("FrontTools (iikoTools)", FrontToolsUrl, FrontToolsButton, handling: ProgramHandling.RevealOnly)),
            new ProgramOption("FrontTools SQLite", () => DownloadProgramAsync("FrontTools SQLite", FrontToolsSqliteUrl, FrontToolsButton, handling: ProgramHandling.ExtractAndRevealFolder)));
    }

    private void ShowProgramOptions(Button button, params ProgramOption[] options)
    {
        HideAllDropdowns();
        ProgramOptionsPanel.Children.Clear();
        ProgramOptionsPopup.PlacementTarget = button;
        ProgramOptionsPopup.MinWidth = Math.Max(button.ActualWidth, 220);
        ProgramOptionsPopup.VerticalOffset = 4;

        foreach (var option in options)
        {
            var optButton = new Button
            {
                Content = option.Title,
                Style = FindResource("BaseButton") as Style,
                Height = 32,
                Margin = new Thickness(0, 4, 0, 4),
                HorizontalContentAlignment = System.Windows.HorizontalAlignment.Left,
                IsEnabled = option.IsEnabled
            };

            optButton.Click += async (_, _) =>
            {
                ProgramOptionsPopup.IsOpen = false;
                await option.Action();
            };

            ProgramOptionsPanel.Children.Add(optButton);
        }

        ProgramOptionsPopup.IsOpen = true;
    }

    private async Task DownloadProgramAsync(
        string displayName,
        string url,
        Button button,
        string? suggestedFileName = null,
        string? customFolderName = null,
        Action<string>? afterDownload = null,
        ProgramHandling handling = ProgramHandling.RevealOnly)
    {
        var originalContent = button.Content;
        button.IsEnabled = false;
        UpdateGlobalProgress($"{displayName}: 0%", 0);
        ShowNotification($"Скачиваем {displayName}...", NotificationType.Info);

        try
        {
            var progress = new Progress<int>(p =>
            {
                if (p > 0 && p < 100)
                {
                    button.Content = $"{originalContent} ({p}%)";
                }
                UpdateGlobalProgress($"{displayName}: {p}%", p);
            });

            var destination = await DownloadToDownloadsAsync(displayName, url, progress, suggestedFileName, customFolderName, CancellationToken.None);
            UpdateGlobalProgress($"{displayName}: 100%", 100);
            ShowNotification($"{displayName} скачан: {destination}", NotificationType.Success);
            _logger.Information("{Name} скачан в {Path}", displayName, destination);
            afterDownload?.Invoke(destination);

            var revealPath = destination;
            var selectFile = true;

            switch (handling)
            {
                case ProgramHandling.ExtractAndRevealFolder:
                    try
                    {
                        revealPath = ExtractArchive(destination, displayName);
                        selectFile = false;
                        ShowNotification($"{displayName} распакован: {revealPath}", NotificationType.Success);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Ошибка распаковки {Name}", displayName);
                        ShowNotification($"Не удалось распаковать {displayName}: {ex.Message}", NotificationType.Error);
                        revealPath = destination;
                        selectFile = true;
                    }
                    break;
                case ProgramHandling.RunInstallerThenReveal:
                    StartInstaller(destination, displayName);
                    break;
                default:
                    break;
            }

            OpenFileInExplorer(revealPath, selectFile);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка скачивания {Name} из {Url}", displayName, url);
            ShowNotification($"Не удалось скачать {displayName}: {ex.Message}", NotificationType.Error);
        }
        finally
        {
            button.Content = originalContent;
            button.IsEnabled = true;
            HideGlobalProgress();
        }
    }

    private async Task<string> DownloadToDownloadsAsync(string displayName, string url, IProgress<int>? progress, string? suggestedFileName, string? customFolderName, CancellationToken token)
    {
        var targetFolder = ResolveDownloadFolder(customFolderName);
        Directory.CreateDirectory(targetFolder);

        using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token);
        response.EnsureSuccessStatusCode();

        var fileName = ResolveFileName(response, url, suggestedFileName, displayName);
        var destination = EnsureUniquePath(Path.Combine(targetFolder, fileName));

        await using var contentStream = await response.Content.ReadAsStreamAsync(token);
        await using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);

        var total = response.Content.Headers.ContentLength ?? -1L;
        var canReport = total > 0 && progress != null;
        var buffer = new byte[81920];
        long totalRead = 0;
        var lastPercent = 0;

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
        return destination;
    }

    private static string ResolveFileName(HttpResponseMessage response, string url, string? suggestedFileName, string displayName)
    {
        var headerName = response.Content.Headers.ContentDisposition?.FileNameStar
                         ?? response.Content.Headers.ContentDisposition?.FileName;

        var candidates = new[]
        {
            headerName,
            GetFileNameFromUrl(url),
            suggestedFileName,
            $"{displayName}.bin"
        };

        foreach (var candidate in candidates)
        {
            var cleaned = SanitizeFileName(candidate);
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                return cleaned;
            }
        }

        return "download.bin";
    }

    private static string? GetFileNameFromUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var segment = uri.Segments.LastOrDefault();
        if (string.IsNullOrWhiteSpace(segment))
        {
            return null;
        }

        var trimmed = segment.Trim('/');
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return null;
        }

        return WebUtility.UrlDecode(trimmed);
    }

    private static string SanitizeFileName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return string.Empty;
        }

        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Where(ch => !invalid.Contains(ch)).ToArray());
        return string.IsNullOrWhiteSpace(cleaned) ? string.Empty : cleaned;
    }

    private static string EnsureUniquePath(string path)
    {
        var directory = Path.GetDirectoryName(path) ?? Directory.GetCurrentDirectory();
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        var candidate = path;
        var counter = 1;

        while (File.Exists(candidate))
        {
            candidate = Path.Combine(directory, $"{name} ({counter}){extension}");
            counter++;
        }

        return candidate;
    }

    private static string EnsureUniqueDirectory(string basePath)
    {
        var directory = basePath;
        var counter = 1;
        while (Directory.Exists(directory))
        {
            directory = $"{basePath} ({counter})";
            counter++;
        }

        return directory;
    }

    private string ResolveDownloadFolder(string? customFolderName)
    {
        if (string.IsNullOrWhiteSpace(customFolderName))
        {
            return _downloadDirectory;
        }

        return Path.Combine(_downloadDirectory, customFolderName);
    }

    private string ExtractArchive(string archivePath, string displayName)
    {
        var baseFolder = Path.Combine(Path.GetDirectoryName(archivePath) ?? _downloadDirectory, Path.GetFileNameWithoutExtension(archivePath));
        var targetFolder = EnsureUniqueDirectory(baseFolder);
        Directory.CreateDirectory(targetFolder);

        using (var archive = ZipFile.OpenRead(archivePath))
        {
            foreach (var entry in archive.Entries)
            {
                var decodedName = DecodeEntryName(entry.FullName);
                if (string.IsNullOrWhiteSpace(decodedName))
                {
                    continue;
                }

                var sanitizedName = decodedName.Replace('\\', '/').TrimStart('/', '\\');
                if (sanitizedName.Length == 0)
                {
                    continue;
                }

                var destinationPath = Path.Combine(targetFolder, sanitizedName);
                var fullPath = Path.GetFullPath(destinationPath);
                if (!fullPath.StartsWith(targetFolder, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (decodedName.EndsWith("/") || decodedName.EndsWith("\\") || entry.Length == 0 && decodedName.EndsWith("/"))
                {
                    Directory.CreateDirectory(fullPath);
                    continue;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                using var entryStream = entry.Open();
                using var fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None);
                entryStream.CopyTo(fileStream);
            }
        }

        TryRenameGibberishFiles(targetFolder, displayName);

        try
        {
            File.Delete(archivePath);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось удалить архив {Path}", archivePath);
        }

        return targetFolder;
    }

    private string DecodeEntryName(string name)
    {
        var enc437 = Encoding.GetEncoding(437);
        var candidates = new[]
        {
            ("cp866", Encoding.GetEncoding(866)),
            ("cp1251", Encoding.GetEncoding(1251))
        };

        var bestName = name;
        var bestScore = ScoreCyrillic(name);

        foreach (var (_, targetEnc) in candidates)
        {
            var reencoded = TryReencode(name, enc437, targetEnc);
            var score = ScoreCyrillic(reencoded);
            if (score > bestScore)
            {
                bestScore = score;
                bestName = reencoded;
            }
        }

        return bestName;
    }

    private static string TryReencode(string text, Encoding from, Encoding to)
    {
        var bytes = from.GetBytes(text);
        return to.GetString(bytes);
    }

    private static int ScoreCyrillic(string text)
    {
        var count = 0;
        foreach (var ch in text)
        {
            if ((ch >= 'А' && ch <= 'я') || ch == 'Ё' || ch == 'ё')
            {
                count++;
            }
        }
        return count;
    }

    private void TryRenameGibberishFiles(string folder, string displayName)
    {
        var files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
        var index = 1;

        foreach (var file in files)
        {
            var name = Path.GetFileName(file);
            if (!IsGibberish(name))
            {
                continue;
            }

            var ext = Path.GetExtension(file);
            var newName = index == 1
                ? $"{displayName}{ext}"
                : $"{displayName} ({index}){ext}";
            index++;

            var newPath = EnsureUniquePath(Path.Combine(Path.GetDirectoryName(file)!, newName));
            try
            {
                File.Move(file, newPath);
                _logger.Information("Переименован файл после распаковки: {Old} -> {New}", file, newPath);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Не удалось переименовать {Path}", file);
            }
        }
    }

    private static bool IsGibberish(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        if (name.Contains('�'))
        {
            return true;
        }

        var hasLetters = name.Any(ch =>
            (ch >= 'A' && ch <= 'Z') ||
            (ch >= 'a' && ch <= 'z') ||
            (ch >= 'А' && ch <= 'я') ||
            ch == 'Ё' || ch == 'ё');

        return !hasLetters;
    }

    private void StartInstaller(string path, string displayName)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = path,
                WorkingDirectory = Path.GetDirectoryName(path) ?? string.Empty,
                UseShellExecute = true
            };

            Process.Start(psi);
            _logger.Information("Установщик запущен: {Name}", displayName);
            ShowNotification($"{displayName}: установщик запущен", NotificationType.Info);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось запустить установщик {Name}", displayName);
            ShowNotification($"Не удалось запустить {displayName}: {ex.Message}", NotificationType.Error);
        }
    }

    private void UpdateGlobalProgress(string text, int percent)
    {
        Dispatcher.Invoke(() =>
        {
            GlobalProgressPanel.Visibility = Visibility.Visible;
            GlobalProgressText.Text = text;
            GlobalProgressBar.Value = Math.Clamp(percent, 0, 100);
        });
    }

    private void HideGlobalProgress()
    {
        Dispatcher.Invoke(() =>
        {
            GlobalProgressPanel.Visibility = Visibility.Collapsed;
            GlobalProgressText.Text = string.Empty;
            GlobalProgressBar.Value = 0;
        });
    }

    private void OpenFileInExplorer(string path, bool selectFile)
    {
        try
        {
            if (selectFile)
            {
                if (!File.Exists(path))
                {
                    ShowNotification($"Файл не найден: {path}", NotificationType.Warning);
                    return;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{path}\"",
                    UseShellExecute = true
                };
                Process.Start(psi);
                return;
            }

            if (!Directory.Exists(path))
            {
                ShowNotification($"Папка не найдена: {path}", NotificationType.Warning);
                return;
            }

            var folderPsi = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            };
            Process.Start(folderPsi);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось открыть проводник для {Path}", path);
            ShowNotification($"Не удалось открыть папку: {ex.Message}", NotificationType.Error);
        }
    }

    private void CreateRhelperPasswordFile(string downloadedPath)
    {
        try
        {
            var directory = Path.GetDirectoryName(downloadedPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            Directory.CreateDirectory(directory);
            var passwordFile = Path.Combine(directory, "пароль.txt");
            File.WriteAllText(passwordFile, "remote-access-setup");
            _logger.Information("Парольный файл создан: {Path}", passwordFile);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Не удалось создать файл пароля для Rhelper");
            ShowNotification("Не удалось записать пароль для Rhelper", NotificationType.Warning);
        }
    }

    private sealed record ProgramOption(string Title, Func<Task> Action, bool IsEnabled = true);

    private enum ProgramHandling
    {
        RevealOnly,
        RunInstallerThenReveal,
        ExtractAndRevealFolder
    }
}
