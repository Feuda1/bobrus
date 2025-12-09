using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Bobrus.App;

public partial class MainWindow
{
    private const string PluginsBaseUrl = "https://rapid.iiko.ru/plugins/";
    private const string PluginsInstallPath = @"C:\Program Files\iiko\iikoRMS\Front.Net\Plugins";
    private readonly List<PluginInfo> _plugins = new();
    private readonly List<PluginInfo> _filteredPlugins = new();
    private readonly List<PluginVersion> _versions = new();

    private async Task LoadPluginsAsync(bool showMessage = false)
    {
        try
        {
            PluginStatusText.Text = "Загрузка списка...";
            PluginsList.ItemsSource = null;
            _plugins.Clear();

            var html = await _httpClient.GetStringAsync(PluginsBaseUrl);
            var items = ParsePluginList(html);
            _plugins.AddRange(items.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase));

            ApplyPluginSearch();
            PluginStatusText.Text = $"Найдено: {_plugins.Count}";
            if (showMessage)
            {
                ShowNotification($"Загружено плагинов: {_plugins.Count}", NotificationType.Info);
            }
        }
        catch (Exception ex)
        {
            PluginStatusText.Text = "Ошибка загрузки";
            _logger.Error(ex, "Не удалось загрузить список плагинов");
            ShowNotification($"Не удалось загрузить список плагинов: {ex.Message}", NotificationType.Error);
        }
    }

    private async Task LoadPluginVersionsAsync(PluginInfo plugin)
    {
        try
        {
            PluginStatusText.Text = $"Загрузка версий {plugin.Name}...";
            PluginVersionsList.ItemsSource = null;
            _versions.Clear();
            InstallPluginButton.IsEnabled = false;

            var html = await _httpClient.GetStringAsync(plugin.Url);
            var versions = ParseVersions(plugin.Url, html);
            _versions.AddRange(versions.OrderByDescending(v => v.Name, StringComparer.OrdinalIgnoreCase));

            PluginVersionsList.ItemsSource = _versions;
            PluginStatusText.Text = $"Версий: {_versions.Count}";
        }
        catch (Exception ex)
        {
            PluginStatusText.Text = "Ошибка загрузки версий";
            _logger.Error(ex, "Не удалось загрузить версии плагина {Plugin}", plugin.Name);
            ShowNotification($"Не удалось загрузить версии: {ex.Message}", NotificationType.Error);
        }
    }

    private async void OnRefreshPluginsClicked(object sender, RoutedEventArgs e)
    {
        await LoadPluginsAsync(showMessage: true);
    }

    private async void OnPluginSelected(object sender, RoutedEventArgs e)
    {
        InstallPluginButton.IsEnabled = false;
        PluginVersionsList.ItemsSource = null;

        if (PluginsList.SelectedItem is not PluginInfo plugin)
        {
            return;
        }

        await LoadPluginVersionsAsync(plugin);
    }

    private void OnPluginsSearchChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        ApplyPluginSearch();
    }

    private void OnPluginVersionSelected(object sender, RoutedEventArgs e)
    {
        InstallPluginButton.IsEnabled = PluginVersionsList.SelectedItem is PluginVersion;
    }

    private async void OnInstallPluginClicked(object sender, RoutedEventArgs e)
    {
        if (PluginsList.SelectedItem is not PluginInfo plugin ||
            PluginVersionsList.SelectedItem is not PluginVersion version)
        {
            return;
        }

        InstallPluginButton.IsEnabled = false;
        PluginsList.IsEnabled = false;
        PluginStatusText.Text = "Скачивание...";

        var tempPath = Path.Combine(Path.GetTempPath(), $"{plugin.Name}_{version.Name}.zip");

        try
        {
            var progress = new Progress<int>(p =>
            {
                PluginStatusText.Text = $"Скачивание {p}%";
                UpdateGlobalProgress($"{plugin.Name} {version.Name}: {p}%", p);
            });

            await DownloadFileWithProgressAsync(version.Url, tempPath, progress, CancellationToken.None);
            PluginStatusText.Text = "Установка...";

            var targetFolder = Path.Combine(PluginsInstallPath, Path.GetFileNameWithoutExtension(version.Name));
            if (Directory.Exists(targetFolder))
            {
                Directory.Delete(targetFolder, true);
            }
            Directory.CreateDirectory(targetFolder);

            ZipFile.ExtractToDirectory(tempPath, targetFolder, overwriteFiles: true);
            RemoveZoneIdentifiers(targetFolder);
            TryDeleteFile(tempPath);

            ShowNotification($"Плагин {plugin.Name} ({version.Name}) установлен", NotificationType.Success);
            PluginStatusText.Text = "Готово";
            OpenFileInExplorer(targetFolder, selectFile: false);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка установки плагина {Plugin} {Version}", plugin.Name, version.Name);
            ShowNotification($"Ошибка установки: {ex.Message}", NotificationType.Error);
            PluginStatusText.Text = "Ошибка";
        }
        finally
        {
            HideGlobalProgress();
            InstallPluginButton.IsEnabled = PluginVersionsList.SelectedItem is not null;
            PluginsList.IsEnabled = true;
            PluginVersionsList.IsEnabled = true;
            TryDeleteFile(tempPath);
        }
    }

    private static IEnumerable<PluginInfo> ParsePluginList(string html)
    {
        var regex = new Regex("href=\"(?<href>[^\"?#]+/)\"", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        var results = new List<PluginInfo>();

        foreach (Match match in regex.Matches(html))
        {
            var href = match.Groups["href"].Value;
            if (string.IsNullOrWhiteSpace(href) || href.StartsWith("../"))
            {
                continue;
            }

            var rawName = href.TrimEnd('/').Trim();
            if (rawName.Length == 0)
            {
                continue;
            }

            var decodedName = WebUtility.UrlDecode(rawName);
            var url = new Uri(new Uri(PluginsBaseUrl), href).ToString();
            results.Add(new PluginInfo(decodedName, url));
        }

        return results;
    }

    private static IEnumerable<PluginVersion> ParseVersions(string baseUrl, string html)
    {
        var regex = new Regex("href=\"(?<href>[^\"?#]+\\.zip)\"", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        var list = new List<PluginVersion>();

        foreach (Match match in regex.Matches(html))
        {
            var href = match.Groups["href"].Value;
            if (string.IsNullOrWhiteSpace(href))
            {
                continue;
            }

            var rawName = Path.GetFileName(href);
            var name = WebUtility.UrlDecode(rawName);
            var url = new Uri(new Uri(baseUrl), href).ToString();
            list.Add(new PluginVersion(name, url));
        }

        return list;
    }

    private void ApplyPluginSearch()
    {
        _filteredPlugins.Clear();
        var query = (PluginsSearchBox.Text ?? string.Empty).Trim().ToLowerInvariant();

        if (string.IsNullOrWhiteSpace(query))
        {
            _filteredPlugins.AddRange(_plugins);
        }
        else
        {
            _filteredPlugins.AddRange(_plugins.Where(p => p.DisplayName.ToLowerInvariant().Contains(query)));
        }

        PluginsList.ItemsSource = null;
        PluginsList.ItemsSource = _filteredPlugins;
        PluginStatusText.Text = $"Найдено: {_filteredPlugins.Count}";
    }

    private void RemoveZoneIdentifiers(string folder)
    {
        try
        {
            foreach (var file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
            {
                var zonePath = file + ":Zone.Identifier";
                try
                {
                    if (File.Exists(zonePath))
                    {
                        File.Delete(zonePath);
                    }
                }
                catch
                {
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Не удалось удалить Zone.Identifier в {Folder}", folder);
        }
    }

    private sealed record PluginInfo(string Name, string Url)
    {
        public string DisplayName => Name;
    }

    private sealed record PluginVersion(string Name, string Url)
    {
        public string DisplayName => Name;
    }
}
