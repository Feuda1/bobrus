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
    private const string PluginsInstallPath = @"C:\Program Files\iiko\iikoRMS\Front.Net\Plugins";
    private readonly Services.PluginRepository _pluginRepo = new(new HttpClient { Timeout = TimeSpan.FromSeconds(30) });
    private readonly List<Services.PluginInfo> _plugins = new();
    private readonly List<Services.PluginInfo> _filteredPlugins = new();
    private readonly List<Services.PluginVersion> _versions = new();

    private async Task LoadPluginsAsync(bool showMessage = false)
    {
        try
        {
            PluginStatusText.Text = "Загрузка списка...";
            PluginsList.ItemsSource = null;
            _plugins.Clear();

            var items = await _pluginRepo.GetPluginsAsync();
            _plugins.AddRange(items);

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

    private async Task LoadPluginVersionsAsync(Services.PluginInfo plugin)
    {
        try
        {
            PluginStatusText.Text = $"Загрузка версий {plugin.Name}...";
            PluginVersionsList.ItemsSource = null;
            _versions.Clear();
            InstallPluginButton.IsEnabled = false;

            var versions = await _pluginRepo.GetVersionsAsync(plugin.Url);
            _versions.AddRange(versions);

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

        if (PluginsList.SelectedItem is not Services.PluginInfo plugin)
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
        InstallPluginButton.IsEnabled = PluginVersionsList.SelectedItem is Services.PluginVersion;
    }

    private async void OnInstallPluginClicked(object sender, RoutedEventArgs e)
    {
        if (PluginsList.SelectedItem is not Services.PluginInfo plugin ||
            PluginVersionsList.SelectedItem is not Services.PluginVersion version)
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
}
