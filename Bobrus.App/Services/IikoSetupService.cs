using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed record IikoSetupOptions
{
    public string ServerUrl { get; init; } = "";
    public bool InstallFront { get; init; }
    public bool InstallOffice { get; init; }
    public bool InstallChain { get; init; }
    public bool InstallCard { get; init; } = true;
    public bool FrontAutostart { get; init; } = true;
    public bool EnableHandCardRoll { get; init; } = true;
    public bool EnableMinimizeButton { get; init; } = true;
    public bool SetServerUrl { get; init; } = true;
    public List<PluginVersion> Plugins { get; init; } = new();
}

internal sealed class IikoSetupService
{
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private const string IikoCardUrl = "https://iiko.biz/ru-RU/About/DownloadPosInstaller?useRc=False";
    private const string IikoFrontExePath = @"C:\Program Files\iiko\iikoRMS\Front.Net\iikoFront.Net.exe";
    private const string IikoPluginsPath = @"C:\Program Files\iiko\iikoRMS\Front.Net\Plugins";
    private static readonly string IikoCashServerPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "iiko", "CashServer", "config.xml");

    public async Task<bool> InstallAsync(
        IikoSetupOptions options,
        IProgress<string>? progress,
        SetupFlowController controller)
    {
        var ct = controller.Token;
        var needsServer = options.InstallFront || options.InstallOffice || options.InstallChain;
        if (needsServer && string.IsNullOrWhiteSpace(options.ServerUrl))
            return false;

        var serverUrl = !string.IsNullOrWhiteSpace(options.ServerUrl) 
            ? NormalizeServerUrl(options.ServerUrl) 
            : "";
        string? version = null;
        bool anyInstalled = false;
        if (options.InstallFront)
        {
            await controller.WaitIfPausedAsync();
            progress?.Report("[IIKO] Получение версии iikoFront...");
            var frontInfo = await GetVersionInfoAsync($"{serverUrl}/update/Front/updates.ini", ct);
            if (frontInfo.HasValue)
            {
                version = frontInfo.Value.Version;
                progress?.Report($"[IIKO] Скачивание iikoFront {version}...");
                var ok = await DownloadAndInstallAsync(frontInfo.Value.DownloadUrl, "iikoFront", progress, ct);
                if (ok) 
                {
                    anyInstalled = true;
                    if (options.FrontAutostart)
                    {
                        await AddIikoFrontToStartupAsync(progress);
                    }
                }
            }
            else
            {
                progress?.Report("⚠ [IIKO] Не удалось получить версию iikoFront");
            }
            if (options.InstallFront) 
            {
                 await ConfigureFrontAsync(options, progress, ct);
                 if (options.Plugins is { Count: > 0 })
                 {
                     await InstallPluginsAsync(options.Plugins, progress, ct);
                 }
            }
        }
        if (options.InstallOffice)
        {
            await controller.WaitIfPausedAsync();
            progress?.Report("[IIKO] Получение версии iikoOffice...");
            var officeInfo = await GetVersionInfoAsync($"{serverUrl}/update/BackOffice/updates.ini", ct);
            if (officeInfo.HasValue)
            {
                version ??= officeInfo.Value.Version;
                progress?.Report($"[IIKO] Скачивание iikoOffice {officeInfo.Value.Version}...");
                var ok = await DownloadAndInstallAsync(officeInfo.Value.DownloadUrl, "iikoOffice", progress, ct);
                if (ok) anyInstalled = true;
            }
            else
            {
                progress?.Report("⚠ [IIKO] Не удалось получить версию iikoOffice");
            }
        }
        if (options.InstallChain)
        {
            await controller.WaitIfPausedAsync();
            if (string.IsNullOrEmpty(version))
            {
                progress?.Report("[IIKO] Получение версии для iikoChain...");
                var frontInfo = await GetVersionInfoAsync($"{serverUrl}/update/Front/updates.ini", ct);
                if (frontInfo.HasValue)
                {
                    version = frontInfo.Value.Version;
                }
                else
                {
                    var officeInfo = await GetVersionInfoAsync($"{serverUrl}/update/BackOffice/updates.ini", ct);
                    if (officeInfo.HasValue)
                    {
                        version = officeInfo.Value.Version;
                    }
                }
            }

            if (!string.IsNullOrEmpty(version))
            {
                progress?.Report($"[IIKO] Скачивание iikoChain {version}...");
                var chainUrl = $"https://downloads.iiko.online/{version}/iiko/Chain/BackOffice/Setup.Chain.BackOffice.exe";
                var ok = await DownloadAndInstallAsync(chainUrl, "iikoChain", progress, ct);
                if (ok) anyInstalled = true;
            }
            else
            {
                progress?.Report("⚠ [IIKO] Не удалось определить версию для iikoChain");
            }
        }
        if (options.InstallCard)
        {
            await controller.WaitIfPausedAsync();
            progress?.Report("[IIKO] Скачивание iikoCard...");
            var ok = await DownloadAndInstallAsync(IikoCardUrl, "iikoCard", progress, ct);
            if (ok) anyInstalled = true;
        }

        progress?.Report(anyInstalled ? "✔ [IIKO] Установка завершена" : "⚠ [IIKO] Ничего не установлено");
        return anyInstalled;
    }
    public static string NormalizeServerUrl(string url)
    {
        url = url.Trim();
        var withoutProtocol = url
            .Replace("https://", "", StringComparison.OrdinalIgnoreCase)
            .Replace("http://", "", StringComparison.OrdinalIgnoreCase);
        withoutProtocol = Regex.Replace(withoutProtocol, @":443\b", "");
        var domain = withoutProtocol.Replace("/resto", "", StringComparison.OrdinalIgnoreCase).TrimEnd('/');
        if (!domain.Contains('.'))
        {
            url = domain + ".iiko.it";
        }
        else
        {
            url = domain;
        }
        url = "https://" + url;
        url = url.TrimEnd('/');
        if (!url.EndsWith("/resto", StringComparison.OrdinalIgnoreCase))
        {
            url += "/resto";
        }

        return url;
    }
    private async Task<(string Version, string DownloadUrl)?> GetVersionInfoAsync(string updatesIniUrl, CancellationToken ct)
    {
        try
        {
            var content = await _http.GetStringAsync(updatesIniUrl, ct);
            var fileMatch = Regex.Match(content, @"FileName=([^\|]+)");
            var versionMatch = Regex.Match(content, @"Version=(\d+\.\d+\.\d+\.\d+)");

            if (fileMatch.Success && versionMatch.Success)
            {
                return (versionMatch.Groups[1].Value, fileMatch.Groups[1].Value);
            }
        }
        catch
        {
        }

        return null;
    }
    private async Task<bool> DownloadAndInstallAsync(
        string downloadUrl,
        string componentName,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        var tempDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Bobrus", "iiko_installers");
        Directory.CreateDirectory(tempDir);

        var fileName = $"{componentName}_Setup.exe";
        var filePath = Path.Combine(tempDir, fileName);
        try { if (File.Exists(filePath)) File.Delete(filePath); } catch { }

        try
        {
            using var response = await _http.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? 0;
            var buffer = new byte[81920]; 
            long downloadedBytes = 0;
            int lastPercent = -1;
            await using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
            await using (var contentStream = await response.Content.ReadAsStreamAsync(ct))
            {
                int bytesRead;
                while ((bytesRead = await contentStream.ReadAsync(buffer, ct)) > 0)
                {
                    await fs.WriteAsync(buffer.AsMemory(0, bytesRead), ct);
                    downloadedBytes += bytesRead;

                    if (totalBytes > 0)
                    {
                        var percent = (int)(downloadedBytes * 100 / totalBytes);
                        if (percent != lastPercent && percent % 10 == 0)
                        {
                            lastPercent = percent;
                            var mb = downloadedBytes / 1024.0 / 1024.0;
                            var totalMb = totalBytes / 1024.0 / 1024.0;
                            progress?.Report($"[IIKO] Скачивание {componentName}: {percent}% ({mb:F1}/{totalMb:F1} МБ)");
                        }
                    }
                }
            }

            progress?.Report($"[IIKO] Установка {componentName}...");
            await Task.Delay(500, ct);
            var psi = new ProcessStartInfo
            {
                FileName = filePath,
                Arguments = "/S /silent /quiet",
                UseShellExecute = true,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            if (proc != null)
            {
                await proc.WaitForExitAsync(ct);
                var success = proc.ExitCode == 0 || proc.ExitCode == 1;
                progress?.Report(success
                    ? $"✔ [IIKO] {componentName} установлен"
                    : $"⚠ [IIKO] {componentName} завершился с кодом {proc.ExitCode}");
                return success;
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"✖ [IIKO] Ошибка {componentName}: {ex.Message}");
        }
        finally
        {
            try { await Task.Delay(1000, ct); } catch (OperationCanceledException) { }
            try { if (File.Exists(filePath)) File.Delete(filePath); } catch { }
        }

        return false;
    }
    private static async Task AddIikoFrontToStartupAsync(IProgress<string>? progress)
    {
        try
        {
            progress?.Report("[IIKO] Добавление iikoFront в автозагрузку...");
            await Task.Delay(3000);
            for (int i = 0; i < 5; i++)
            {
                if (File.Exists(IikoFrontExePath))
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                    
                    key?.SetValue("iikoFront", $"\"{IikoFrontExePath}\"");
                    
                    progress?.Report("✔ [IIKO] iikoFront добавлен в автозагрузку");
                    return;
                }
                await Task.Delay(2000);
            }
            
            progress?.Report("⚠ [IIKO] Не удалось найти iikoFront.exe для автозагрузки");
        }
        catch (Exception ex)
        {
            progress?.Report($"⚠ [IIKO] Ошибка автозагрузки: {ex.Message}");
        }
    }

    public async Task InstallIikoAsync(WindowsSetupOptions options, IProgress<string> progress, SetupFlowController controller)
    {
        var iikoOptions = new IikoSetupOptions
        {
            ServerUrl = options.IikoServerUrl,
            InstallFront = options.IikoFront,
            InstallOffice = options.IikoOffice,
            InstallChain = options.IikoChain,
            InstallCard = options.IikoCard,
            FrontAutostart = options.IikoFrontAutostart,
            EnableHandCardRoll = options.IikoHandCardRoll,
            EnableMinimizeButton = options.IikoMinimizeButton,
            SetServerUrl = options.IikoSetServerUrl,
            Plugins = options.IikoPlugins
        };

        await InstallAsync(iikoOptions, progress, controller);
    }

    private async Task ConfigureFrontAsync(IikoSetupOptions options, IProgress<string>? progress, CancellationToken ct)
    {
        if (!options.EnableHandCardRoll && !options.EnableMinimizeButton && !options.SetServerUrl)
            return;

        try
        {
             if (!File.Exists(IikoFrontExePath))
             {
                 progress?.Report("⚠ [IIKO] iikoFront не найден, пропуск настройки конфига");
                 return;
             }

             progress?.Report("[IIKO] Запуск iikoFront для генерации конфига (10 сек)...");
             var psi = new ProcessStartInfo
             {
                 FileName = IikoFrontExePath,
                 UseShellExecute = true,
                 WorkingDirectory = Path.GetDirectoryName(IikoFrontExePath)
             };

             using var process = Process.Start(psi);
             if (process != null)
             {
                 try 
                 {
                    await Task.Delay(TimeSpan.FromSeconds(10), ct);
                    if (!process.HasExited)
                    {
                        process.Kill();
                    }
                 }
                 catch {}
             }
             
             if (!File.Exists(IikoCashServerPath))
             {
                 progress?.Report($"⚠ [IIKO] Конфиг не найден: {IikoCashServerPath}");
                 return;
             }

             progress?.Report("[IIKO] Применение настроек конфига...");
             var xml = await File.ReadAllTextAsync(IikoCashServerPath, ct);
             bool changed = false;

             if (options.EnableHandCardRoll)
             {
                 if (TryPatchXmlTag(ref xml, "AllowHandCardRoll", "true")) changed = true;
             }
             
             if (options.EnableMinimizeButton)
             {
                 if (TryPatchXmlTag(ref xml, "ShowMinimizeButton", "true")) changed = true;
             }

             if (options.SetServerUrl && !string.IsNullOrWhiteSpace(options.ServerUrl))
             {
                 var normUrl = NormalizeServerUrl(options.ServerUrl);
                 if (TryPatchXmlTag(ref xml, "serverUrl", normUrl)) changed = true;
             }

             if (changed)
             {
                 await File.WriteAllTextAsync(IikoCashServerPath, xml, ct);
                 progress?.Report("✔ [IIKO] Конфиг обновлен");
             }
             else
             {
                 progress?.Report("[IIKO] Конфиг актуален");
             }

        }
        catch (Exception ex)
        {
             progress?.Report($"⚠ [IIKO] Ошибка настройки конфига: {ex.Message}");
        }
    }

    private static bool TryPatchXmlTag(ref string xml, string tagName, string newValue)
    {
        var pattern = $@"<{tagName}>.*?</{tagName}>";
        var replacement = $"<{tagName}>{newValue}</{tagName}>";
        
        if (Regex.IsMatch(xml, pattern))
        {
             var match = Regex.Match(xml, pattern);
             if (match.Value != replacement)
             {
                 xml = Regex.Replace(xml, pattern, replacement);
                 return true;
             }
        }
        else
        {
        }
        return false;
    }

    private async Task InstallPluginsAsync(List<PluginVersion> plugins, IProgress<string>? progress, CancellationToken ct)
    {
        if (!Directory.Exists(IikoPluginsPath))
        {
            try { Directory.CreateDirectory(IikoPluginsPath); }
            catch {
                progress?.Report("⚠ [IIKO] Не удалось создать папку Plugins");
                return;
            }
        }

        foreach (var plugin in plugins)
        {
            ct.ThrowIfCancellationRequested();
            
            try
            {
                progress?.Report($"[IIKO] Установка плагина {plugin.Name}...");
                var tempZip = Path.Combine(Path.GetTempPath(), $"iiko_plugin_{Guid.NewGuid()}.zip");
                using (var response = await _http.GetAsync(plugin.Url, HttpCompletionOption.ResponseHeadersRead, ct))
                {
                    response.EnsureSuccessStatusCode();
                    await using var fs = new FileStream(tempZip, FileMode.Create);
                    await response.Content.CopyToAsync(fs, ct);
                }
                UnblockFile(tempZip);
                var pluginName = Path.GetFileNameWithoutExtension(plugin.Name); 
                
                var targetDir = Path.Combine(IikoPluginsPath, pluginName);
                if (Directory.Exists(targetDir)) Directory.Delete(targetDir, true);
                Directory.CreateDirectory(targetDir);

                ZipFile.ExtractToDirectory(tempZip, targetDir, true);
                RemoveZoneIdentifiers(targetDir);
                File.Delete(tempZip);
                
                progress?.Report($"✔ [IIKO] Плагин {plugin.Name} установлен");
            }
            catch (Exception ex)
            {
                progress?.Report($"⚠ [IIKO] Ошибка установки плагина {plugin.Name}: {ex.Message}");
            }
        }
    }

    private void RemoveZoneIdentifiers(string folder)
    {
        try
        {
            foreach (var file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
            {
                UnblockFile(file);
            }
        }
        catch { }
    }

    private void UnblockFile(string filePath)
    {
        try
        {
            var zonePath = filePath + ":Zone.Identifier";
            File.Delete(zonePath);
        }
        catch {  }
    }
}
