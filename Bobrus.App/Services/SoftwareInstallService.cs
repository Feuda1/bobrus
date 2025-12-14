using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

public enum SoftwareInstallType
{
    SilentExe,       
    RunVisible,      
    ExtractZip,      
    RevealOnly       
}

public sealed record SoftwareItem(
    string Name,
    string Url,
    SoftwareInstallType InstallType,
    string? SilentArgs = null,
    string? TargetFolder = null
);

public sealed class SoftwareInstallService
{
    private readonly HttpClient _httpClient = new();
    private readonly string _downloadDir = AppPaths.DownloadsDirectory;
    private readonly string _desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

    public static readonly SoftwareItem[] AvailableSoftware = new[]
    {
        new SoftwareItem("7-Zip", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/7z2500-x64.exe", SoftwareInstallType.SilentExe, "/S"),
        new SoftwareItem("Notepad++", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/npp.8.8.2.Installer.x64.exe", SoftwareInstallType.SilentExe, "/S"),
        new SoftwareItem("Rhelper", "https://repo.denvic.ru/remote-access/remote-access-setup.exe", SoftwareInstallType.SilentExe, "/PASSWORD=\"remote-access-setup\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-"),
        new SoftwareItem("AnyDesk", "https://download.anydesk.com/AnyDesk.exe", SoftwareInstallType.RevealOnly),
        new SoftwareItem("OrderCheck", "https://clearbat.iiko.online/downloads/OrderCheck.exe", SoftwareInstallType.RevealOnly),
        new SoftwareItem("CLEAR.bat", "https://clearbat.iiko.online/downloads/CLEAR.bat.exe", SoftwareInstallType.SilentExe, "/S"),
        new SoftwareItem("FrontTools", "https://fronttools.iiko.it/FrontTools.exe", SoftwareInstallType.RevealOnly),
        new SoftwareItem("FrontTools SQLite", "https://fronttools.iiko.it/fronttools_sqlite.zip", SoftwareInstallType.ExtractZip),
        new SoftwareItem("ComPort Checker", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/ComPortChecker.1.1.zip", SoftwareInstallType.ExtractZip),
        new SoftwareItem("Database.NET", "https://fishcodelib.com/files/DatabaseNet4.zip", SoftwareInstallType.ExtractZip),
        new SoftwareItem("Advanced IP Scanner", "https://download.advanced-ip-scanner.com/download/files/Advanced_IP_Scanner_2.5.4594.1.exe", SoftwareInstallType.SilentExe, "/S"),
        new SoftwareItem("Printer Test", "https://725920.selcdn.ru/upload_portkkm/iblock/929/9291b0d0c76c7c5756b81732df929086/Printer-TEST-V3.1C.zip", SoftwareInstallType.ExtractZip),
        new SoftwareItem("Ассистент ФН", "https://мойассистент.рф/%D1%81%D0%BA%D0%B0%D1%87%D0%B0%D1%82%D1%8C/Download/1369", SoftwareInstallType.SilentExe, "/S")
    };

    public async Task InstallAsync(
        SoftwareItem item,
        IProgress<string>? progress,
        IProgress<int>? downloadProgress,
        CancellationToken ct)
    {
        progress?.Report($"Скачивание {item.Name}...");
        
        var fileName = GetFileNameFromUrl(item.Url) ?? $"{item.Name}.bin";
        var downloadPath = Path.Combine(_downloadDir, fileName);
        Directory.CreateDirectory(_downloadDir);
        await DownloadFileAsync(item.Url, downloadPath, downloadProgress, ct);
        
        await InstallFromFileAsync(item, downloadPath, ct);
    }
    public async Task<string> DownloadOnlyAsync(SoftwareItem item, CancellationToken ct)
    {
        var fileName = GetFileNameFromUrl(item.Url) ?? $"{item.Name}.bin";
        var downloadPath = Path.Combine(_downloadDir, fileName);
        Directory.CreateDirectory(_downloadDir);
        
        await DownloadFileAsync(item.Url, downloadPath, null, ct);
        return downloadPath;
    }
    public async Task InstallFromFileAsync(SoftwareItem item, string downloadPath, CancellationToken ct)
    {
        switch (item.InstallType)
        {
            case SoftwareInstallType.SilentExe:
                await RunSilentInstallerAsync(downloadPath, item.SilentArgs ?? "/S", ct);
                CleanupInstaller(downloadPath);
                break;
            
            case SoftwareInstallType.RunVisible:
                await RunVisibleInstallerAsync(downloadPath, ct);
                CleanupInstaller(downloadPath);
                break;
                
            case SoftwareInstallType.ExtractZip:
                ExtractZipToDesktop(downloadPath, item.Name);
                break;
                
            case SoftwareInstallType.RevealOnly:
                CopyToDesktop(downloadPath, item.Name);
                break;
        }
    }

    private async Task DownloadFileAsync(string url, string destination, IProgress<int>? progress, CancellationToken ct)
    {
        using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
        response.EnsureSuccessStatusCode();
        
        var total = response.Content.Headers.ContentLength ?? -1L;
        await using var contentStream = await response.Content.ReadAsStreamAsync(ct);
        await using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
        
        var buffer = new byte[81920];
        long totalRead = 0;
        var lastPercent = 0;
        
        while (true)
        {
            var read = await contentStream.ReadAsync(buffer, ct);
            if (read == 0) break;
            
            await fileStream.WriteAsync(buffer.AsMemory(0, read), ct);
            totalRead += read;
            
            if (total > 0)
            {
                var percent = (int)Math.Clamp((totalRead * 100L) / total, 0, 100);
                if (percent != lastPercent && percent % 5 == 0)
                {
                    lastPercent = percent;
                    progress?.Report(percent);
                }
            }
        }
        progress?.Report(100);
    }

    private static async Task RunSilentInstallerAsync(string installerPath, string args, CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = installerPath,
            Arguments = args,
            UseShellExecute = true,
            CreateNoWindow = true
        };
        
        using var process = Process.Start(psi);
        if (process != null)
        {
            await process.WaitForExitAsync(ct);
        }
    }

    private static async Task RunVisibleInstallerAsync(string installerPath, CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = installerPath,
            UseShellExecute = true,
            CreateNoWindow = false
        };
        
        using var process = Process.Start(psi);
        if (process != null)
        {
            await process.WaitForExitAsync(ct);
        }
    }

    private static void RevealFile(string filePath)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{filePath}\"",
                UseShellExecute = true
            });
        }
        catch { }
    }

    private string ExtractZipToDesktop(string zipPath, string displayName)
    {
        var targetFolder = Path.Combine(_desktopDir, displayName);
        var counter = 1;
        while (Directory.Exists(targetFolder))
        {
            targetFolder = Path.Combine(_desktopDir, $"{displayName} ({counter++})");
        }
        
        Directory.CreateDirectory(targetFolder);
        ZipFile.ExtractToDirectory(zipPath, targetFolder);
        
        try { File.Delete(zipPath); } catch { }
        
        return targetFolder;
    }

    private string CopyToDesktop(string filePath, string displayName)
    {
        var ext = Path.GetExtension(filePath);
        var destPath = Path.Combine(_desktopDir, displayName + ext);
        var counter = 1;
        while (File.Exists(destPath))
        {
            destPath = Path.Combine(_desktopDir, $"{displayName} ({counter++}){ext}");
        }
        
        File.Copy(filePath, destPath, overwrite: false);
        try { File.Delete(filePath); } catch { }
        
        return destPath;
    }

    private static void CleanupInstaller(string installerPath)
    {
        try { File.Delete(installerPath); } catch { }
    }

    private static string? GetFileNameFromUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return null;
        var segment = uri.Segments.Length > 0 ? uri.Segments[^1] : null;
        return segment?.Trim('/');
    }
}
