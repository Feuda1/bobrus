using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed class RhelperSetupService
{
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromMinutes(10) };
    
    private const string DownloadUrl = "https://repo.denvic.ru/remote-access/remote-access-setup.exe";
    private const string InstallerPassword = "remote-access-setup";
    private const string RhelperExePath = @"C:\Program Files (x86)\Rhelper39\RHurs_v39.exe";
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
    public async Task WaitForDefenderDisabledAsync(IProgress<string>? progress, SetupFlowController controller)
    {
        var ct = controller.Token;
        progress?.Report("[Rhelper] Открываю настройки защитника...");
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
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "ms-settings:windowsdefender",
                    UseShellExecute = true
                });
            }
            catch { }
        }
        await Task.Delay(2000, ct);
        
        var instructions = new[]
        {
            "",
            "!!! ВНИМАНИЕ: ТРЕБУЕТСЯ ДЕЙСТВИЕ !!!",
            "Отключите ВСЕ галочки в открывшемся окне:",
            "1. Защита в реальном времени (!) ВАЖНО",
            "2. Облачная защита",
            "3. Автоматическая отправка образцов",
            "4. Защита от подделки",
            "",
            "--> ЗАКРОЙТЕ ОКНО ЗАЩИТНИКА ДЛЯ ПРОДОЛЖЕНИЯ <--",
            ""
        };

        foreach (var line in instructions)
        {
            progress?.Report(line);
        }
        
        string[] processNames = { "SecHealthUI", "WindowsSecurityHealthUI" };
        string[] windowTitles = { "Безопасность Windows", "Windows Security" };
        while (HasVisibleWindow(processNames, windowTitles))
        {
            await controller.WaitIfPausedAsync(); 
            
            ct.ThrowIfCancellationRequested();
            await Task.Delay(500, ct);
        }
        
        progress?.Report("[Rhelper] ✔ Окно закрыто.");
    }
    

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

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
            _ = GetWindowThreadProcessId(hwnd, out var pid);
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

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    public async Task<string?> DownloadAsync(IProgress<string>? progress, CancellationToken ct)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), "remote-access-setup.exe");
        
        try
        {
            progress?.Report("[Rhelper] Скачивание установщика...");
            
            using var response = await _http.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead, ct);
            response.EnsureSuccessStatusCode();
            
            var totalBytes = response.Content.Headers.ContentLength ?? 0;
            await using var contentStream = await response.Content.ReadAsStreamAsync(ct);
            await using var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true);
            
            var buffer = new byte[8192];
            long downloadedBytes = 0;
            int bytesRead;
            
            while ((bytesRead = await contentStream.ReadAsync(buffer, ct)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), ct);
                downloadedBytes += bytesRead;
                
                if (totalBytes > 0)
                {
                    var percent = (int)(downloadedBytes * 100 / totalBytes);
                    progress?.Report($"[PROGRESS]{percent}");
                }
            }
            
            progress?.Report("[Rhelper] ✔ Установщик скачан");
            return tempPath;
        }
        catch (Exception ex)
        {
            progress?.Report($"[Rhelper] ✖ Ошибка скачивания: {ex.Message}");
            return null;
        }
    }
    public Process? StartInstaller(string installerPath, IProgress<string>? progress)
    {
        try
        {
            progress?.Report("[Rhelper] Запуск установщика...");
            
            var psi = new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = $"/PASSWORD=\"{InstallerPassword}\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-",
                UseShellExecute = true
            };
            
            var process = Process.Start(psi);
            
            if (process == null)
            {
                progress?.Report("[Rhelper] ✖ Не удалось запустить установщик");
                return null;
            }
            
            progress?.Report("[Rhelper] Установка Rhelper...");
            return process;
        }
        catch (Exception ex)
        {
            progress?.Report($"[Rhelper] ✖ Ошибка запуска: {ex.Message}");
            return null;
        }
    }
    public async Task WaitForShowIdClosedAsync(IProgress<string>? progress, SetupFlowController controller)
    {
        var ct = controller.Token;
        progress?.Report("[Rhelper] Ожидание окна ID...");
        Process? showIdProcess = null;
        while (showIdProcess == null)
        {
            await controller.WaitIfPausedAsync();
            ct.ThrowIfCancellationRequested();
            
            var processes = Process.GetProcessesByName("show-id");
            if (processes.Length > 0)
            {
                showIdProcess = processes[0];
                break;
            }
            
            await Task.Delay(500, ct);
        }
        
        progress?.Report("[Rhelper] Запишите ID и нажмите кнопку...");
        while (!showIdProcess.HasExited)
        {
            await controller.WaitIfPausedAsync();
            ct.ThrowIfCancellationRequested();
            await Task.Delay(500, ct);
        }
        
        progress?.Report("[Rhelper] ✔ ID записан");
    }
    public static bool IsInstalled() => File.Exists(RhelperExePath);
    public static string ExePath => RhelperExePath;
    
    private static IntPtr FindWindowByPartialTitle(string partialTitle)
    {
        IntPtr result = IntPtr.Zero;
        
        EnumWindows((hwnd, lParam) =>
        {
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            var title = sb.ToString();
            
            if (title.Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
            {
                result = hwnd;
                return false;
            }
            
            return true;
        }, IntPtr.Zero);
        
        return result;
    }
}
