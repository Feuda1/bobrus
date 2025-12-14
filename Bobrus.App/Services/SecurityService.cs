using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed record SecurityActionResult(bool Ok, string Output);

internal sealed class SecurityService
{
    public async Task<SecurityActionResult> DisableDefenderAsync()
    {
        var command = @"
            $ErrorActionPreference='SilentlyContinue'

            # ===== ГРУБОЕ ОТКЛЮЧЕНИЕ DEFENDER И БРАНДМАУЭРА =====

            # Реестр - Defender
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender' /v DisableAntiSpyware /t REG_DWORD /d 1 /f *>$null
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender' /v DisableAntiVirus /t REG_DWORD /d 1 /f *>$null
            reg add 'HKLM\SOFTWARE\Microsoft\Windows Defender' /v DisableAntiSpyware /t REG_DWORD /d 1 /f *>$null
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection' /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f *>$null
            reg add 'HKLM\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection' /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f *>$null

            # SmartScreen
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v EnableSmartScreen /t REG_DWORD /d 0 /f *>$null
            reg add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' /v SmartScreenEnabled /t REG_SZ /d 'Off' /f *>$null

            # Скрыть Security Center
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications' /v DisableNotifications /t REG_DWORD /d 1 /f *>$null
            reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Systray' /v HideSystray /t REG_DWORD /d 1 /f *>$null
            reg delete 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' /v SecurityHealth /f *>$null

            # Службы Defender - отключить через реестр
            reg add 'HKLM\SYSTEM\CurrentControlSet\Services\WinDefend' /v Start /t REG_DWORD /d 4 /f *>$null
            reg add 'HKLM\SYSTEM\CurrentControlSet\Services\WdNisSvc' /v Start /t REG_DWORD /d 4 /f *>$null
            reg add 'HKLM\SYSTEM\CurrentControlSet\Services\SecurityHealthService' /v Start /t REG_DWORD /d 4 /f *>$null
            
            # PowerShell команды (Set-MpPreference) для надежности
            Try { Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableBehaviorMonitoring $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableBlockAtFirstSeen $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableIOAVProtection $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisablePrivacyMode $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -SignatureDisableUpdateOnStartupWithoutEngine $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableArchiveScanning $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableIntrusionPreventionSystem $true -ErrorAction SilentlyContinue } Catch {}
            Try { Set-MpPreference -DisableScriptScanning $true -ErrorAction SilentlyContinue } Catch {}
            
            Try { sc.exe stop WinDefend | Out-Null } Catch {}
            Try { sc.exe config WinDefend start= disabled | Out-Null } Catch {}
";

        var (ok, output) = await RunPowerShellAsync(command, requireElevation: true);
        return new SecurityActionResult(ok, output);
    }

    public async Task<SecurityActionResult> DisableFirewallAsync()
    {
        var command = string.Join("; ",
            "Try { Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -ErrorAction SilentlyContinue } Catch {}",
            "netsh advfirewall set allprofiles state off");

        var (ok, output) = await RunPowerShellAsync(command, requireElevation: true);
        return new SecurityActionResult(ok, output);
    }

    private static async Task<(bool ok, string output)> RunPowerShellAsync(string command, bool requireElevation = false)
    {
        string? tempLog = null;

        if (requireElevation)
        {
            tempLog = Path.Combine(Path.GetTempPath(), $"bobrus-pwsh-{Guid.NewGuid():N}.log");
            var escapedLogPath = tempLog.Replace("'", "''");
            command =
                $"$ErrorActionPreference='Continue'; $logPath='{escapedLogPath}'; $output = & {{ {command} }} *>&1; " +
                "$output | Out-File -FilePath $logPath -Encoding UTF8; " +
                "if ($LASTEXITCODE -is [int] -and $LASTEXITCODE -ne 0) { exit $LASTEXITCODE } else { exit 0 }";
        }

        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"{command}\"",
            UseShellExecute = requireElevation,
            RedirectStandardOutput = !requireElevation,
            RedirectStandardError = !requireElevation,
            CreateNoWindow = true,
            Verb = requireElevation ? "runas" : string.Empty,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            return (false, "Не удалось запустить PowerShell");
        }

        process.WaitForExit(30000);

        string output = string.Empty;

        if (requireElevation && tempLog is not null && File.Exists(tempLog))
        {
            output = await File.ReadAllTextAsync(tempLog);
            try
            {
                File.Delete(tempLog);
            }
            catch
            {
            }
        }
        else if (!requireElevation)
        {
            var stdout = await process.StandardOutput.ReadToEndAsync();
            var stderr = await process.StandardError.ReadToEndAsync();
            output = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        }

        return (process.ExitCode == 0, output);
    }
}
