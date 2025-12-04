using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed class SecurityService
{
    public async Task<bool> DisableDefenderAsync()
    {
        var command = string.Join("; ",
            "Set-MpPreference -DisableRealtimeMonitoring $true -DisableBehaviorMonitoring $true -DisableIOAVProtection $true -DisableScriptScanning $true -SubmitSamplesConsent NeverSend -MAPSReporting 0",
            "Try { sc.exe stop WinDefend | Out-Null } Catch {}",
            "Try { sc.exe config WinDefend start= disabled | Out-Null } Catch {}");

        var (ok, _) = await RunPowerShellAsync(command);
        return ok;
    }

    public async Task<bool> DisableFirewallAsync()
    {
        var command = string.Join("; ",
            "Try { Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -ErrorAction SilentlyContinue } Catch {}",
            "netsh advfirewall set allprofiles state off");

        var (ok, _) = await RunPowerShellAsync(command);
        return ok;
    }

    private static async Task<(bool ok, string output)> RunPowerShellAsync(string command)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"{command}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            return (false, "Не удалось запустить PowerShell");
        }

        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        process.WaitForExit(20000);

        var output = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        return (process.ExitCode == 0, output);
    }
}
