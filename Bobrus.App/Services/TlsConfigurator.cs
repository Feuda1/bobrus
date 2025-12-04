using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Bobrus.App.Services;

internal sealed record TlsConfigResult(bool Ok, string Message);

internal sealed class TlsConfigurator
{
    private static readonly string[] ProtocolPaths =
    {
        @"SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client",
        @"SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"
    };

    public Task<TlsConfigResult> EnableTls12Async()
    {
        return Task.Run(() =>
        {
            var cmd = string.Join("; ",
                "New-Item -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Client' -Force",
                "New-Item -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Server' -Force",
                "Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Client' -Name 'Enabled' -Type DWord -Value 1",
                "Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Client' -Name 'DisabledByDefault' -Type DWord -Value 0",
                "Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Server' -Name 'Enabled' -Type DWord -Value 1",
                "Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\SecurityProviders\\SCHANNEL\\Protocols\\TLS 1.2\\Server' -Name 'DisabledByDefault' -Type DWord -Value 0",
                "Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\.NETFramework\\v4.0.30319' -Name 'SchUseStrongCrypto' -Type DWord -Value 1",
                "Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\.NETFramework\\v4.0.30319' -Name 'SchUseStrongCrypto' -Type DWord -Value 1",
                "Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\.NETFramework\\v4.0.30319' -Name 'SystemDefaultTlsVersions' -Type DWord -Value 1",
                "Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\.NETFramework\\v4.0.30319' -Name 'SystemDefaultTlsVersions' -Type DWord -Value 1");

            var result = RunElevatedPowerShell(cmd);
            return result;
        });
    }

    private static TlsConfigResult RunElevatedPowerShell(string command)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"{command}\"",
                Verb = "runas",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                return new TlsConfigResult(false, "Не удалось запустить PowerShell");
            }

            process.WaitForExit(30000);
            return process.ExitCode == 0
                ? new TlsConfigResult(true, "Настройки TLS 1.2 применены")
                : new TlsConfigResult(false, $"Код выхода: {process.ExitCode}");
        }
        catch (Exception ex)
        {
            return new TlsConfigResult(false, ex.Message);
        }
    }
}
