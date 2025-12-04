using System;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Bobrus.App.Services;

internal sealed class TlsConfigurator
{
    private static readonly string[] ProtocolPaths =
    {
        @"SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client",
        @"SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"
    };

    public Task<bool> EnableTls12Async()
    {
        return Task.Run(() =>
        {
            try
            {
                foreach (var path in ProtocolPaths)
                {
                    using var key = Registry.LocalMachine.CreateSubKey(path, writable: true);
                    if (key is null)
                    {
                        return false;
                    }
                    key.SetValue("Enabled", 1, RegistryValueKind.DWord);
                    key.SetValue("DisabledByDefault", 0, RegistryValueKind.DWord);
                }

                EnableDotNetStrongCrypto();
                return true;
            }
            catch
            {
                return false;
            }
        });
    }

    private static void EnableDotNetStrongCrypto()
    {
        SetDword(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\.NETFramework\v4.0.30319", "SchUseStrongCrypto", 1);
        SetDword(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319", "SchUseStrongCrypto", 1);
        SetDword(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\.NETFramework\v4.0.30319", "SystemDefaultTlsVersions", 1);
        SetDword(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319", "SystemDefaultTlsVersions", 1);
    }

    private static void SetDword(RegistryHive hive, string path, string name, int value)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using var key = baseKey.CreateSubKey(path, writable: true);
        key?.SetValue(name, value, RegistryValueKind.DWord);
    }
}
