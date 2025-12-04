using System;
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
            try
            {
                foreach (var path in ProtocolPaths)
                {
                    using var key = Registry.LocalMachine.CreateSubKey(path, writable: true);
                    if (key is null)
                    {
                        return new TlsConfigResult(false, $"Не удалось открыть ветку реестра: {path}");
                    }
                    key.SetValue("Enabled", 1, RegistryValueKind.DWord);
                    key.SetValue("DisabledByDefault", 0, RegistryValueKind.DWord);
                }

                var strong = EnableDotNetStrongCrypto();
                if (!strong.Ok)
                {
                    return strong;
                }

                return new TlsConfigResult(true, "Настройки TLS 1.2 применены");
            }
            catch (Exception ex)
            {
                return new TlsConfigResult(false, ex.Message);
            }
        });
    }

    private static TlsConfigResult EnableDotNetStrongCrypto()
    {
        try
        {
            SetDword(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\.NETFramework\v4.0.30319", "SchUseStrongCrypto", 1);
            SetDword(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319", "SchUseStrongCrypto", 1);
            SetDword(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\.NETFramework\v4.0.30319", "SystemDefaultTlsVersions", 1);
            SetDword(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319", "SystemDefaultTlsVersions", 1);
            return new TlsConfigResult(true, "StrongCrypto включен");
        }
        catch (Exception ex)
        {
            return new TlsConfigResult(false, $"StrongCrypto: {ex.Message}");
        }
    }

    private static void SetDword(RegistryHive hive, string path, string name, int value)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using var key = baseKey.CreateSubKey(path, writable: true);
        key?.SetValue(name, value, RegistryValueKind.DWord);
    }
}
