using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed record ComPortDevice(string Name, string InstanceId, bool IsEnabled);

internal sealed class ComPortManager
{
    private static readonly string[] PortKeywords = new[]
    {
        "(com",
        "serial",
        "rs232"
    };

    public async Task<IReadOnlyList<ComPortDevice>> GetPortsAsync()
    {
        return await Task.Run(() =>
        {
            var devices = new List<ComPortDevice>();
            using var searcher = new ManagementObjectSearcher("SELECT Name, PNPDeviceID, ConfigManagerErrorCode, PNPClass FROM Win32_PnPEntity");
            foreach (var obj in searcher.Get().Cast<ManagementObject>())
            {
                var name = (obj["Name"] as string) ?? string.Empty;
                var id = (obj["PNPDeviceID"] as string) ?? string.Empty;
                var pnpClass = (obj["PNPClass"] as string) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                if (!IsPortName(name) && !pnpClass.Equals("Ports", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var errorCode = obj["ConfigManagerErrorCode"] is int code ? code : 0;
                var isEnabled = errorCode != 22; // 22 = disabled
                devices.Add(new ComPortDevice(name, id, isEnabled));
            }

            return (IReadOnlyList<ComPortDevice>)devices;
        });
    }

    public async Task<bool> RestartPortsAsync()
    {
        var devices = await GetPortsAsync();
        if (devices.Count == 0)
        {
            return false;
        }

        foreach (var device in devices)
        {
            await RunPnpUtilAsync(false, device.InstanceId);
        }

        await Task.Delay(400);

        foreach (var device in devices)
        {
            await RunPnpUtilAsync(true, device.InstanceId);
        }

        return true;
    }

    private static bool IsPortName(string name)
    {
        var lower = name.ToLowerInvariant();
        return PortKeywords.Any(k => lower.Contains(k));
    }

    private static Task RunPnpUtilAsync(bool enable, string instanceId)
    {
        return Task.Run(() =>
        {
            var args = enable ? $"/enable-device \"{instanceId}\"" : $"/disable-device \"{instanceId}\"";
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            process?.WaitForExit(12000);
        });
    }
}
