using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed class TouchDeviceManager
{
    private static readonly string[] TouchKeywords = new[]
    {
        "touch screen",
        "touchscreen",
        "touch digitizer",
        "hid-compliant touch",
        "touch panel"
    };

    public async Task<IReadOnlyList<TouchDevice>> GetTouchDevicesAsync()
    {
        return await Task.Run(() =>
        {
            var devices = new List<TouchDevice>();
            using var searcher = new ManagementObjectSearcher("SELECT Name, PNPDeviceID, ConfigManagerErrorCode FROM Win32_PnPEntity");
            foreach (var obj in searcher.Get().Cast<ManagementObject>())
            {
                var name = (obj["Name"] as string) ?? string.Empty;
                var id = (obj["PNPDeviceID"] as string) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                if (!IsTouchName(name))
                {
                    continue;
                }

                var errorCode = obj["ConfigManagerErrorCode"] is int code ? code : 0;
                var isEnabled = errorCode != 22; // 22 = disabled
                devices.Add(new TouchDevice(name, id, isEnabled));
            }

            return (IReadOnlyList<TouchDevice>)devices;
        });
    }

    public async Task<bool> SetTouchEnabledAsync(bool enable)
    {
        var devices = await GetTouchDevicesAsync();
        if (devices.Count == 0)
        {
            return false;
        }

        foreach (var device in devices)
        {
            await RunPnpUtilAsync(enable, device.InstanceId);
        }

        return true;
    }

    public async Task<bool> RestartTouchAsync()
    {
        var devices = await GetTouchDevicesAsync();
        if (devices.Count == 0)
        {
            return false;
        }

        foreach (var device in devices)
        {
            await RunPnpUtilAsync(false, device.InstanceId);
        }

        await Task.Delay(500);

        foreach (var device in devices)
        {
            await RunPnpUtilAsync(true, device.InstanceId);
        }

        return true;
    }

    private static bool IsTouchName(string name)
    {
        var lower = name.ToLowerInvariant();
        return TouchKeywords.Any(k => lower.Contains(k));
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

internal sealed record TouchDevice(string Name, string InstanceId, bool IsEnabled);
