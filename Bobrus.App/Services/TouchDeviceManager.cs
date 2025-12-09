using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed class TouchDeviceManager
{
    private static readonly string[] TouchNameKeywords = new[]
    {
        "touch screen",
        "touchscreen",
        "touch screen device",
        "touchscreen device",
        "touch digitizer",
        "hid-compliant touch",
        "hid compliant touch",
        "hid-compliant digitizer",
        "hid compliant digitizer",
        "hid-compliant pen",
        "touch panel",
        "touch-panel",
        "multi-touch",
        "multitouch",
        "digitizer",
        "pen and touch",
        "touch input",
        "сенсор",
        "сенсорный",
        "сенсорная панель",
        "сенсорный экран",
        "hid-совместимый сенсорный экран",
        "hid-совместимый сенсор",
        "hid compliant touch screen"
    };

    private static readonly string[] TouchHardwareKeywords = new[]
    {
        "touch",
        "digitizer",
        "multitouch",
        "hid_touch",
        "hid\\elan",
        "elan",
        "wacom",
        "silead",
        "goodix",
        "melfas",
        "maxtouch",
        "atmel",
        "synaptics",
        "stmicro",
        "ilitek",
        "raydium",
        "sensel",
        "egalax",
        "egalax_touch",
        "touchpanel",
        "touch-panel",
        "hid_sensor_touch",
        "hid\\vid_04f3", // Elan
        "hid\\vid_056a", // Wacom
        "hid\\vid_1fd2", // Goodix
        "hid\\vid_27c6", // Goodix
        "hid\\vid_222a", // Silead
        "hid\\vid_0eef", // Elan older
        "hid\\vid_0488", // ITE / Raydium / Synaptics touch chips
        "hid\\vid_0b05", // ASUS touch variants
        "hid\\vid_0bda", // Realtek touch firmware on some panels
        "hid\\vid_03eb", // Atmel/Microchip
        "hid\\vid_328f", // Atmel touch variants
        "hid\\vid_258a", // Egalax / GoodTouch
        "hid\\vid_04b4"  // Cypress
    };

    public async Task<IReadOnlyList<TouchDevice>> GetTouchDevicesAsync()
    {
        return await Task.Run(() =>
        {
            var devices = new List<TouchDevice>();
            using var searcher = new ManagementObjectSearcher("SELECT Name, Description, Manufacturer, PNPDeviceID, ConfigManagerErrorCode, HardwareID, CompatibleID, PNPClass, ClassGuid FROM Win32_PnPEntity");
            foreach (var obj in searcher.Get().Cast<ManagementObject>())
            {
                var name = (obj["Name"] as string) ?? string.Empty;
                var id = (obj["PNPDeviceID"] as string) ?? string.Empty;
                var pnpClass = (obj["PNPClass"] as string) ?? string.Empty;
                var classGuid = (obj["ClassGuid"] as string) ?? string.Empty;
                var description = (obj["Description"] as string) ?? string.Empty;
                var manufacturer = (obj["Manufacturer"] as string) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                var hardwareIds = obj["HardwareID"] as string[] ?? Array.Empty<string>();
                var compatibleIds = obj["CompatibleID"] as string[] ?? Array.Empty<string>();

                if (!IsLikelyTouch(name, description, manufacturer, id, pnpClass, classGuid, hardwareIds, compatibleIds))
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

    private static bool IsLikelyTouch(
        string name,
        string description,
        string manufacturer,
        string pnpDeviceId,
        string pnpClass,
        string classGuid,
        IEnumerable<string> hardwareIds,
        IEnumerable<string> compatibleIds)
    {
        var lowerName = name.ToLowerInvariant();
        var lowerDescription = description.ToLowerInvariant();
        var lowerManufacturer = manufacturer.ToLowerInvariant();
        var lowerId = pnpDeviceId.ToLowerInvariant();
        var isHid = lowerId.StartsWith("hid\\") || string.Equals(pnpClass, "hidclass", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(classGuid, "{745a17a0-74d3-11d0-b6fe-00a0c90f57da}", StringComparison.OrdinalIgnoreCase);

        if (!isHid)
        {
            return false;
        }

        if (TouchNameKeywords.Any(k => lowerName.Contains(k) || lowerDescription.Contains(k)))
        {
            return true;
        }

        if (TouchNameKeywords.Any(k => lowerManufacturer.Contains(k)))
        {
            return true;
        }

        foreach (var hw in hardwareIds)
        {
            var hwLower = (hw ?? string.Empty).ToLowerInvariant();
            if (TouchHardwareKeywords.Any(k => hwLower.Contains(k)))
            {
                return true;
            }
        }

        foreach (var cid in compatibleIds)
        {
            var cidLower = (cid ?? string.Empty).ToLowerInvariant();
            if (TouchHardwareKeywords.Any(k => cidLower.Contains(k)))
            {
                return true;
            }
        }

        if (lowerName.Contains("hid-compliant device") || lowerName.Contains("hid compliant device"))
        {
            if (lowerName.Contains("mouse") || lowerName.Contains("keyboard") || lowerName.Contains("keybd"))
                return false;

            if (TouchHardwareKeywords.Any(k => lowerId.Contains(k)))
            {
                return true;
            }
        }

        return false;
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
