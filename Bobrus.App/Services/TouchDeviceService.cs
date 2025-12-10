using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace Bobrus.App.Services;

internal sealed class TouchDeviceService
{
    private static readonly ILogger Logger = Log.ForContext<TouchDeviceService>();

    public async Task<IReadOnlyList<TouchDevice>> GetTouchDevicesAsync()
    {
        Logger.Information("GetTouchDevicesAsync: starting");

        var script = @"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Get-PnpDevice -Class 'HIDClass' -ErrorAction SilentlyContinue | Where-Object {
    $_.FriendlyName -like '*сенсор*' -or $_.FriendlyName -like '*touch*'
} | ForEach-Object {
    $s = if ($_.Status -eq 'OK') { 'E' } else { 'D' }
    Write-Output ""$($_.InstanceId)|$($_.FriendlyName)|$s""
}
";
        var output = await RunPowerShellScriptAsync(script);
        var devices = ParseOutput(output);

        Logger.Information("GetTouchDevicesAsync: found {Count} devices", devices.Count);
        return devices;
    }

    public async Task<bool> SetTouchEnabledAsync(bool enable)
    {
        var cmd = enable ? "Enable-PnpDevice" : "Disable-PnpDevice";
        Logger.Information("SetTouchEnabledAsync: {Cmd}", cmd);

        // Один скрипт для всех устройств сразу
        var script = $@"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$devices = Get-PnpDevice -Class 'HIDClass' -ErrorAction SilentlyContinue | Where-Object {{
    $_.FriendlyName -like '*сенсор*' -or $_.FriendlyName -like '*touch*'
}}
if ($devices) {{
    $devices | {cmd} -Confirm:$false -ErrorAction SilentlyContinue
    Write-Output ""OK:$($devices.Count)""
}} else {{
    Write-Output ""NONE""
}}
";
        var output = await RunPowerShellScriptAsync(script);
        Logger.Information("SetTouchEnabledAsync: result={Result}", output);

        return output?.StartsWith("OK") == true;
    }

    public async Task<bool> RestartTouchAsync()
    {
        Logger.Information("RestartTouchAsync: starting");

        // Всё в одном скрипте: отключить, подождать, включить
        var script = @"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$devices = Get-PnpDevice -Class 'HIDClass' -ErrorAction SilentlyContinue | Where-Object {
    $_.FriendlyName -like '*сенсор*' -or $_.FriendlyName -like '*touch*'
}
if ($devices) {
    $devices | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    $devices | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Write-Output ""OK:$($devices.Count)""
} else {
    Write-Output ""NONE""
}
";
        var output = await RunPowerShellScriptAsync(script);
        Logger.Information("RestartTouchAsync: result={Result}", output);

        return output?.StartsWith("OK") == true;
    }

    private static List<TouchDevice> ParseOutput(string output)
    {
        var devices = new List<TouchDevice>();
        if (string.IsNullOrWhiteSpace(output))
            return devices;

        foreach (var line in output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            var p = line.Split('|');
            if (p.Length >= 3 && !string.IsNullOrEmpty(p[0]))
            {
                devices.Add(new TouchDevice(p[1].Trim(), p[0].Trim(), p[2].Trim() == "E"));
            }
        }

        return devices;
    }

    private static async Task<string> RunPowerShellScriptAsync(string script)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), "bobrus_touch.ps1");
        try
        {
            await File.WriteAllTextAsync(tempFile, script, new UTF8Encoding(true));

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoLogo -NoProfile -ExecutionPolicy Bypass -File \"{tempFile}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using var p = Process.Start(psi);
            if (p == null)
            {
                Logger.Error("RunPowerShellScriptAsync: Process.Start returned null");
                return string.Empty;
            }

            var output = await p.StandardOutput.ReadToEndAsync();
            var error = await p.StandardError.ReadToEndAsync();
            await p.WaitForExitAsync();

            if (!string.IsNullOrWhiteSpace(error))
            {
                Logger.Warning("RunPowerShellScriptAsync: stderr={Error}", error.Trim());
            }

            return output.Trim();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "RunPowerShellScriptAsync: exception");
            return string.Empty;
        }
        finally
        {
            try { File.Delete(tempFile); } catch { }
        }
    }
}

internal sealed record TouchDevice(string FriendlyName, string InstanceId, bool IsEnabled);
