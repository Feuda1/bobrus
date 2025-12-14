using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using WpfApplication = System.Windows.Application;

namespace Bobrus.App;

internal sealed class OverlayDaemon
{
    private const string SettingsFileName = "settings.json";
    private FileSystemWatcher? _watcher;
    private OverlayWindow? _overlayWindow;
    private DispatcherTimer? _refreshTimer;

    public void Run()
    {
        ApplySettings();
        SetupWatcher();
        StartRefreshTimer();
    }

    private string SettingsPath => Path.Combine(AppPaths.AppDataRoot, SettingsFileName);

    private void SetupWatcher()
    {
        try
        {
            _watcher = new FileSystemWatcher(AppPaths.AppDataRoot, SettingsFileName)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
            };
            _watcher.Changed += (_, _) => OnSettingsChanged();
            _watcher.Created += (_, _) => OnSettingsChanged();
            _watcher.Renamed += (_, _) => OnSettingsChanged();
            _watcher.EnableRaisingEvents = true;
        }
        catch
        {
            WpfApplication.Current.Shutdown();
        }
    }

    private void StartRefreshTimer()
    {
        _refreshTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(5)
        };
        _refreshTimer.Tick += (_, _) => ApplySettings();
        _refreshTimer.Start();
    }

    private void OnSettingsChanged()
    {
        WpfApplication.Current.Dispatcher.InvokeAsync(async () =>
        {
            await Dispatcher.Yield(DispatcherPriority.Background);
            ApplySettings();
        });
    }

    private void ApplySettings()
    {
        var settings = LoadSettingsWithRetry();
        if (settings is null || !settings.OverlayEnabled || string.IsNullOrWhiteSpace(settings.OverlayCrm) || string.IsNullOrWhiteSpace(settings.OverlayCashDesk))
        {
            _overlayWindow?.Hide();
            return;
        }

        if (_overlayWindow == null)
        {
            _overlayWindow = new OverlayWindow();
            _overlayWindow.Closed += (_, _) => _overlayWindow = null;
        }

        var bounds = GetOverlayScreenBounds(settings.OverlayDisplay);
        _overlayWindow.UpdateOverlay(settings.OverlayCrm ?? string.Empty, settings.OverlayCashDesk ?? string.Empty, bounds);
        _overlayWindow.RefreshLayout();
        _overlayWindow.Show();
        _overlayWindow.Topmost = true;
    }

    private static AppSettings? LoadSettingsWithRetry()
    {
        var path = Path.Combine(AppPaths.AppDataRoot, SettingsFileName);
        const int maxAttempts = 3;
        for (var i = 0; i < maxAttempts; i++)
        {
            try
            {
                if (!File.Exists(path))
                {
                    Thread.Sleep(50);
                    continue;
                }

                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = new StreamReader(stream);
                var json = reader.ReadToEnd();
                var settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json);
                if (settings != null)
                {
                    return settings;
                }
            }
            catch
            {
            }

            Thread.Sleep(100);
        }

        return null;
    }

    private static Rect GetOverlayScreenBounds(string? displayName)
    {
        var screens = System.Windows.Forms.Screen.AllScreens;
        var screen = screens.FirstOrDefault(s => string.Equals(s.DeviceName, displayName, StringComparison.OrdinalIgnoreCase))
                     ?? System.Windows.Forms.Screen.PrimaryScreen
                     ?? screens.FirstOrDefault();

        if (screen is null)
        {
            return new Rect(SystemParameters.VirtualScreenLeft, SystemParameters.VirtualScreenTop, SystemParameters.VirtualScreenWidth, SystemParameters.VirtualScreenHeight);
        }

        var b = screen.Bounds;
        return new Rect(b.X, b.Y, b.Width, b.Height);
    }
}
