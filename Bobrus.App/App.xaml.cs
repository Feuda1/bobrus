using System;
using System.IO;
using System.Windows;
using System.Threading;
using System.Text;
using Application = System.Windows.Application;
using Bobrus.App.Services;
using Serilog;
using System.Runtime.InteropServices;

namespace Bobrus.App;

public partial class App : Application
{
    private Mutex? _singleInstanceMutex;
    private const string MainMutexName = "Global\\BobrusSingleInstance";
    private const string OverlayMutexName = "Global\\BobrusOverlayRunner";
    private bool _isOverlayRunner;

    protected override void OnStartup(StartupEventArgs e)
    {
        _isOverlayRunner = e.Args.Any(a => string.Equals(a, "--overlay-runner", StringComparison.OrdinalIgnoreCase));
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var mutexName = _isOverlayRunner ? OverlayMutexName : MainMutexName;
        _singleInstanceMutex = new Mutex(initiallyOwned: true, name: mutexName, createdNew: out var isFirstInstance);
        if (!_isOverlayRunner && !isFirstInstance)
        {
            TryShowExistingWindow();
            Shutdown();
            return;
        }

        if (_isOverlayRunner && !isFirstInstance)
        {
            Shutdown();
            return;
        }

        AppPaths.EnsureBaseDirectories();
        ConfigureLogging();

        if (UpdateInstaller.TryHandleUpdateMode(e.Args))
        {
            Shutdown();
            return;
        }

        if (_isOverlayRunner)
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var daemon = new OverlayDaemon();
            daemon.Run();
            base.OnStartup(e);
            return;
        }

        ShutdownMode = ShutdownMode.OnMainWindowClose;
        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();
        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        Log.CloseAndFlush();
        base.OnExit(e);
    }
    private static void ConfigureLogging()
    {
        var logPath = Path.Combine(AppPaths.LogsDirectory, "bobrus-.log");
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Verbose()
            .WriteTo.File(logPath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                rollOnFileSizeLimit: true,
                shared: true)
            .WriteTo.Sink(new UiLogSink(), restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Information)
            .CreateLogger();

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            Log.Fatal(args.ExceptionObject as Exception, "Необработанная ошибка приложения");
            Log.CloseAndFlush();
        };
    }

    private void TryShowExistingWindow()
    {
        try
        {
            var hwnd = FindWindowByTitle("Bobrus");
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SW_RESTORE);
                SetForegroundWindow(hwnd);
            }
        }
        catch
        {
        }
    }

    private static IntPtr FindWindowByTitle(string title)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, lParam) =>
        {
            var length = GetWindowTextLength(hWnd);
            if (length == 0) return true;
            Span<char> buffer = stackalloc char[length + 1];
            _ = GetWindowText(hWnd, buffer, buffer.Length);
            var currentTitle = new string(buffer[..Math.Max(0, buffer.Length - 1)]);
            if (string.Equals(currentTitle, title, StringComparison.OrdinalIgnoreCase))
            {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private const int SW_RESTORE = 9;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, Span<char> lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
