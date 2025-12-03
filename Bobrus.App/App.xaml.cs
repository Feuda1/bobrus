using System;
using System.IO;
using System.Windows;
using Serilog;

namespace Bobrus.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        AppPaths.EnsureBaseDirectories();
        ConfigureLogging();

        if (UpdateInstaller.TryHandleUpdateMode(e.Args))
        {
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Log.CloseAndFlush();
        base.OnExit(e);
    }

    private static void ConfigureLogging()
    {
        var logPath = Path.Combine(AppPaths.LogsDirectory, "bobrus-.log");
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(logPath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                rollOnFileSizeLimit: true,
                shared: true)
            .CreateLogger();

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            Log.Fatal(args.ExceptionObject as Exception, "Необработанная ошибка приложения");
            Log.CloseAndFlush();
        };
    }
}
