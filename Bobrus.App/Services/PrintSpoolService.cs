using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed class PrintSpoolService
{
    private const string SpoolerServiceName = "Spooler";

    public async Task RestartAndCleanAsync(Action<string>? log = null)
    {
        log?.Invoke("Остановка диспетчера печати");
        await StopServiceAsync();

        log?.Invoke("Очистка очереди печати");
        CleanSpoolDirectory();

        log?.Invoke("Запуск диспетчера печати");
        await StartServiceAsync();
    }

    private static Task StopServiceAsync()
    {
        return Task.Run(() =>
        {
            try
            {
                using var controller = new ServiceController(SpoolerServiceName);
                if (controller.Status != ServiceControllerStatus.Stopped &&
                    controller.Status != ServiceControllerStatus.StopPending)
                {
                    controller.Stop();
                    controller.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
                }
            }
            catch
            {
                // ignore
            }
        });
    }

    private static Task StartServiceAsync()
    {
        return Task.Run(() =>
        {
            try
            {
                using var controller = new ServiceController(SpoolerServiceName);
                if (controller.Status != ServiceControllerStatus.Running &&
                    controller.Status != ServiceControllerStatus.StartPending)
                {
                    controller.Start();
                    controller.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                }
            }
            catch
            {
                // ignore
            }
        });
    }

    private static void CleanSpoolDirectory()
    {
        try
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "spool", "PRINTERS");
            if (!Directory.Exists(path))
            {
                return;
            }

            foreach (var file in Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var info = new FileInfo(file);
                    info.IsReadOnly = false;
                    info.Delete();
                }
                catch
                {
                    // ignore
                }
            }
        }
        catch
        {
            // ignore
        }
    }
}
