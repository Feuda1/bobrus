using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using Serilog;

namespace Bobrus.App;

internal static class UpdateInstaller
{
    public static bool TryHandleUpdateMode(string[] args)
    {
        if (args.Length < 3)
        {
            return false;
        }

        var applyIndex = Array.IndexOf(args, "--apply-update");
        if (applyIndex < 0 || args.Length <= applyIndex + 2)
        {
            return false;
        }

        var sourceFolder = args[applyIndex + 1];
        var targetFolder = args[applyIndex + 2];

        var launchExecutable = "Bobrus.exe";
        int? waitProcessId = null;

        for (var i = applyIndex + 3; i < args.Length; i++)
        {
            var current = args[i];
            if (string.Equals(current, "--launch", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                launchExecutable = args[i + 1];
                i++;
            }
            else if (string.Equals(current, "--wait-process", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var pid))
            {
                waitProcessId = pid;
                i++;
            }
        }

        Log.Information("Запуск режима установки обновления. Source={Source}, Target={Target}, Launch={Launch}, WaitPid={WaitPid}", sourceFolder, targetFolder, launchExecutable, waitProcessId);
        try
        {
            ApplyUpdate(sourceFolder, targetFolder, launchExecutable, waitProcessId);
            Log.Information("Установка обновления завершена.");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при установке обновления.");
            System.Windows.MessageBox.Show($"Не удалось установить обновление: {ex.Message}", "Обновление", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        return true;
    }

    private static void ApplyUpdate(string sourceFolder, string targetFolder, string launchExecutableName, int? waitProcessId)
    {
        if (waitProcessId is int pid)
        {
            WaitForProcessExit(pid, TimeSpan.FromSeconds(20));
        }

        CopyDirectory(sourceFolder, targetFolder);

        var launchPath = Path.Combine(targetFolder, launchExecutableName);
        if (File.Exists(launchPath))
        {
            Log.Information("Запуск обновленной версии: {LaunchPath}", launchPath);
            Process.Start(new ProcessStartInfo
            {
                FileName = launchPath,
                WorkingDirectory = targetFolder,
                UseShellExecute = true
            });
        }
    }

    private static void CopyDirectory(string sourceFolder, string targetFolder)
    {
        Directory.CreateDirectory(targetFolder);

        foreach (var file in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
        {
            var relativePath = Path.GetRelativePath(sourceFolder, file);
            var destinationPath = Path.Combine(targetFolder, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
            File.Copy(file, destinationPath, overwrite: true);
        }
    }

    private static void WaitForProcessExit(int processId, TimeSpan timeout)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.HasExited)
            {
                process.WaitForExit((int)timeout.TotalMilliseconds);
            }
        }
        catch
        {
            // Процесс уже завершился или не найден.
        }

        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            if (!IsProcessRunning(processId))
            {
                return;
            }

            Thread.Sleep(200);
        }
    }

    private static bool IsProcessRunning(int processId)
    {
        try
        {
            var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch
        {
            return false;
        }
    }
}
