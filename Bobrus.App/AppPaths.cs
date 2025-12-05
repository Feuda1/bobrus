using System;
using System.IO;

namespace Bobrus.App;

internal static class AppPaths
{
    private const string AppFolderName = "Bobrus";

    public static string AppDataRoot =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppFolderName);

    public static string LogsDirectory => Path.Combine(AppDataRoot, "logs");

    public static string UpdatesDirectory => Path.Combine(AppDataRoot, "updates");

    public static string UpdatesVersionDirectory(string version) =>
        Path.Combine(UpdatesDirectory, version);

    public static void EnsureBaseDirectories()
    {
        Directory.CreateDirectory(AppDataRoot);
        Directory.CreateDirectory(LogsDirectory);
        Directory.CreateDirectory(UpdatesDirectory);
    }
}
