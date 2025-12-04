using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed record CleanupResult(string Name, long BytesFreed);
internal sealed record CleanupProgress(string Name, bool IsStart, long BytesFreed);

internal sealed class CleaningService
{
    public async Task<IReadOnlyList<CleanupResult>> RunCleanupAsync(Action<CleanupProgress>? progress = null)
    {
        var results = new List<CleanupResult>();

        var targets = new List<Func<Task<CleanupResult>>>
        {
            () => CleanDirectory("Пользовательский TEMP", Path.GetTempPath()),
            () => CleanDirectory("Windows\\Temp", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp")),
            () => CleanDirectory("Центр обновлений (Download)", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download")),
            () => CleanDirectory("Delivery Optimization", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "DeliveryOptimization", "Cache")),
            () => CleanDirectory("WER отчёты", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft", "Windows", "WER")),
            () => CleanDirectory("Edge кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Edge\User Data\Default\Cache")),
            () => CleanDirectory("Chrome кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\User Data\Default\Cache")),
            () => CleanDirectory("Yandex кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Yandex\YandexBrowser\User Data\Default\Cache")),
            () => CleanDirectory("Opera кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Opera Software\Opera Stable\Cache")),
            () => CleanDirectory("Opera GX кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Opera Software\Opera GX Stable\Cache")),
            () => CleanDirectory("Brave кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"BraveSoftware\Brave-Browser\User Data\Default\Cache")),
            () => CleanDirectory("Vivaldi кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Vivaldi\User Data\Default\Cache")),
            () => CleanDirectory("Chromium кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Chromium\User Data\Default\Cache")),
            () => CleanFirefoxCaches(),
            () => CleanDirectory("IE/INet кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Windows\INetCache")),
            EmptyRecycleBin
        };

        foreach (var target in targets)
        {
            try
            {
                progress?.Invoke(new CleanupProgress(GetTargetName(target), true, 0));
                var result = await target();
                progress?.Invoke(new CleanupProgress(result.Name, false, result.BytesFreed));
                results.Add(result);
            }
            catch
            {
                // Игнорируем индивидуальные ошибки, продолжаем остальные шаги.
            }
        }

        return results;
    }

    private static string GetTargetName(Func<Task<CleanupResult>> target) =>
        target.Method.Name switch
        {
            nameof(EmptyRecycleBin) => "Корзина",
            _ => "Очистка"
        };

    private static Task<CleanupResult> CleanDirectory(string name, string path)
    {
        return Task.Run(() =>
        {
            var freed = 0L;

            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                return new CleanupResult(name, 0);
            }

            foreach (var file in SafeEnumFiles(path))
            {
                try
                {
                    var info = new FileInfo(file);
                    freed += info.Exists ? info.Length : 0;
                    info.IsReadOnly = false;
                    info.Delete();
                }
                catch
                {
                    // ignore
                }
            }

            foreach (var dir in SafeEnumDirectories(path))
            {
                try
                {
                    Directory.Delete(dir, recursive: true);
                }
                catch
                {
                    // ignore
                }
            }

            return new CleanupResult(name, freed);
        });
    }

    private static Task<CleanupResult> CleanFirefoxCaches()
    {
        return Task.Run(() =>
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var profilesRoot = Path.Combine(appData, "Mozilla", "Firefox", "Profiles");
            if (!Directory.Exists(profilesRoot))
            {
                return new CleanupResult("Firefox кеш", 0);
            }

            var freed = 0L;
            foreach (var profile in SafeEnumDirectories(profilesRoot))
            {
                var cache = Path.Combine(profile, "cache2");
                if (!Directory.Exists(cache))
                {
                    continue;
                }

                foreach (var file in SafeEnumFiles(cache))
                {
                    try
                    {
                        var info = new FileInfo(file);
                        freed += info.Exists ? info.Length : 0;
                        info.IsReadOnly = false;
                        info.Delete();
                    }
                    catch
                    {
                        // ignore
                    }
                }

                foreach (var dir in SafeEnumDirectories(cache))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            return new CleanupResult("Firefox кеш", freed);
        });
    }

    private static IEnumerable<string> SafeEnumFiles(string path)
    {
        try
        {
            return Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories);
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private static IEnumerable<string> SafeEnumDirectories(string path)
    {
        try
        {
            return Directory.EnumerateDirectories(path, "*", SearchOption.AllDirectories).OrderByDescending(p => p.Length);
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private static Task<CleanupResult> EmptyRecycleBin()
    {
        return Task.Run(() =>
        {
            try
            {
                SHEmptyRecycleBin(IntPtr.Zero, null, RecycleFlags.SHERB_NOCONFIRMATION | RecycleFlags.SHERB_NOPROGRESSUI | RecycleFlags.SHERB_NOSOUND);
            }
            catch
            {
                // ignore
            }

            return new CleanupResult("Корзина", 0);
        });
    }

    [Flags]
    private enum RecycleFlags : uint
    {
        SHERB_NOCONFIRMATION = 0x00000001,
        SHERB_NOPROGRESSUI = 0x00000002,
        SHERB_NOSOUND = 0x00000004
    }

    [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHEmptyRecycleBin(IntPtr hwnd, string? pszRootPath, RecycleFlags dwFlags);
}
