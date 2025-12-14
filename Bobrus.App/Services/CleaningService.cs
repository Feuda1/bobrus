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

        var targets = new List<(string Name, Func<Task<CleanupResult>> Task)>
        {
            ("Пользовательский TEMP", () => CleanDirectory("Пользовательский TEMP", Path.GetTempPath())),
            ("Windows\\Temp", () => CleanDirectory("Windows\\Temp", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp"))),
            ("Центр обновлений (Download)", () => CleanDirectory("Центр обновлений (Download)", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download"))),
            ("Delivery Optimization", () => CleanDirectory("Delivery Optimization", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "DeliveryOptimization", "Cache"))),
            ("WER отчёты", () => CleanDirectory("WER отчёты", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft", "Windows", "WER"))),
            ("Edge кеш", () => CleanDirectory("Edge кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Edge\User Data\Default\Cache"))),
            ("Chrome кеш", () => CleanDirectory("Chrome кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\User Data\Default\Cache"))),
            ("Yandex кеш", () => CleanDirectory("Yandex кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Yandex\YandexBrowser\User Data\Default\Cache"))),
            ("Opera кеш", () => CleanDirectory("Opera кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Opera Software\Opera Stable\Cache"))),
            ("Opera GX кеш", () => CleanDirectory("Opera GX кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Opera Software\Opera GX Stable\Cache"))),
            ("Brave кеш", () => CleanDirectory("Brave кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"BraveSoftware\Brave-Browser\User Data\Default\Cache"))),
            ("Vivaldi кеш", () => CleanDirectory("Vivaldi кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Vivaldi\User Data\Default\Cache"))),
            ("Chromium кеш", () => CleanDirectory("Chromium кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Chromium\User Data\Default\Cache"))),
            ("Firefox кеш", CleanFirefoxCaches),
            ("IE/INet кеш", () => CleanDirectory("IE/INet кеш", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Windows\INetCache"))),
            ("Логи iiko CashServer (старше 30 дней)", () => CleanOldLogs("Логи iiko CashServer (старше 30 дней)", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"iiko\CashServer\Logs"), TimeSpan.FromDays(30))),
            ("Корзина", EmptyRecycleBin)
        };

        foreach (var target in targets)
        {
            try
            {
                progress?.Invoke(new CleanupProgress(target.Name, true, 0));
                var result = await target.Task();
                progress?.Invoke(new CleanupProgress(result.Name, false, result.BytesFreed));
                results.Add(result);
            }
            catch
            {
            }
        }

        return results;
    }

    private static Task<CleanupResult> CleanDirectory(string name, string path)
    {
        return Task.Run(() =>
        {
            var freed = 0L;

            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                return new CleanupResult(name, 0);
            }

            freed += CleanDirectoryRecursive(path);

            return new CleanupResult(name, freed);
        });
    }

    private static long CleanDirectoryRecursive(string path)
    {
        long freed = 0;

        IEnumerable<string> files;
        try
        {
            files = Directory.EnumerateFiles(path, "*", SearchOption.TopDirectoryOnly);
        }
        catch
        {
            return 0;
        }

        foreach (var file in files)
        {
            try
            {
                var info = new FileInfo(file);
                if (!info.Exists) continue;
                
                var size = info.Length;
                info.IsReadOnly = false;
                info.Delete();
                freed += size;
            }
            catch
            {
            }
        }

        IEnumerable<string> dirs;
        try
        {
            dirs = Directory.EnumerateDirectories(path, "*", SearchOption.TopDirectoryOnly);
        }
        catch
        {
            return freed;
        }

        foreach (var dir in dirs)
        {
            freed += CleanDirectoryRecursive(dir);
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch
            {
            }
        }

        return freed;
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
            IEnumerable<string> profiles;
            try
            {
                profiles = Directory.EnumerateDirectories(profilesRoot, "*", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                profiles = Enumerable.Empty<string>();
            }

            foreach (var profile in profiles)
            {
                var cache = Path.Combine(profile, "cache2");
                if (!Directory.Exists(cache))
                {
                    continue;
                }

                IEnumerable<string> files;
                try
                {
                    files = Directory.EnumerateFiles(cache, "*", SearchOption.AllDirectories);
                }
                catch
                {
                    files = Enumerable.Empty<string>();
                }

                foreach (var file in files)
                {
                    try
                    {
                        var info = new FileInfo(file);
                        if (!info.Exists) continue;

                        var size = info.Length;
                        info.IsReadOnly = false;
                        info.Delete();
                        freed += size;
                    }
                    catch
                    {
                    }
                }

                IEnumerable<string> dirs;
                try
                {
                    dirs = Directory.EnumerateDirectories(cache, "*", SearchOption.AllDirectories).OrderByDescending(p => p.Length);
                }
                catch
                {
                    dirs = Enumerable.Empty<string>();
                }

                foreach (var dir in dirs)
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                    }
                    catch
                    {
                    }
                }
            }

            return new CleanupResult("Firefox кеш", freed);
        });
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
            }

            return new CleanupResult("Корзина", 0);
        });
    }

    private static Task<CleanupResult> CleanOldLogs(string name, string path, TimeSpan maxAge)
    {
        return Task.Run(() =>
        {
            var freed = 0L;
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                return new CleanupResult(name, 0);
            }

            var threshold = DateTime.UtcNow - maxAge;
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories);
            }
            catch
            {
                files = Enumerable.Empty<string>();
            }

            foreach (var file in files)
            {
                try
                {
                    var info = new FileInfo(file);
                    if (!info.Exists || info.LastWriteTimeUtc > threshold)
                    {
                        continue;
                    }

                    var size = info.Length;
                    info.IsReadOnly = false;
                    info.Delete();
                    freed += size;
                }
                catch
                {
                }
            }

            return new CleanupResult(name, freed);
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
