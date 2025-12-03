using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

internal sealed record UpdateAsset(string Name, string DownloadUrl, long? SizeBytes);

internal sealed record UpdateInfo(Version LatestVersion, UpdateAsset Asset);

internal sealed record UpdateCheckResult(bool IsUpdateAvailable, UpdateInfo? Update, string Message);

internal sealed class UpdateService
{
    private const string RepoOwner = "Feuda1";
    private const string RepoName = "bobrus";
    private const string ExecutableName = "Bobrus.exe";
    private readonly HttpClient _httpClient;

    public UpdateService(HttpClient httpClient)
    {
        _httpClient = httpClient;
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Bobrus-Updater/0.1");
        _httpClient.Timeout = TimeSpan.FromSeconds(30);

        var assemblyVersion = typeof(UpdateService).Assembly.GetName().Version;
        CurrentVersion = assemblyVersion ?? new Version(0, 1, 0);
    }

    public Version CurrentVersion { get; }

    public async Task<UpdateCheckResult> CheckForUpdatesAsync(CancellationToken cancellationToken)
    {
        var latest = await GetLatestReleaseAsync(cancellationToken);
        if (latest is null)
        {
            return new UpdateCheckResult(false, null, "Не удалось прочитать информацию о релизах.");
        }

        if (latest.LatestVersion <= CurrentVersion)
        {
            return new UpdateCheckResult(false, latest, $"Установлена актуальная версия ({CurrentVersion}).");
        }

        return new UpdateCheckResult(true, latest, $"Найдена новая версия {latest.LatestVersion}.");
    }

    public string GetPackageCachePath(Version version, string packageName)
    {
        var folder = AppPaths.UpdatesVersionDirectory(version.ToString());
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, packageName);
    }

    public async Task DownloadAssetAsync(UpdateAsset asset, string destinationPath, IProgress<double>? progress, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);

        using var response = await _httpClient.GetAsync(asset.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        response.EnsureSuccessStatusCode();

        await using var contentStream = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var fileStream = File.Create(destinationPath);

        var totalBytes = asset.SizeBytes ?? response.Content.Headers.ContentLength;
        var buffer = new byte[81920];
        long totalRead = 0;
        int read;

        while ((read = await contentStream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken)) > 0)
        {
            await fileStream.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            totalRead += read;

            if (totalBytes.HasValue && progress is not null)
            {
                progress.Report((double)totalRead / totalBytes.Value);
            }
        }
    }

    public string ExtractPackage(string archivePath, Version version)
    {
        var versionFolder = AppPaths.UpdatesVersionDirectory(version.ToString());
        var extractFolder = Path.Combine(versionFolder, "extracted");
        if (Directory.Exists(extractFolder))
        {
            Directory.Delete(extractFolder, recursive: true);
        }

        if (archivePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            ZipFile.ExtractToDirectory(archivePath, extractFolder);
            return extractFolder;
        }

        Directory.CreateDirectory(extractFolder);
        var destination = Path.Combine(extractFolder, Path.GetFileName(archivePath));
        File.Copy(archivePath, destination, overwrite: true);
        return extractFolder;
    }

    public Process? StartApplyUpdate(string extractedFolder, int currentProcessId)
    {
        var updaterPath = Path.Combine(extractedFolder, ExecutableName);
        if (!File.Exists(updaterPath))
        {
            return null;
        }

        var targetDirectory = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var arguments = BuildApplyArguments(extractedFolder, targetDirectory, currentProcessId);

        var startInfo = new ProcessStartInfo
        {
            FileName = updaterPath,
            Arguments = arguments,
            UseShellExecute = true,
            WorkingDirectory = extractedFolder
        };

        return Process.Start(startInfo);
    }

    private static string BuildApplyArguments(string sourceFolder, string targetFolder, int currentProcessId)
    {
        var builder = new StringBuilder();
        builder.Append("--apply-update ");
        builder.Append('"').Append(sourceFolder).Append('"').Append(' ');
        builder.Append('"').Append(targetFolder).Append('"').Append(' ');
        builder.Append("--launch ").Append(ExecutableName).Append(' ');
        builder.Append("--wait-process ").Append(currentProcessId);
        return builder.ToString();
    }

    private async Task<UpdateInfo?> GetLatestReleaseAsync(CancellationToken cancellationToken)
    {
        var apiUrl = $"https://api.github.com/repos/{RepoOwner}/{RepoName}/releases/latest";
        using var response = await _httpClient.GetAsync(apiUrl, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);

        var root = document.RootElement;
        var version = ParseVersion(root.GetPropertyOrDefault("tag_name"));

        UpdateAsset? asset = null;
        if (root.TryGetProperty("assets", out var assetsElement))
        {
            foreach (var assetElement in assetsElement.EnumerateArray())
            {
                var name = assetElement.GetPropertyOrDefault("name");
                var url = assetElement.GetPropertyOrDefault("browser_download_url");
                long? size = null;
                if (assetElement.TryGetProperty("size", out var sizeProp))
                {
                    size = sizeProp.GetInt64();
                }

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(url))
                {
                    continue;
                }

                if (name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    asset = new UpdateAsset(name, url, size);
                    break;
                }

                asset ??= new UpdateAsset(name, url, size);
            }
        }

        if (asset is null)
        {
            return null;
        }

        return new UpdateInfo(version, asset);
    }

    private static Version ParseVersion(string? tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
        {
            return new Version(0, 0, 0);
        }

        var cleaned = tag.Trim();
        if (cleaned.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            cleaned = cleaned[1..];
        }

        if (Version.TryParse(cleaned, out var version))
        {
            return version;
        }

        return new Version(0, 0, 0);
    }
}

internal static class JsonExtensions
{
    public static string? GetPropertyOrDefault(this JsonElement element, string name)
    {
        if (element.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String)
        {
            return property.GetString();
        }

        return null;
    }
}
