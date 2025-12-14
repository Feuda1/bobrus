using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

public record PluginInfo(string Name, string Url)
{
    public string DisplayName => Name;
}

public record PluginVersion(string Name, string Url)
{
    public string DisplayName => Name;
}

public class PluginRepository
{
    private const string PluginsBaseUrl = "https://rapid.iiko.ru/plugins/";
    private readonly HttpClient _httpClient;

    public PluginRepository(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
    }

    public async Task<List<PluginInfo>> GetPluginsAsync(CancellationToken ct = default)
    {
        var html = await _httpClient.GetStringAsync(PluginsBaseUrl, ct);
        return ParsePluginList(html).OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public async Task<List<PluginVersion>> GetVersionsAsync(string pluginUrl, CancellationToken ct = default)
    {
        var html = await _httpClient.GetStringAsync(pluginUrl, ct);
        return ParseVersions(pluginUrl, html).OrderByDescending(v => v.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static IEnumerable<PluginInfo> ParsePluginList(string html)
    {
        var regex = new Regex("href=\"(?<href>[^\"?#]+/)\"", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        var results = new List<PluginInfo>();

        foreach (Match match in regex.Matches(html))
        {
            var href = match.Groups["href"].Value;
            if (string.IsNullOrWhiteSpace(href) || href.StartsWith("../"))
            {
                continue;
            }

            var rawName = href.TrimEnd('/').Trim();
            if (rawName.Length == 0)
            {
                continue;
            }

            var decodedName = WebUtility.UrlDecode(rawName);
            var url = new Uri(new Uri(PluginsBaseUrl), href).ToString();
            results.Add(new PluginInfo(decodedName, url));
        }

        return results;
    }

    private static IEnumerable<PluginVersion> ParseVersions(string baseUrl, string html)
    {
        var regex = new Regex("href=\"(?<href>[^\"?#]+\\.zip)\"", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        var list = new List<PluginVersion>();

        foreach (Match match in regex.Matches(html))
        {
            var href = match.Groups["href"].Value;
            if (string.IsNullOrWhiteSpace(href))
            {
                continue;
            }

            var rawName = Path.GetFileName(href);
            var name = WebUtility.UrlDecode(rawName);
            var url = new Uri(new Uri(baseUrl), href).ToString();
            list.Add(new PluginVersion(name, url));
        }

        return list;
    }
}
