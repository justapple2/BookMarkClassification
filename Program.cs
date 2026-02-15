using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using MihaZupan;
using Spectre.Console;

var options = PromptUserOptions();
var bookmarkFiles = DiscoverBookmarkFiles(options.Browsers);

if (bookmarkFiles.Count == 0)
{
    AnsiConsole.MarkupLine("[red]没有找到任何 Chrome/Edge 书签文件。[/]");
    return;
}

AnsiConsole.MarkupLine($"[green]共发现 {bookmarkFiles.Count} 个书签文件。[/]");

using var httpClient = BuildHttpClient(options.Proxy);

foreach (var bookmarkFile in bookmarkFiles)
{
    await ProcessBookmarkFileAsync(bookmarkFile, httpClient, options.TimeoutSeconds);
}

AnsiConsole.MarkupLine("[bold green]处理完成。[/]");

static UserOptions PromptUserOptions()
{
    var browsers = new List<BrowserType>();

    if (AnsiConsole.Confirm("是否扫描 Chrome 书签?", true))
    {
        browsers.Add(BrowserType.Chrome);
    }

    if (AnsiConsole.Confirm("是否扫描 Edge 书签?", true))
    {
        browsers.Add(BrowserType.Edge);
    }

    if (browsers.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]未选择浏览器，默认同时扫描 Chrome 和 Edge。[/]");
        browsers.AddRange([BrowserType.Chrome, BrowserType.Edge]);
    }

    var timeoutSeconds = AnsiConsole.Ask<int>("访问超时秒数(建议 5~15)", 8);

    var useProxy = AnsiConsole.Confirm("是否使用代理?", false);
    string? proxy = null;

    if (useProxy)
    {
        proxy = AnsiConsole.Ask<string>(
            "请输入代理地址（示例: [grey]http://127.0.0.1:7890[/] 或 [grey]socks5://127.0.0.1:1080[/]）");
    }

    return new UserOptions(browsers, timeoutSeconds <= 0 ? 8 : timeoutSeconds, proxy);
}

static HttpClient BuildHttpClient(string? proxyText)
{
    var handler = new HttpClientHandler
    {
        AllowAutoRedirect = true,
        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
        ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
    };

    if (!string.IsNullOrWhiteSpace(proxyText))
    {
        if (!Uri.TryCreate(proxyText, UriKind.Absolute, out var proxyUri))
        {
            AnsiConsole.MarkupLine($"[yellow]代理地址无效，忽略代理: {proxyText}[/]");
        }
        else if (proxyUri.Scheme.StartsWith("socks", StringComparison.OrdinalIgnoreCase))
        {
            handler.UseProxy = true;
            handler.Proxy = new HttpToSocks5Proxy(proxyUri.Host, proxyUri.Port);
            AnsiConsole.MarkupLine($"[green]已启用 SOCKS 代理: {proxyUri}[/]");
        }
        else
        {
            handler.UseProxy = true;
            handler.Proxy = new WebProxy(proxyUri);
            AnsiConsole.MarkupLine($"[green]已启用 HTTP 代理: {proxyUri}[/]");
        }
    }

    var client = new HttpClient(handler);
    client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("BookMarkClassification", "1.0"));
    client.Timeout = TimeSpan.FromSeconds(30);
    return client;
}

static List<BookmarkFile> DiscoverBookmarkFiles(IReadOnlyCollection<BrowserType> browsers)
{
    var result = new List<BookmarkFile>();
    var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

    foreach (var browser in browsers)
    {
        var basePath = browser switch
        {
            BrowserType.Chrome => Path.Combine(localAppData, "Google", "Chrome", "User Data"),
            BrowserType.Edge => Path.Combine(localAppData, "Microsoft", "Edge", "User Data"),
            _ => string.Empty
        };

        if (string.IsNullOrWhiteSpace(basePath) || !Directory.Exists(basePath))
        {
            continue;
        }

        foreach (var dir in Directory.EnumerateDirectories(basePath))
        {
            var name = Path.GetFileName(dir);
            if (!name.Equals("Default", StringComparison.OrdinalIgnoreCase)
                && !name.StartsWith("Profile", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var bookmarkPath = Path.Combine(dir, "Bookmarks");
            if (File.Exists(bookmarkPath))
            {
                result.Add(new BookmarkFile(browser, name, bookmarkPath));
            }
        }
    }

    return result;
}

static async Task ProcessBookmarkFileAsync(BookmarkFile bookmarkFile, HttpClient httpClient, int timeoutSeconds)
{
    AnsiConsole.WriteLine();
    AnsiConsole.MarkupLine($"[bold cyan]处理 {bookmarkFile.Browser} - {bookmarkFile.ProfileName}[/]");

    var text = await File.ReadAllTextAsync(bookmarkFile.Path);
    var json = JsonNode.Parse(text)?.AsObject();

    if (json is null)
    {
        AnsiConsole.MarkupLine("[red]书签文件解析失败，已跳过。[/]");
        return;
    }

    var roots = json["roots"]?.AsObject();
    if (roots is null)
    {
        AnsiConsole.MarkupLine("[red]未找到 roots 节点，已跳过。[/]");
        return;
    }

    var deadByRoot = new Dictionary<string, List<DeadBookmark>>();
    var checkedCount = 0;
    var deadCount = 0;

    foreach (var rootName in new[] { "bookmark_bar", "other", "synced" })
    {
        if (roots[rootName]?.AsObject() is not JsonObject rootNode)
        {
            continue;
        }

        var deadList = new List<DeadBookmark>();
        await TraverseFolderAsync(rootNode, [], deadList, httpClient, timeoutSeconds, () => checkedCount++, () => deadCount++);
        if (deadList.Count > 0)
        {
            deadByRoot[rootName] = deadList;
        }
    }

    if (deadCount == 0)
    {
        AnsiConsole.MarkupLine($"[green]未发现失效书签（检测 {checkedCount} 条 URL）。[/]");
        return;
    }

    var folderName = $"Unresponsive_Bookmarks_{DateTime.Now:yyyyMMdd_HHmmss}";

    foreach (var pair in deadByRoot)
    {
        if (roots[pair.Key]?.AsObject() is not JsonObject rootNode)
        {
            continue;
        }

        var rootChildren = rootNode["children"]?.AsArray();
        if (rootChildren is null)
        {
            continue;
        }

        var container = CreateFolder(folderName);
        foreach (var dead in pair.Value)
        {
            AppendWithPath(container, dead.PathFolders, dead.UrlNode);
        }

        rootChildren.Add(container);
    }

    var backupPath = bookmarkFile.Path + ".bak_" + DateTime.Now.ToString("yyyyMMddHHmmss");
    File.Copy(bookmarkFile.Path, backupPath, overwrite: false);

    var jsonText = json.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    await File.WriteAllTextAsync(bookmarkFile.Path, jsonText);

    AnsiConsole.MarkupLine($"[yellow]检测完成: 共检测 {checkedCount} 条 URL，失效 {deadCount} 条。[/]");
    AnsiConsole.MarkupLine($"[green]已写回书签并备份到: {backupPath}[/]");
}

static async Task TraverseFolderAsync(
    JsonObject folder,
    List<string> currentPath,
    List<DeadBookmark> deadList,
    HttpClient httpClient,
    int timeoutSeconds,
    Action checkedCounter,
    Action deadCounter)
{
    var children = folder["children"]?.AsArray();
    if (children is null)
    {
        return;
    }

    for (var i = children.Count - 1; i >= 0; i--)
    {
        if (children[i] is not JsonObject child)
        {
            continue;
        }

        var type = child["type"]?.GetValue<string>();

        if (string.Equals(type, "url", StringComparison.OrdinalIgnoreCase))
        {
            var url = child["url"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(url))
            {
                continue;
            }

            checkedCounter();
            var ok = await IsAliveAsync(httpClient, url, timeoutSeconds);
            if (!ok)
            {
                var deepClone = child.DeepClone().AsObject();
                deadList.Add(new DeadBookmark([.. currentPath], deepClone));
                children.RemoveAt(i);
                deadCounter();
            }
        }
        else if (string.Equals(type, "folder", StringComparison.OrdinalIgnoreCase))
        {
            var name = child["name"]?.GetValue<string>() ?? "未命名目录";
            currentPath.Add(name);
            await TraverseFolderAsync(child, currentPath, deadList, httpClient, timeoutSeconds, checkedCounter, deadCounter);
            currentPath.RemoveAt(currentPath.Count - 1);
        }
    }
}

static async Task<bool> IsAliveAsync(HttpClient httpClient, string url, int timeoutSeconds)
{
    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));

    try
    {
        using var headRequest = new HttpRequestMessage(HttpMethod.Head, url);
        using var headResponse = await httpClient.SendAsync(headRequest, HttpCompletionOption.ResponseHeadersRead, cts.Token);

        if ((int)headResponse.StatusCode is >= 200 and < 400)
        {
            return true;
        }

        if (headResponse.StatusCode == HttpStatusCode.MethodNotAllowed)
        {
            using var getRequest = new HttpRequestMessage(HttpMethod.Get, url);
            using var getResponse = await httpClient.SendAsync(getRequest, HttpCompletionOption.ResponseHeadersRead, cts.Token);
            return (int)getResponse.StatusCode is >= 200 and < 400;
        }

        return false;
    }
    catch
    {
        return false;
    }
}

static JsonObject CreateFolder(string folderName)
{
    return new JsonObject
    {
        ["type"] = "folder",
        ["name"] = folderName,
        ["date_added"] = ToChromeTimestamp(DateTime.UtcNow),
        ["date_modified"] = ToChromeTimestamp(DateTime.UtcNow),
        ["children"] = new JsonArray()
    };
}

static void AppendWithPath(JsonObject rootFolder, IReadOnlyList<string> pathFolders, JsonObject urlNode)
{
    var cursor = rootFolder;

    foreach (var folderName in pathFolders)
    {
        var children = cursor["children"]?.AsArray();
        if (children is null)
        {
            children = new JsonArray();
            cursor["children"] = children;
        }

        JsonObject? existing = null;
        foreach (var child in children)
        {
            if (child is not JsonObject childObj)
            {
                continue;
            }

            if (!string.Equals(childObj["type"]?.GetValue<string>(), "folder", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(childObj["name"]?.GetValue<string>(), folderName, StringComparison.Ordinal))
            {
                existing = childObj;
                break;
            }
        }

        if (existing is null)
        {
            existing = CreateFolder(folderName);
            children.Add(existing);
        }

        cursor = existing;
    }

    var finalChildren = cursor["children"]?.AsArray();
    if (finalChildren is null)
    {
        finalChildren = new JsonArray();
        cursor["children"] = finalChildren;
    }

    finalChildren.Add(urlNode);
}

static string ToChromeTimestamp(DateTime utc)
{
    var epoch = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    var microseconds = (long)(utc - epoch).TotalMilliseconds * 1000;
    return microseconds.ToString();
}

enum BrowserType
{
    Chrome,
    Edge
}

sealed record UserOptions(IReadOnlyCollection<BrowserType> Browsers, int TimeoutSeconds, string? Proxy);
sealed record BookmarkFile(BrowserType Browser, string ProfileName, string Path);
sealed record DeadBookmark(IReadOnlyList<string> PathFolders, JsonObject UrlNode);
