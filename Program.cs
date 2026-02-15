using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
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

// 构建两个 HttpClient：一个带代理（如果用户指定），一个直连（无代理）
// 两个都创建以便后续“代理/不代理各试一次”的逻辑。
using var clientWithProxy = BuildHttpClient(options.Proxy);
using var clientDirect = BuildHttpClient(null);

foreach (var bookmarkFile in bookmarkFiles)
{
    await ProcessBookmarkFileAsync(bookmarkFile, clientWithProxy, clientDirect, options.TimeoutSeconds, !string.IsNullOrWhiteSpace(options.Proxy));
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
        browsers.AddRange(new[] { BrowserType.Chrome, BrowserType.Edge });
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

static async Task ProcessBookmarkFileAsync(
    BookmarkFile bookmarkFile,
    HttpClient clientWithProxy,
    HttpClient clientDirect,
    int timeoutSeconds,
    bool hasProxyConfigured)
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
        await TraverseFolderAsync(rootNode, new List<string>(), deadList, clientWithProxy, clientDirect, timeoutSeconds, hasProxyConfigured, () => checkedCount++, () => deadCount++);
        if (deadList.Count > 0)
        {
            deadByRoot[rootName] = deadList;
        }
    }

    AnsiConsole.MarkupLine($"[yellow]检测完成: 共检测 {checkedCount} 条 URL，失效 {deadCount} 条。[/]");

    // 导出到 HTML，保留可访问链接的目录结构，并把所有不可访问的链接放到顶层“不可访问”目录下（按 root / 原路径分组）
    var exportPath = ExportBookmarksToHtml(bookmarkFile, roots, deadByRoot);
    AnsiConsole.MarkupLine($"[green]已导出 HTML 书签到: {exportPath}[/]");
}
    
static async Task TraverseFolderAsync(
    JsonObject folder,
    List<string> currentPath,
    List<DeadBookmark> deadList,
    HttpClient clientWithProxy,
    HttpClient clientDirect,
    int timeoutSeconds,
    bool hasProxy,
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

            var ok = await TryIsAliveEitherAsync(clientWithProxy, clientDirect, hasProxy, url, timeoutSeconds);
            if (!ok)
            {
                var deepClone = child.DeepClone().AsObject();
                deadList.Add(new DeadBookmark(currentPath.ToArray(), deepClone));
                deadCounter();
            }
        }
        else if (string.Equals(type, "folder", StringComparison.OrdinalIgnoreCase))
        {
            var name = child["name"]?.GetValue<string>() ?? "未命名目录";
            currentPath.Add(name);
            await TraverseFolderAsync(child, currentPath, deadList, clientWithProxy, clientDirect, timeoutSeconds, hasProxy, checkedCounter, deadCounter);
            currentPath.RemoveAt(currentPath.Count - 1);
        }
    }
}

static async Task<bool> TryIsAliveEitherAsync(HttpClient clientWithProxy, HttpClient clientDirect, bool hasProxy, string url, int timeoutSeconds)
{
    // 如果配置了代理，先尝试带代理，再尝试直连；任意一次成功即视为可访问。
    if (hasProxy)
    {
        try
        {
            if (await IsAliveAsync(clientWithProxy, url, timeoutSeconds))
            {
                return true;
            }
        }
        catch
        {
            // 忽略，继续尝试直连
        }
    }

    // 直连尝试（或者代理尝试失败后尝试直连）
    try
    {
        if (await IsAliveAsync(clientDirect, url, timeoutSeconds))
        {
            return true;
        }
    }
    catch
    {
        // 最终失败
    }

    return false;
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

static long ChromeTimestampToUnixSeconds(string? chromeTimestamp)
{
    // Chrome timestamp stored as microseconds since 1601-01-01 UTC (as string)
    if (string.IsNullOrWhiteSpace(chromeTimestamp) || !long.TryParse(chromeTimestamp, out var micro))
    {
        return DateTimeOffset.UtcNow.ToUnixTimeSeconds();
    }

    var epoch = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    try
    {
        var ms = micro / 1000; // to milliseconds
        var dt = epoch.AddMilliseconds(ms);
        return new DateTimeOffset(dt).ToUnixTimeSeconds();
    }
    catch
    {
        return DateTimeOffset.UtcNow.ToUnixTimeSeconds();
    }
}

static string ExportBookmarksToHtml(BookmarkFile bookmarkFile, JsonObject roots, Dictionary<string, List<DeadBookmark>> deadByRoot)
{
    var myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    var safeProfile = string.Concat(bookmarkFile.ProfileName.Split(Path.GetInvalidFileNameChars()));
    var fileName = $"Bookmarks_Export_{bookmarkFile.Browser}_{safeProfile}_{DateTime.Now:yyyyMMdd_HHmmss}.html";
    var fullPath = Path.Combine(myDocs, fileName);

    // 收集所有不可访问的 URL（按字符串）以便在导出原目录时排除
    var deadUrlSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var list in deadByRoot.Values)
    {
        foreach (var dead in list)
        {
            var u = dead.UrlNode["url"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(u))
            {
                deadUrlSet.Add(u);
            }
        }
    }

    var sb = new StringBuilder();
    sb.AppendLine("<!DOCTYPE NETSCAPE-Bookmark-file-1>");
    sb.AppendLine("<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; charset=UTF-8\">");
    sb.AppendLine("<TITLE>Bookmarks</TITLE>");
    sb.AppendLine("<H1>Bookmarks</H1>");
    sb.AppendLine("<DL><p>");

    var exportFolderName = $"Exported_Bookmarks_{DateTime.Now:yyyyMMdd_HHmmss}";

    // 顶层容器
    sb.AppendLine($"  <DT><H3>{EscapeHtml(exportFolderName)}</H3>");
    sb.AppendLine("  <DL><p>");

    // 先写出各 root 的可访问项（保留目录结构），跳过 deadUrlSet 中的 URL
    foreach (var rootName in new[] { "bookmark_bar", "other", "synced" })
    {
        if (roots[rootName]?.AsObject() is not JsonObject rootNode)
        {
            continue;
        }

        sb.AppendLine($"    <DT><H3>{EscapeHtml(rootName)}</H3>");
        sb.AppendLine("    <DL><p>");

        // 递归写入结构
        WriteFolderHtml(rootNode, sb, deadUrlSet);

        sb.AppendLine("    </DL><p>");
    }

    // 追加“不可访问”顶层目录，并按 root -> path 分组放入
    sb.AppendLine($"    <DT><H3>不可访问</H3>");
    sb.AppendLine("    <DL><p>");

    foreach (var kv in deadByRoot)
    {
        var rootName = kv.Key;
        var list = kv.Value;

        sb.AppendLine($"      <DT><H3>{EscapeHtml(rootName)}</H3>");
        sb.AppendLine("      <DL><p>");

        // group by path (preserve original path)
        var groups = list.GroupBy(d => string.Join(" / ", d.PathFolders ?? Array.Empty<string>()), StringComparer.Ordinal);

        foreach (var group in groups)
        {
            var folderLabel = string.IsNullOrEmpty(group.Key) ? "根目录" : group.Key;
            sb.AppendLine($"        <DT><H3>{EscapeHtml(folderLabel)}</H3>");
            sb.AppendLine("        <DL><p>");

            foreach (var dead in group)
            {
                var url = dead.UrlNode["url"]?.GetValue<string>() ?? string.Empty;
                var title = dead.UrlNode["name"]?.GetValue<string>() ?? url;
                var addDate = ChromeTimestampToUnixSeconds(dead.UrlNode["date_added"]?.GetValue<string>());
                sb.AppendLine($"          <DT><A HREF=\"{EscapeHtml(url)}\" ADD_DATE=\"{addDate}\">{EscapeHtml(title)}</A>");
            }

            sb.AppendLine("        </DL><p>");
        }

        sb.AppendLine("      </DL><p>");
    }

    sb.AppendLine("    </DL><p>"); // 结束 不可访问
    sb.AppendLine("  </DL><p>"); // 结束 导出根
    sb.AppendLine("</DL><p>");

    File.WriteAllText(fullPath, sb.ToString(), Encoding.UTF8);
    return fullPath;
}

static void WriteFolderHtml(JsonObject folderNode, StringBuilder sb, HashSet<string> deadUrlSet)
{
    // 如果 folderNode 本身是一个 folder 对象，写 name，然后处理 children
    var nameAttr = folderNode["name"]?.GetValue<string>();
    if (!string.IsNullOrEmpty(nameAttr))
    {
        // 这里 folder 的标题写在调用处已经写入，所以此方法只负责 children
    }

    var children = folderNode["children"]?.AsArray();
    if (children is null)
    {
        return;
    }

    foreach (var child in children)
    {
        if (child is not JsonObject childObj)
        {
            continue;
        }

        var type = childObj["type"]?.GetValue<string>();
        if (string.Equals(type, "folder", StringComparison.OrdinalIgnoreCase))
        {
            var fname = childObj["name"]?.GetValue<string>() ?? "未命名目录";
            sb.AppendLine($"      <DT><H3>{EscapeHtml(fname)}</H3>");
            sb.AppendLine("      <DL><p>");
            WriteFolderHtml(childObj, sb, deadUrlSet);
            sb.AppendLine("      </DL><p>");
        }
        else if (string.Equals(type, "url", StringComparison.OrdinalIgnoreCase))
        {
            var url = childObj["url"]?.GetValue<string>() ?? string.Empty;
            if (deadUrlSet.Contains(url))
            {
                // 跳过不可访问的 URL（它们会被收集到“不可访问”目录）
                continue;
            }

            var title = childObj["name"]?.GetValue<string>() ?? url;
            var addDate = ChromeTimestampToUnixSeconds(childObj["date_added"]?.GetValue<string>());
            var icon = childObj["icon"]?.GetValue<string>();
            var iconAttr = string.IsNullOrEmpty(icon) ? "" : $" ICON=\"{EscapeHtml(icon)}\"";
            sb.AppendLine($"        <DT><A HREF=\"{EscapeHtml(url)}\" ADD_DATE=\"{addDate}\"{iconAttr}>{EscapeHtml(title)}</A>");
        }
    }
}

static string EscapeHtml(string? s)
{
    if (string.IsNullOrEmpty(s)) return string.Empty;
    return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}

enum BrowserType
{
    Chrome,
    Edge
}

sealed record UserOptions(IReadOnlyCollection<BrowserType> Browsers, int TimeoutSeconds, string? Proxy);
sealed record BookmarkFile(BrowserType Browser, string ProfileName, string Path);
sealed record DeadBookmark(IReadOnlyList<string> PathFolders, JsonObject UrlNode);
