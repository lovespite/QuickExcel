using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>HTTP is performed by the host after an exact-host grant, never by the sandboxed worker.</summary>
public static class ScriptNetworkBroker
{
    internal const int MaxBodyBytes = 8 * 1024 * 1024;
    internal const int MaxHeaderBytes = 32 * 1024;
    private const int MaxRedirects = 10;
    private static readonly HttpClient Client = new(new SocketsHttpHandler
    {
        AllowAutoRedirect = false,
        UseCookies = false,
        UseProxy = false,
        Credentials = null,
        PreAuthenticate = false,
        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli,
        MaxResponseHeadersLength = MaxHeaderBytes / 1024,
        PooledConnectionLifetime = TimeSpan.FromMinutes(2),
        ConnectTimeout = Timeout.InfiniteTimeSpan
    }) { Timeout = Timeout.InfiniteTimeSpan };

    public static string NormalizeHost(string host)
    {
        if (string.IsNullOrWhiteSpace(host) || host.Length > 253 || host != host.Trim() ||
            host.Any(c => char.IsControl(c) || char.IsWhiteSpace(c) || c is '/' or '\\' or '@' or '?' or '#' or '%' or '*'))
            throw new InvalidOperationException("网络权限必须是完整域名或 IP 地址，不允许 URL、通配符或端口。");
        if (host.StartsWith('[') && host.EndsWith(']')) host = host[1..^1];
        if (IPAddress.TryParse(host, out var address))
            return (address.IsIPv4MappedToIPv6 ? address.MapToIPv4() : address).ToString();
        if (host.Contains(':')) throw new InvalidOperationException("网络权限不能包含端口。");
        if (host.EndsWith('.')) host = host[..^1];
        try { host = new IdnMapping { UseStd3AsciiRules = true }.GetAscii(host).ToLowerInvariant(); }
        catch (ArgumentException) { throw new InvalidOperationException("网络权限包含无效域名。"); }
        if (host.Length is 0 or > 253 || host.Split('.').Any(label => label.Length is 0 or > 63 ||
            label.StartsWith('-') || label.EndsWith('-') || label.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-')))
            throw new InvalidOperationException("网络权限包含无效域名。");
        return host;
    }

    public static string[] NormalizeHosts(IReadOnlyList<string> hosts)
    {
        ArgumentNullException.ThrowIfNull(hosts);
        if (hosts.Count > 128) throw new InvalidOperationException("最多声明 128 个网络地址。");
        return hosts.Select(NormalizeHost).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
    }

    internal static JsonElement Fetch(JsonElement call, IReadOnlyList<string> authorizedHosts, CancellationToken cancellation,
        Action<ScriptNetworkEntry>? record = null, ScriptNetworkOptions? options = null) => FetchAsync(call, authorizedHosts, cancellation, record, options).GetAwaiter().GetResult();

    internal static async Task<JsonElement> FetchAsync(JsonElement call, IReadOnlyList<string> authorizedHosts,
        CancellationToken cancellation, Action<ScriptNetworkEntry>? record = null, ScriptNetworkOptions? options = null)
    {
        var started = DateTimeOffset.UtcNow;
        var timer = Stopwatch.StartNew();
        var originalUrl = ""; var method = "GET"; var responseUrl = ""; var statusText = "";
        int? status = null;
        List<KeyValuePair<string, string>> headers = [], responseHeaders = [];
        IReadOnlyList<KeyValuePair<string, string>> recordedHeaders = [];
        byte[]? body = null, responseBody = null;
        byte[]? recordedBody = null;
        var recordedMethod = "GET";
        string? failure = null;
        try
        {
            cancellation.ThrowIfCancellationRequested();
            options ??= new();
            options.Validate();
            if (!options.AllowNetwork) throw new InvalidOperationException("设置中已禁止脚本访问网络。");
            var bodyLimit = options.BodyLimitBytes;
            if (call.ValueKind != JsonValueKind.Object) throw new InvalidOperationException("无效的 fetch 参数。");
            originalUrl = String(call, "url", null) ?? throw new InvalidOperationException("fetch 缺少 URL。");
            if (originalUrl.Length > 8192) throw new InvalidOperationException("fetch URL 不能超过 8192 字符。");
            var allowed = NormalizeHosts(authorizedHosts).ToHashSet(StringComparer.Ordinal);
            var uri = Address(originalUrl, allowed);
            responseUrl = uri.AbsoluteUri;
            method = String(call, "method", "GET")!;
            if (method.Length is 0 or > 32 || !method.All(TokenCharacter)) throw new InvalidOperationException("无效的 HTTP 方法。");
            if (method.ToUpperInvariant() is "CONNECT" or "TRACE" or "TRACK") throw new InvalidOperationException("fetch 不支持此 HTTP 方法。");
            if (method.ToUpperInvariant() is "DELETE" or "GET" or "HEAD" or "OPTIONS" or "POST" or "PUT") method = method.ToUpperInvariant();
            recordedMethod = method;
            var redirect = String(call, "redirect", "follow");
            if (redirect is not ("follow" or "error" or "manual")) throw new InvalidOperationException("无效的 fetch redirect 选项。");
            var credentials = String(call, "credentials", "same-origin");
            if (credentials is not ("omit" or "same-origin")) throw new InvalidOperationException("脚本 fetch 不提供浏览器 Cookie 或系统凭据；请使用 credentials: 'omit'。");
            headers = ReadHeaders(call);
            var encoded = String(call, "bodyBase64", null);
            if (encoded is not null)
            {
                if (encoded.Length > (bodyLimit + 2) / 3 * 4) throw new InvalidOperationException($"fetch 请求体超过 {options.BodyLimitMiB} MiB 限制。");
                try { body = Convert.FromBase64String(encoded.Replace('-', '+').Replace('_', '/')); }
                catch (FormatException) { throw new InvalidOperationException("fetch 请求体编码无效。"); }
                if (body.Length > bodyLimit) throw new InvalidOperationException($"fetch 请求体超过 {options.BodyLimitMiB} MiB 限制。");
                if (method is "GET" or "HEAD") throw new InvalidOperationException("GET 和 HEAD 请求不能包含请求体。");
            }
            recordedHeaders = headers.ToArray(); recordedBody = body;
            var redirected = false;
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            timeout.CancelAfter(options.ActionTimeout);
            for (var hop = 0; ; hop++)
            {
                cancellation.ThrowIfCancellationRequested();
                using var request = new HttpRequestMessage(new HttpMethod(method), uri);
                if (body is not null) request.Content = new ByteArrayContent(body);
                foreach (var header in headers)
                    if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value))
                    {
                        request.Content ??= new ByteArrayContent([]);
                        if (!request.Content.Headers.TryAddWithoutValidation(header.Key, header.Value)) throw new InvalidOperationException("无效的 HTTP 请求头。");
                    }
                using var response = await Client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(false);
                status = (int)response.StatusCode; statusText = response.ReasonPhrase ?? "";
                responseUrl = uri.AbsoluteUri;
                responseHeaders = response.Headers.Concat(response.Content.Headers)
                    .Where(h => !h.Key.Equals("Set-Cookie", StringComparison.OrdinalIgnoreCase) && !h.Key.Equals("Set-Cookie2", StringComparison.OrdinalIgnoreCase))
                    .SelectMany(h => h.Value.Select(v => new KeyValuePair<string, string>(h.Key.ToLowerInvariant(), v))).ToList();
                if (responseHeaders.Count > 128 || responseHeaders.Sum(h => Encoding.UTF8.GetByteCount(h.Key) + Encoding.UTF8.GetByteCount(h.Value)) > MaxHeaderBytes)
                    throw new InvalidOperationException("fetch 响应头超过限制。");
                if (status is 301 or 302 or 303 or 307 or 308)
                {
                    if (redirect == "error") throw new InvalidOperationException("fetch 禁止重定向。");
                    if (redirect == "manual")
                    {
                        responseBody = [];
                        return Result("", 0, "", [], [], false, "opaqueredirect");
                    }
                    if (response.Headers.Location is { } location)
                    {
                        if (hop >= MaxRedirects) throw new InvalidOperationException("fetch 重定向超过 10 次限制。");
                        var next = Address(new Uri(uri, location).AbsoluteUri, allowed);
                        if (uri.Scheme == "https" && next.Scheme == "http") throw new InvalidOperationException("fetch 禁止从 HTTPS 重定向到 HTTP。");
                        if (uri.Scheme != next.Scheme || uri.IdnHost != next.IdnHost || uri.Port != next.Port)
                            headers.RemoveAll(h => h.Key.Equals("authorization", StringComparison.OrdinalIgnoreCase));
                        if (status == 303 && method is not ("GET" or "HEAD") || status is 301 or 302 && method == "POST")
                        {
                            method = "GET"; body = null;
                            headers.RemoveAll(h => h.Key.StartsWith("content-", StringComparison.OrdinalIgnoreCase));
                        }
                        uri = next; redirected = true; continue;
                    }
                }
                if (method != "HEAD" && response.Content.Headers.ContentLength > bodyLimit) throw new InvalidOperationException($"fetch 响应体超过 {options.BodyLimitMiB} MiB 限制。");
                using var stream = await response.Content.ReadAsStreamAsync(timeout.Token).ConfigureAwait(false);
                using var buffer = new MemoryStream();
                var chunk = new byte[16 * 1024];
                int read;
                while ((read = await stream.ReadAsync(chunk, timeout.Token).ConfigureAwait(false)) != 0)
                {
                    if (buffer.Length + read > bodyLimit) throw new InvalidOperationException($"fetch 响应体超过 {options.BodyLimitMiB} MiB 限制。");
                    // MemoryStream normally doubles capacity. Cap that growth for
                    // non-power-of-two preferences such as 3 or 10 MiB as well.
                    if (buffer.Capacity < buffer.Length + read)
                        buffer.Capacity = (int)Math.Min(bodyLimit, Math.Max((long)buffer.Capacity * 2, buffer.Length + read));
                    buffer.Write(chunk, 0, read);
                }
                responseBody = buffer.ToArray();
                return Result(responseUrl, status.Value, statusText, responseHeaders, responseBody, redirected, "basic");
            }
        }
        catch (OperationCanceledException error) when (!cancellation.IsCancellationRequested)
        {
            failure = "网络请求超时。";
            throw new HttpRequestException(failure, error);
        }
        catch (Exception error)
        {
            failure = error is OperationCanceledException ? "网络请求已取消或超时。" : error.Message;
            throw;
        }
        finally
        {
            record?.Invoke(new ScriptNetworkEntry
            {
                Started = started, Url = Limit(originalUrl, 8192), Method = Limit(recordedMethod, 32), ResponseUrl = Limit(responseUrl, 8192),
                Status = status, StatusText = Limit(statusText, 256), RequestHeaders = recordedHeaders, ResponseHeaders = responseHeaders,
                RequestBody = recordedBody, ResponseBody = responseBody, Duration = timer.Elapsed, Error = failure is null ? null : Limit(failure, 1024)
            });
        }
    }

    private static Uri Address(string url, HashSet<string> allowed)
    {
        if (url.Length > 8192 || !Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https") ||
            !string.IsNullOrEmpty(uri.UserInfo) || url.Any(char.IsControl))
            throw new InvalidOperationException("fetch 只接受不含用户名、密码的绝对 HTTP/HTTPS URL。");
        var host = NormalizeHost(uri.IdnHost);
        if (!allowed.Contains(host)) throw new InvalidOperationException("未授权访问网络地址：" + host);
        return string.IsNullOrEmpty(uri.Fragment) ? uri : new UriBuilder(uri) { Fragment = "" }.Uri;
    }

    private static List<KeyValuePair<string, string>> ReadHeaders(JsonElement call)
    {
        var result = new List<KeyValuePair<string, string>>();
        if (!call.TryGetProperty("headers", out var headers)) return result;
        if (headers.ValueKind != JsonValueKind.Array || headers.GetArrayLength() > 128) throw new InvalidOperationException("无效的 fetch 请求头。");
        var bytes = 0;
        foreach (var header in headers.EnumerateArray())
        {
            if (header.ValueKind != JsonValueKind.Array || header.GetArrayLength() != 2 ||
                header[0].ValueKind != JsonValueKind.String || header[1].ValueKind != JsonValueKind.String) throw new InvalidOperationException("无效的 fetch 请求头。");
            var name = header[0].GetString()!.ToLowerInvariant(); var value = header[1].GetString()!;
            if (name.Length is 0 or > 256 || !name.All(TokenCharacter) || value.Any(c => c < 32 && c != '\t' || c == 127) ||
                (bytes += Encoding.UTF8.GetByteCount(name) + Encoding.UTF8.GetByteCount(value)) > MaxHeaderBytes)
                throw new InvalidOperationException("无效的 fetch 请求头或超过 32 KiB 限制。");
            if (name is "accept-charset" or "accept-encoding" or "access-control-request-headers" or "access-control-request-method" or
                "connection" or "content-length" or "cookie" or "cookie2" or "date" or "dnt" or "expect" or "host" or "keep-alive" or
                "origin" or "permissions-policy" or "referer" or "te" or "trailer" or "transfer-encoding" or "upgrade" or "via" ||
                name.StartsWith("proxy-", StringComparison.Ordinal) || name.StartsWith("sec-", StringComparison.Ordinal))
                throw new InvalidOperationException("fetch 不允许设置请求头：" + name);
            result.Add(new(name, value));
        }
        return result;
    }

    private static bool TokenCharacter(char c) => char.IsAsciiLetterOrDigit(c) || "!#$%&'*+-.^_`|~".Contains(c);
    private static string? String(JsonElement call, string name, string? fallback) => !call.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null
        ? fallback : value.ValueKind == JsonValueKind.String ? value.GetString() : throw new InvalidOperationException("无效的 fetch 参数：" + name);
    private static string Limit(string value, int length) => value.Length <= length ? value : value[..length];
    private static JsonElement Result(string url, int status, string statusText, IReadOnlyList<KeyValuePair<string, string>> headers,
        byte[] body, bool redirected, string type)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject(); writer.WriteString("url", url); writer.WriteNumber("status", status); writer.WriteString("statusText", statusText);
            writer.WriteBoolean("redirected", redirected); writer.WriteString("type", type); writer.WriteStartArray("headers");
            foreach (var header in headers) { writer.WriteStartArray(); writer.WriteStringValue(header.Key); writer.WriteStringValue(header.Value); writer.WriteEndArray(); }
            // JSON's default encoder expands '+' sixfold. Fetch uses base64url
            // inside its private wire payload so an 8 MiB body fits the 16 MiB frame.
            writer.WriteEndArray(); writer.WriteString("bodyBase64", Convert.ToBase64String(body).Replace('+', '-').Replace('/', '_')); writer.WriteEndObject();
        }
        using var document = JsonDocument.Parse(buffer.ToArray());
        return document.RootElement.Clone();
    }
}
