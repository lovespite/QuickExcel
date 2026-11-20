using System.Collections.Concurrent;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Runtime;

namespace QuickExcel.Core.Test;

public sealed class ScriptNetworkBrokerTests
{
    [Theory]
    [InlineData("EXAMPLE.com.", "example.com")]
    [InlineData("例子.中国", "xn--fsqu00a.xn--fiqs8s")]
    [InlineData("[2001:DB8::1]", "2001:db8::1")]
    [InlineData("::ffff:127.0.0.1", "127.0.0.1")]
    public void HostGrantsAreCanonicalAndExact(string input, string expected) => Assert.Equal(expected, ScriptNetworkBroker.NormalizeHost(input));

    [Theory]
    [InlineData("*.example.com")]
    [InlineData("https://example.com")]
    [InlineData("example.com:443")]
    [InlineData("example.com/path")]
    [InlineData("user@example.com")]
    [InlineData("example.com..")]
    [InlineData(" example.com")]
    [InlineData("fe80::1%4")]
    public void MalformedDeclarationsAreRejected(string host) => Assert.Throws<InvalidOperationException>(() => ScriptNetworkBroker.NormalizeHost(host));

    [Fact]
    public async Task HttpErrorsResolveWithBodyAndHeadersAndRequestBodyIsSent()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(404, "{\"missing\":true}",
            new Dictionary<string, string> { ["Content-Type"] = "application/json", ["Set-Cookie"] = "hidden=1" })));
        ScriptNetworkEntry? entry = null;
        var result = await ScriptNetworkBroker.FetchAsync(Call(server.Url, "POST", "hello", headers: new() { ["X-Test"] = "yes", ["Content-Type"] = "text/plain" }),
            ["127.0.0.1"], default, value => entry = value);
        Assert.Equal(404, result.GetProperty("status").GetInt32());
        Assert.Equal("{\"missing\":true}", Encoding.UTF8.GetString(Convert.FromBase64String(result.GetProperty("bodyBase64").GetString()!)));
        Assert.DoesNotContain(result.GetProperty("headers").EnumerateArray(), h => h[0].GetString() == "set-cookie");
        var request = Assert.Single(server.Requests);
        Assert.Equal("POST", request.Method); Assert.Equal("yes", request.Headers["X-Test"]); Assert.Equal("hello", Encoding.UTF8.GetString(request.Body));
        Assert.NotNull(entry); Assert.Equal(404, entry.Status); Assert.Null(entry.Error); Assert.Equal("hello", Encoding.UTF8.GetString(entry.RequestBody!));
    }

    [Fact]
    public async Task DeniedHostsAndRedirectsNeverReachTheirDestination()
    {
        await using var target = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        await using var redirect = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(302, "", new Dictionary<string, string>
        { ["Location"] = target.Url.Replace("127.0.0.1", "localhost", StringComparison.Ordinal) })));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(target.Url), ["example.com"], default));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(redirect.Url), ["127.0.0.1"], default));
        Assert.Empty(target.Requests); Assert.Single(redirect.Requests);
    }

    [Fact]
    public async Task RedirectModesAndCrossOriginCredentialsFollowFetchRules()
    {
        await using var target = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        await using var redirect = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(302, "", new Dictionary<string, string> { ["Location"] = target.Url })));
        var result = await ScriptNetworkBroker.FetchAsync(Call(redirect.Url, "POST", "secret", headers: new() { ["Authorization"] = "Bearer secret" }), ["127.0.0.1"], default);
        Assert.True(result.GetProperty("redirected").GetBoolean());
        var request = Assert.Single(target.Requests);
        Assert.Equal("GET", request.Method); Assert.Empty(request.Body); Assert.False(request.Headers.ContainsKey("Authorization"));
        result = await ScriptNetworkBroker.FetchAsync(Call(redirect.Url, redirect: "manual"), ["127.0.0.1"], default);
        Assert.Equal(0, result.GetProperty("status").GetInt32()); Assert.Equal("opaqueredirect", result.GetProperty("type").GetString());
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(redirect.Url, redirect: "error"), ["127.0.0.1"], default));
        Assert.Single(target.Requests);
    }

    [Fact]
    public async Task ResponseSizeLimitIncludesDecompressedBytesAndRequestsAreBounded()
    {
        var options = new ScriptNetworkOptions(BodyLimitMiB: 1);
        using var compressed = new MemoryStream();
        using (var gzip = new GZipStream(compressed, CompressionMode.Compress, true)) gzip.Write(new byte[options.BodyLimitBytes + 1]);
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(200, "", new Dictionary<string, string> { ["Content-Encoding"] = "gzip" }, compressed.ToArray())));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], default, options: options));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url, "POST", new string('x', options.BodyLimitBytes + 1)), ["127.0.0.1"], default, options: options));
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task DefaultEightMiBLimitAllowsLargerRequestAndResponseBodies()
    {
        Assert.Equal(8 * 1024 * 1024, new ScriptNetworkOptions().BodyLimitBytes);
        var body = new string('x', 3 * 1024 * 1024);
        await using var server = new ScriptFetchTestServer(request => Task.FromResult(new ScriptFetchTestServer.Response(Bytes: request.Body)));
        var result = await ScriptNetworkBroker.FetchAsync(Call(server.Url, "POST", body), ["127.0.0.1"], default);
        Assert.Equal(body.Length, Convert.FromBase64String(result.GetProperty("bodyBase64").GetString()!).Length);
        Assert.Equal(body.Length, Assert.Single(server.Requests).Body.Length);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task MasterNetworkSwitchRejectsBeforeHttpEvenWithoutDeclaredHosts(bool declared)
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url),
            declared ? ["127.0.0.1"] : [], default, options: new(AllowNetwork: false)));
        Assert.Contains("禁止脚本访问网络", error.Message);
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task ConfiguredTimeoutStopsRequestWithoutRequiringCallerCancellation()
    {
        var accepted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new ScriptFetchTestServer(async _ => { accepted.TrySetResult(); await release.Task; return new(); });
        ScriptNetworkEntry? entry = null;
        var fetch = ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], default, value => entry = value,
            new(ActionTimeoutSeconds: 1));
        try
        {
            await accepted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            var error = await Assert.ThrowsAsync<HttpRequestException>(() => fetch.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Contains("超时", error.Message); Assert.Contains("超时", entry!.Error);
        }
        finally { release.TrySetResult(); }
    }

    [Fact]
    public async Task InvalidPolicyCannotRelaxLimitsOrSendRequests()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], default,
            options: new(BodyLimitMiB: 11)));
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task CancellationStopsPendingRequestsAndRecordsFailure()
    {
        var accepted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new ScriptFetchTestServer(async _ => { accepted.TrySetResult(); await release.Task; return new(); });
        using var cancellation = new CancellationTokenSource();
        ScriptNetworkEntry? entry = null;
        var fetch = ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], cancellation.Token, value => entry = value);
        await accepted.Task.WaitAsync(TimeSpan.FromSeconds(5)); cancellation.Cancel();
        try { await Assert.ThrowsAnyAsync<OperationCanceledException>(() => fetch); Assert.NotNull(entry?.Error); }
        finally { release.TrySetResult(); }
    }

    [Fact]
    public async Task CookiesAreNeverRetainedAndFragmentsAreExcludedFromResponseUrl()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(200, "ok", new Dictionary<string, string> { ["Set-Cookie"] = "session=secret" })));
        var result = await ScriptNetworkBroker.FetchAsync(Call(server.Url + "#private"), ["127.0.0.1"], default);
        Assert.Equal(server.Url, result.GetProperty("url").GetString());
        await ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], default);
        Assert.All(server.Requests, request => { Assert.False(request.Headers.ContainsKey("Cookie")); Assert.Equal("/", request.Path); });
    }

    [Fact]
    public async Task RedirectModeAppliesEvenWithoutALocationHeader()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(302)));
        var result = await ScriptNetworkBroker.FetchAsync(Call(server.Url, redirect: "manual"), ["127.0.0.1"], default);
        Assert.Equal("opaqueredirect", result.GetProperty("type").GetString());
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url, redirect: "error"), ["127.0.0.1"], default));
        result = await ScriptNetworkBroker.FetchAsync(Call(server.Url), ["127.0.0.1"], default);
        Assert.Equal(302, result.GetProperty("status").GetInt32());
    }

    [Theory]
    [InlineData("GET", "body", null)]
    [InlineData("TRACE", null, null)]
    [InlineData("POST", null, "Host")]
    [InlineData("POST", null, "Cookie")]
    [InlineData("POST", null, "Content-Length")]
    public async Task InvalidMethodsBodiesAndForbiddenHeadersFailBeforeSending(string method, string? body, string? header)
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptNetworkBroker.FetchAsync(Call(server.Url, method, body,
            headers: header is null ? null : new() { [header] = "bad" }), ["127.0.0.1"], default));
        Assert.Empty(server.Requests);
    }

    private static JsonElement Call(string url, string method = "GET", string? body = null, string redirect = "follow", Dictionary<string, string>? headers = null)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject(); writer.WriteString("url", url); writer.WriteString("method", method); writer.WriteString("redirect", redirect);
            if (body is not null) writer.WriteBase64String("bodyBase64", Encoding.UTF8.GetBytes(body));
            writer.WriteStartArray("headers");
            foreach (var h in headers ?? []) { writer.WriteStartArray(); writer.WriteStringValue(h.Key); writer.WriteStringValue(h.Value); writer.WriteEndArray(); }
            writer.WriteEndArray(); writer.WriteEndObject();
        }
        using var document = JsonDocument.Parse(buffer.ToArray()); return document.RootElement.Clone();
    }
}

internal sealed class ScriptFetchTestServer : IAsyncDisposable
{
    internal sealed record Request(string Method, string Path, IReadOnlyDictionary<string, string> Headers, byte[] Body);
    internal sealed record Response(int Status = 200, string Body = "ok", IReadOnlyDictionary<string, string>? Headers = null, byte[]? Bytes = null);
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _accept;
    private readonly ConcurrentBag<Task> _clients = [];
    internal ConcurrentQueue<Request> Requests { get; } = [];
    internal string Url { get; }
    internal ScriptFetchTestServer(Func<Request, Task<Response>> handler)
    {
        _listener.Start(); Url = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/";
        _accept = Accept(handler);
    }
    private async Task Accept(Func<Request, Task<Response>> handler)
    {
        try { while (!_stop.IsCancellationRequested) { var client = await _listener.AcceptTcpClientAsync(_stop.Token); _clients.Add(Serve(client, handler)); } }
        catch (OperationCanceledException) { }
        catch (SocketException) when (_stop.IsCancellationRequested) { }
    }
    private async Task Serve(TcpClient client, Func<Request, Task<Response>> handler)
    {
        using (client)
        try
        {
            var stream = client.GetStream(); var bytes = new List<byte>(); var next = new byte[1];
            while (bytes.Count < 65536)
            {
                if (await stream.ReadAsync(next, _stop.Token) == 0) return;
                bytes.Add(next[0]);
                if (bytes.Count >= 4 && bytes[^4] == 13 && bytes[^3] == 10 && bytes[^2] == 13 && bytes[^1] == 10) break;
            }
            var lines = Encoding.ASCII.GetString(bytes.ToArray()).Split("\r\n"); var first = lines[0].Split(' ');
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in lines.Skip(1).Where(l => l.Contains(':'))) { var split = line.IndexOf(':'); headers[line[..split]] = line[(split + 1)..].Trim(); }
            var body = new byte[headers.TryGetValue("Content-Length", out var length) ? int.Parse(length, System.Globalization.CultureInfo.InvariantCulture) : 0];
            await stream.ReadExactlyAsync(body, _stop.Token);
            var request = new Request(first[0], first[1], headers, body); Requests.Enqueue(request);
            var response = await handler(request); var payload = response.Bytes ?? Encoding.UTF8.GetBytes(response.Body);
            var head = new StringBuilder($"HTTP/1.1 {response.Status} Test\r\nConnection: close\r\nContent-Length: {payload.Length}\r\n");
            foreach (var h in response.Headers ?? new Dictionary<string, string>()) head.Append(h.Key).Append(": ").Append(h.Value).Append("\r\n");
            head.Append("\r\n"); await stream.WriteAsync(Encoding.ASCII.GetBytes(head.ToString()), _stop.Token); await stream.WriteAsync(payload, _stop.Token);
        }
        catch (Exception error) when (error is IOException or OperationCanceledException or SocketException) { }
    }
    public async ValueTask DisposeAsync()
    {
        _stop.Cancel(); _listener.Stop(); await _accept;
        await Task.WhenAll(_clients).WaitAsync(TimeSpan.FromSeconds(5)); _stop.Dispose();
    }
}
