using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace QuickExcel.AvaloniaProbe;

/// <summary>A0-only loopback origin. Serves a verified, immutable allowlist, never a filesystem root.</summary>
internal sealed class ProbeResources : IAsyncDisposable
{
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly ConcurrentDictionary<int, Task> _requests = new();
    private readonly Dictionary<string, byte[]> _resources = new(StringComparer.Ordinal);
    private readonly SemaphoreSlim _slots = new(16);
    private readonly Task _accept;
    private int _nextId;
    internal string Instance { get; } = Convert.ToHexString(RandomNumberGenerator.GetBytes(24)).ToLowerInvariant();
    internal Uri Page { get; }
    internal const string Csp = "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; style-src 'self' 'unsafe-inline'; font-src 'self'; img-src 'self' data: blob:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'";

    internal ProbeResources(string directory, string pdfPath)
    {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "manifest.json")));
        foreach (var item in manifest.RootElement.GetProperty("files").EnumerateArray())
        {
            var name = item.GetProperty("name").GetString()!;
            if (name.StartsWith('/') || name.Contains('\\') || name.Split('/').Any(p => p is ".." or "." or ""))
                throw new InvalidDataException("Invalid resource name.");
            var bytes = File.ReadAllBytes(Path.Combine(directory, name));
            if (bytes.Length > 32 * 1024 * 1024 || bytes.Length != item.GetProperty("size").GetInt32() ||
                !Convert.ToHexString(SHA256.HashData(bytes)).Equals(item.GetProperty("sha256").GetString(), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"Resource integrity mismatch: {name}");
            _resources.Add('/' + Instance + '/' + name, bytes);
        }
        _resources.Add('/' + Instance + "/fixture.pdf", File.ReadAllBytes(pdfPath));
        _listener.Start();
        Page = new Uri($"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/{Instance}/probe.html");
        _accept = AcceptAsync();
    }

    private async Task AcceptAsync()
    {
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                await _slots.WaitAsync(_stop.Token);
                TcpClient client;
                try { client = await _listener.AcceptTcpClientAsync(_stop.Token); }
                catch { _slots.Release(); throw; }
                var id = Interlocked.Increment(ref _nextId);
                var task = ServeAsync(client);
                _requests[id] = task;
                _ = task.ContinueWith(_ => { _requests.TryRemove(id, out var ignored); _slots.Release(); }, TaskScheduler.Default);
            }
        }
        catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
        catch (SocketException) when (_stop.IsCancellationRequested) { }
    }

    private async Task ServeAsync(TcpClient client)
    {
        using (client)
        using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token))
        {
            timeout.CancelAfter(TimeSpan.FromSeconds(10));
            try
            {
                var stream = client.GetStream();
                var buffer = new byte[8192];
                var length = 0;
                var end = -1;
                while (length < buffer.Length && end < 0)
                {
                    var count = await stream.ReadAsync(buffer.AsMemory(length), timeout.Token);
                    if (count == 0) return;
                    length += count;
                    end = Encoding.ASCII.GetString(buffer, 0, length).IndexOf("\r\n\r\n", StringComparison.Ordinal);
                }
                if (end < 0) return;
                var lines = Encoding.ASCII.GetString(buffer, 0, end).Split("\r\n");
                var request = lines[0].Split(' ');
                var host = lines.Skip(1).Where(l => l.StartsWith("Host:", StringComparison.OrdinalIgnoreCase)).ToArray();
                var crossSite = lines.Any(l => l.Equals("Sec-Fetch-Site: cross-site", StringComparison.OrdinalIgnoreCase));
                var valid = request.Length == 3 && request[0] == "GET" && request[2] == "HTTP/1.1" &&
                    host.Length == 1 && host[0][5..].Trim() == Page.Authority && !crossSite;
                byte[]? content = null;
                var found = valid && _resources.TryGetValue(request[1], out content);
                var type = found ? Path.GetExtension(request[1]) switch
                {
                    ".html" => "text/html; charset=utf-8", ".js" or ".mjs" => "text/javascript; charset=utf-8",
                    ".css" => "text/css; charset=utf-8", ".pdf" => "application/pdf", ".wasm" => "application/wasm",
                    ".ttf" => "font/ttf", _ => "application/octet-stream"
                } : "text/plain";
                var body = content ?? "Not found"u8.ToArray();
                var headers = $"HTTP/1.1 {(found ? "200 OK" : "404 Not Found")}\r\nContent-Type: {type}\r\nContent-Length: {body.Length}\r\nConnection: close\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nReferrer-Policy: no-referrer\r\nCross-Origin-Resource-Policy: same-origin\r\nPermissions-Policy: camera=(), microphone=(), geolocation=(), usb=()\r\nContent-Security-Policy: {Csp}\r\n\r\n";
                await stream.WriteAsync(Encoding.ASCII.GetBytes(headers), timeout.Token);
                await stream.WriteAsync(body, timeout.Token);
            }
            catch (OperationCanceledException) when (timeout.IsCancellationRequested) { }
            catch (IOException) { /* Browser canceled the request. */ }
        }
    }

    internal async Task CheckBoundaryAsync()
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        using var valid = await client.GetAsync(Page);
        if (valid.StatusCode != HttpStatusCode.OK || !valid.Headers.Contains("Content-Security-Policy"))
            throw new InvalidDataException("Local page/CSP unavailable.");
        foreach (var suffix in new[] { "/probe.html", $"/{Instance}/../probe.html", $"/{Instance}/%2e%2e/probe.html", $"/{Instance}/fixture.pdf?extra=1" })
        {
            using var denied = await client.GetAsync(new Uri(Page.GetLeftPart(UriPartial.Authority) + suffix));
            if (denied.StatusCode != HttpStatusCode.NotFound) throw new InvalidDataException("Resource escaped the exact allowlist.");
        }
        using var wrongHost = new HttpRequestMessage(HttpMethod.Get, Page);
        wrongHost.Headers.Host = "untrusted.invalid";
        using var wrongHostResult = await client.SendAsync(wrongHost);
        if (wrongHostResult.StatusCode != HttpStatusCode.NotFound) throw new InvalidDataException("Unexpected host accepted.");
    }

    internal string WriteVirtualDirectory()
    {
        // Probe PDFs contain generated test data only. This is not a production
        // document cache, and never maps an arbitrary workbook directory.
        var root = Path.Combine(Program.Output, "virtual-origin");
        if (Directory.Exists(root)) throw new IOException("Virtual-origin evidence already exists.");
        foreach (var (name, bytes) in _resources)
        {
            var file = Path.Combine(root, name.TrimStart('/'));
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            File.WriteAllBytes(file, bytes);
        }
        return root;
    }

    public async ValueTask DisposeAsync()
    {
        await _stop.CancelAsync();
        _listener.Stop();
        await _accept.ConfigureAwait(false);
        await Task.WhenAll(_requests.Values).ConfigureAwait(false);
        _resources.Clear();
        _stop.Dispose();
        // Continuations may still be returning slots; SemaphoreSlim has no OS handle to release here.
    }
}
