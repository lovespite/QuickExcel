using QuickExcel.Core;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;

namespace QuickExcel.Infrastructure.Web;

/// <summary>Per-view random loopback origin. Exact resource names, bounded HTTP, and revocable PDF leases.</summary>
public sealed class LocalWebOrigin : IAsyncDisposable
{
    public const string Csp = "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; style-src 'self' 'unsafe-inline'; font-src 'self' blob:; img-src 'self' data: blob:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'";
    private readonly WebAssetCatalog _assets;
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly SemaphoreSlim _slots = new(16);
    private readonly ConcurrentDictionary<int, Task> _requests = new();
    private readonly ConcurrentDictionary<string, PdfResource> _pdfs = new(StringComparer.Ordinal);
    private readonly Task _accept;
    private readonly object _lifecycle = new();
    private Task? _disposeTask;
    private int _nextRequest;
    private bool _disposed;
    public string Instance { get; } = Convert.ToHexString(RandomNumberGenerator.GetBytes(24)).ToLowerInvariant();
    public Uri Page { get; }
    public LocalWebOrigin(WebAssetCatalog assets, string page, string document, long generation, string language = "zh-CN")
    {
        if (page is not ("script.html" or "preview.html") || !assets.Files.ContainsKey(page) || !Guid.TryParseExact(document, "N", out _) || generation < 1)
            throw new ArgumentException(MessageText.Get("无效的本地页面或文档身份。"));
        _assets = assets; _listener.Start();
        Page = new($"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/{Instance}/{page}#document={document}&generation={generation}&lang={Uri.EscapeDataString(language)}");
        _accept = AcceptAsync();
    }
    public async Task<PdfLease> AddPdfAsync(string path, CancellationToken token = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_pdfs.Count >= 2) throw new InvalidOperationException(MessageText.Get("每个预览最多持有两份 PDF 资源。"));
        var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024, FileOptions.Asynchronous | FileOptions.RandomAccess);
        try
        {
            if (input.Length is < 5 or > 128 * 1024 * 1024) throw new InvalidDataException(MessageText.Get("PDF 资源长度无效或超过 128 MiB。"));
            var header = new byte[5]; await input.ReadExactlyAsync(header, token);
            if (!header.AsSpan().SequenceEqual("%PDF-"u8)) throw new InvalidDataException(MessageText.Get("预览资源不是 PDF。"));
            var resource = new PdfResource(input); var name = "documents/" + Convert.ToHexString(RandomNumberGenerator.GetBytes(24)).ToLowerInvariant() + ".pdf";
            lock (_lifecycle)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_pdfs.Count >= 2) throw new InvalidOperationException(MessageText.Get("每个预览最多持有两份 PDF 资源。"));
                if (!_pdfs.TryAdd(name, resource)) throw new IOException(MessageText.Get("PDF 资源标识冲突。"));
            }
            return new(this, name, new Uri(Page, name));
        }
        catch { await input.DisposeAsync(); throw; }
    }
    public sealed class PdfLease(LocalWebOrigin owner, string name, Uri uri) : IAsyncDisposable
    {
        public Uri Uri { get; } = uri;
        public async ValueTask DisposeAsync() { if (owner._pdfs.TryRemove(name, out var resource)) await resource.RetireAsync(); }
    }
    private sealed class PdfResource(FileStream stream)
    {
        private readonly object _gate = new();
        private readonly TaskCompletionSource _idle = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _readers; private bool _retired;
        public FileStream Stream => stream;
        public bool Acquire() { lock (_gate) { if (_retired) return false; _readers++; return true; } }
        public void Release() { lock (_gate) { _readers--; if (_retired && _readers == 0) _idle.TrySetResult(); } }
        public async Task RetireAsync()
        {
            lock (_gate) { _retired = true; if (_readers == 0) _idle.TrySetResult(); }
            await _idle.Task.ConfigureAwait(false); await stream.DisposeAsync().ConfigureAwait(false);
        }
    }
    private async Task AcceptAsync()
    {
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                await _slots.WaitAsync(_stop.Token).ConfigureAwait(false); TcpClient client;
                try { client = await _listener.AcceptTcpClientAsync(_stop.Token).ConfigureAwait(false); }
                catch { _slots.Release(); throw; }
                var id = Interlocked.Increment(ref _nextRequest); var task = ServeAsync(client); _requests[id] = task;
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
            timeout.CancelAfter(TimeSpan.FromSeconds(10)); PdfResource? held = null;
            try
            {
                var stream = client.GetStream(); var buffer = new byte[8192]; var size = 0; var end = -1;
                while (size < buffer.Length && end < 0)
                { var read = await stream.ReadAsync(buffer.AsMemory(size), timeout.Token); if (read == 0) return; size += read; end = Encoding.ASCII.GetString(buffer, 0, size).IndexOf("\r\n\r\n", StringComparison.Ordinal); }
                if (end < 0) return;
                var lines = Encoding.ASCII.GetString(buffer, 0, end).Split("\r\n"); var request = lines[0].Split(' ');
                var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); var valid = true;
                foreach (var line in lines.Skip(1)) { var colon = line.IndexOf(':'); if (colon <= 0 || !headers.TryAdd(line[..colon], line[(colon + 1)..].Trim())) { valid = false; break; } }
                valid &= request.Length == 3 && request[0] is "GET" or "HEAD" && request[2] == "HTTP/1.1" &&
                    headers.GetValueOrDefault("Host") == Page.Authority && headers.GetValueOrDefault("Sec-Fetch-Site") != "cross-site" &&
                    (!headers.TryGetValue("Origin", out var origin) || origin == Page.GetLeftPart(UriPartial.Authority)) &&
                    !headers.ContainsKey("Transfer-Encoding") && (!headers.TryGetValue("Content-Length", out var contentLength) || contentLength == "0");
                var prefix = '/' + Instance + '/'; var name = valid && request[1].StartsWith(prefix, StringComparison.Ordinal) ? request[1][prefix.Length..] : "";
                byte[]? asset = null; _assets.Files.TryGetValue(name, out asset);
                if (name.Length > 0 && _pdfs.TryGetValue(name, out var resource) && resource.Acquire()) held = resource;
                var found = asset is not null || held is not null;
                var length = asset?.LongLength ?? held?.Stream.Length ?? 0; long first = 0, last = length - 1;
                var status = found ? "200 OK" : "404 Not Found"; string? rangeHeader = null;
                if (found && headers.TryGetValue("Range", out var range))
                {
                    if (!ParseRange(range, length, out first, out last)) { status = "416 Range Not Satisfiable"; found = false; rangeHeader = $"bytes */{length}"; }
                    else { status = "206 Partial Content"; rangeHeader = $"bytes {first}-{last}/{length}"; }
                }
                var type = found ? Path.GetExtension(name) switch { ".html" => "text/html; charset=utf-8", ".js" or ".mjs" => "text/javascript; charset=utf-8",
                    ".css" => "text/css; charset=utf-8", ".pdf" => "application/pdf", ".wasm" => "application/wasm", ".ttf" => "font/ttf", ".svg" => "image/svg+xml", ".png" => "image/png", _ => "application/octet-stream" } : "text/plain";
                var response = $"HTTP/1.1 {status}\r\nContent-Type: {type}\r\nContent-Length: {(found ? last - first + 1 : 0)}\r\nConnection: close\r\nAccept-Ranges: bytes\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nReferrer-Policy: no-referrer\r\nCross-Origin-Resource-Policy: same-origin\r\nPermissions-Policy: camera=(), microphone=(), geolocation=(), usb=(), payment=(), display-capture=()\r\nContent-Security-Policy: {Csp}\r\n";
                if (rangeHeader is not null) response += "Content-Range: " + rangeHeader + "\r\n";
                await stream.WriteAsync(Encoding.ASCII.GetBytes(response + "\r\n"), timeout.Token);
                if (!found || request[0] == "HEAD") return;
                if (asset is not null) await stream.WriteAsync(asset.AsMemory((int)first, checked((int)(last - first + 1))), timeout.Token);
                else
                {
                    var chunk = new byte[64 * 1024];
                    for (var offset = first; offset <= last;)
                    {
                        var read = await RandomAccess.ReadAsync(held!.Stream.SafeFileHandle, chunk.AsMemory(0, (int)Math.Min(chunk.Length, last - offset + 1)), offset, timeout.Token);
                        if (read == 0) throw new EndOfStreamException(MessageText.Get("PDF 资源在读取中截断。"));
                        await stream.WriteAsync(chunk.AsMemory(0, read), timeout.Token); offset += read;
                    }
                }
            }
            catch (OperationCanceledException) when (timeout.IsCancellationRequested) { }
            catch (IOException) { /* Canceled browser read; other leases remain valid. */ }
            catch (SocketException) { }
            finally { held?.Release(); }
        }
    }
    private static bool ParseRange(string text, long length, out long first, out long last)
    {
        first = 0; last = length - 1;
        if (!text.StartsWith("bytes=", StringComparison.Ordinal) || text.Contains(',')) return false;
        var pair = text[6..].Split('-'); if (pair.Length != 2 || length == 0) return false;
        if (pair[0].Length == 0) { if (!long.TryParse(pair[1], out var tail) || tail <= 0) return false; first = Math.Max(0, length - tail); }
        else if (!long.TryParse(pair[0], out first) || first < 0 || first >= length) return false;
        else if (pair[1].Length > 0) { if (!long.TryParse(pair[1], out last) || last < first) return false; last = Math.Min(last, length - 1); }
        return true;
    }
    public ValueTask DisposeAsync()
    { lock (_lifecycle) return new(_disposeTask ??= DisposeCoreAsync()); }
    private async Task DisposeCoreAsync()
    {
        _disposed = true;
        await _stop.CancelAsync(); _listener.Stop(); await _accept.ConfigureAwait(false); await Task.WhenAll(_requests.Values).ConfigureAwait(false);
        foreach (var name in _pdfs.Keys) if (_pdfs.TryRemove(name, out var resource)) await resource.RetireAsync().ConfigureAwait(false);
        _stop.Dispose();
    }
}
