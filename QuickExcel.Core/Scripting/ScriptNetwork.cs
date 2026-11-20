using System.Collections.Immutable;

namespace QuickExcel.Core;

/// <summary>A diagnostic HTTP exchange. Evicted bodies are released even when a row remains selected.</summary>
public sealed class ScriptNetworkEntry
{
    private byte[]? _requestBody, _responseBody;
    public Guid Id { get; init; } = Guid.NewGuid();
    public DateTimeOffset Started { get; init; } = DateTimeOffset.UtcNow;
    public string Url { get; init; } = "";
    public string Method { get; init; } = "GET";
    public int? Status { get; init; }
    public string StatusText { get; init; } = "";
    public string ResponseUrl { get; init; } = "";
    public IReadOnlyList<KeyValuePair<string, string>> RequestHeaders { get; init; } = [];
    public IReadOnlyList<KeyValuePair<string, string>> ResponseHeaders { get; init; } = [];
    public TimeSpan Duration { get; init; }
    public string? Error { get; init; }
    public byte[]? RequestBody { get => Volatile.Read(ref _requestBody); init { _requestBody = value; RequestBytes = value?.Length ?? 0; } }
    public byte[]? ResponseBody { get => Volatile.Read(ref _responseBody); init { _responseBody = value; ResponseBytes = value?.Length ?? 0; } }
    public int RequestBytes { get; private set; }
    public int ResponseBytes { get; private set; }
    public bool BodiesEvicted { get; private set; }
    internal long RetainedBodyBytes => (long)(RequestBody?.Length ?? 0) + (ResponseBody?.Length ?? 0);
    // Conservative retained metadata accounting, including UTF-16 strings and containers.
    internal long MetadataBytes => 256L + 2L * (Url.Length + Method.Length + StatusText.Length + ResponseUrl.Length + (Error?.Length ?? 0)) +
        RequestHeaders.Concat(ResponseHeaders).Sum(h => 64L + 2L * (h.Key.Length + h.Value.Length));
    internal void ReleaseBodies()
    {
        Interlocked.Exchange(ref _requestBody, null); Interlocked.Exchange(ref _responseBody, null);
        BodiesEvicted = true;
    }
}

/// <summary>Explicit network consent and bounded diagnostics belong to one open document only.</summary>
public sealed class ScriptNetworkSession : IDisposable
{
    public const int DefaultMaxEntries = 200;
    public const long DefaultMaxBodyBytes = 50 * 1024 * 1024;
    private readonly object _gate = new();
    private readonly List<ScriptNetworkEntry> _entries = [];
    private readonly HashSet<string> _granted = new(StringComparer.Ordinal);
    private int _maxEntries;
    private long _maxBodyBytes;
    private long _maxRetainedBytes = DefaultMaxBodyBytes;
    private ScriptNetworkOptions _options = new();
    private CancellationTokenSource _networkLifetime = new();
    private long _retainedBytes, _metadataBytes, _generation;
    private bool _recording, _disposed;

    public ScriptNetworkSession(int maxEntries = DefaultMaxEntries, long maxBodyBytes = DefaultMaxBodyBytes)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxEntries);
        ArgumentOutOfRangeException.ThrowIfNegative(maxBodyBytes);
        _maxEntries = maxEntries; _maxBodyBytes = maxBodyBytes;
    }

    public Func<IReadOnlyList<string>, CancellationToken, bool>? Authorize { get; set; }
    public ScriptNetworkOptions Options { get { lock (_gate) return _options; } }
    public event Action? Changed;
    public bool RecordingEnabled { get { lock (_gate) return _recording && !_disposed; } set { lock (_gate) _recording = value && !_disposed; } }
    public long RetainedBodyBytes { get { lock (_gate) return _retainedBytes; } }
    public long RetainedMemoryBytes { get { lock (_gate) return _retainedBytes + _metadataBytes; } }
    public ImmutableArray<ScriptNetworkEntry> Entries { get { lock (_gate) return [.. _entries]; } }

    public void Configure(ScriptNetworkOptions options)
    {
        ArgumentNullException.ThrowIfNull(options); options.Validate();
        CancellationTokenSource? revoked = null;
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_options == options && _maxEntries == options.PanelEntryLimit && _maxBodyBytes == options.PanelMemoryLimitBytes) return;
            if (_options.AllowNetwork != options.AllowNetwork)
            {
                _granted.Clear(); _generation++;
                // Cancel outside the lock: HTTP cancellation may complete on another thread.
                revoked = _networkLifetime;
                _networkLifetime = new();
            }
            _options = options; _maxEntries = options.PanelEntryLimit;
            _maxBodyBytes = _maxRetainedBytes = options.PanelMemoryLimitBytes;
            EvictExcess();
        }
        if (revoked is not null) CancelAndDispose(revoked);
        NotifyChanged();
    }

    /// <summary>Every network call takes this lease, including calls made after a prior grant.</summary>
    public CancellationToken CaptureNetworkAccess(IReadOnlyList<string>? authorizedHosts = null)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_options.AllowNetwork) throw new InvalidOperationException("设置已禁止访问网络。");
            if (authorizedHosts is not null && !authorizedHosts.All(_granted.Contains))
                throw new InvalidOperationException("网络授权已失效，请重新应用脚本。");
            return _networkLifetime.Token;
        }
    }

    public void EnsureAuthorized(IReadOnlyList<string> hosts, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(hosts);
        cancellationToken.ThrowIfCancellationRequested();
        if (hosts.Count > 128 || hosts.Any(h => string.IsNullOrWhiteSpace(h) || h.Length > 253 || h.Any(char.IsControl)))
            throw new InvalidOperationException("网络权限声明无效。");
        var declared = hosts.Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        long generation;
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (declared.Length != 0 && !_options.AllowNetwork) throw new InvalidOperationException("设置已禁止访问网络，授权申请已拒绝。");
            if (declared.All(_granted.Contains)) return;
            generation = _generation;
        }
        // The UI callback must not run under the session lock: closing or revoking while
        // the dialog is open invalidates this response instead of resurrecting a grant.
        var allowed = Authorize?.Invoke(declared, cancellationToken) == true;
        cancellationToken.ThrowIfCancellationRequested();
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (generation != _generation) throw new OperationCanceledException("脚本授权已变化，请重新加载脚本。", cancellationToken);
            if (!allowed) throw new InvalidOperationException("未授权脚本访问所声明的网络地址。");
            _granted.UnionWith(declared);
        }
    }

    public void Revoke()
    {
        lock (_gate) { _granted.Clear(); _generation++; }
    }

    public void Record(ScriptNetworkEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        lock (_gate)
        {
            if (!_recording || _disposed) { entry.ReleaseBodies(); return; }
            // Metadata also has a finite bound; body retention is accounted separately.
            if (entry.Url.Length > 8192 || entry.ResponseUrl.Length > 8192 || entry.Method.Length > 32 ||
                entry.StatusText.Length > 1024 || entry.Error?.Length > 8192 ||
                entry.RequestHeaders.Count > 128 || entry.ResponseHeaders.Count > 128 ||
                entry.RequestHeaders.Concat(entry.ResponseHeaders).Sum(h => (long)h.Key.Length + h.Value.Length) > 65536)
            { entry.ReleaseBodies(); return; }
            _entries.Add(entry); _retainedBytes += entry.RetainedBodyBytes; _metadataBytes += entry.MetadataBytes;
            EvictExcess();
        }
        NotifyChanged();
    }

    private void EvictExcess()
    {
        while (_entries.Count > _maxEntries)
        {
            var oldest = _entries[0]; _retainedBytes -= oldest.RetainedBodyBytes;
            _metadataBytes -= oldest.MetadataBytes;
            oldest.ReleaseBodies(); _entries.RemoveAt(0);
        }
        foreach (var oldest in _entries)
        {
            if (_retainedBytes <= _maxBodyBytes && _retainedBytes + _metadataBytes <= _maxRetainedBytes) break;
            _retainedBytes -= oldest.RetainedBodyBytes; oldest.ReleaseBodies();
        }
        while (_retainedBytes + _metadataBytes > _maxRetainedBytes && _entries.Count != 0)
        {
            var oldest = _entries[0]; _retainedBytes -= oldest.RetainedBodyBytes; _metadataBytes -= oldest.MetadataBytes;
            oldest.ReleaseBodies(); _entries.RemoveAt(0);
        }
    }

    public void Clear()
    {
        lock (_gate)
        {
            foreach (var entry in _entries) entry.ReleaseBodies();
            _entries.Clear(); _retainedBytes = _metadataBytes = 0;
        }
        NotifyChanged();
    }

    private void NotifyChanged()
    {
        if (Changed is not { } changed) return;
        foreach (Action observer in changed.GetInvocationList())
            try { observer(); } catch (Exception error) { System.Diagnostics.Trace.TraceError($"Network diagnostic observer failed: {error.GetType().Name}"); }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true; _recording = false; _generation++; _granted.Clear(); Authorize = null;
            foreach (var entry in _entries) entry.ReleaseBodies();
            _entries.Clear(); _retainedBytes = _metadataBytes = 0; Changed = null;
        }
        CancelAndDispose(_networkLifetime);
    }

    private static void CancelAndDispose(CancellationTokenSource cancellation)
    {
        try { cancellation.Cancel(); }
        catch (AggregateException error) { System.Diagnostics.Trace.TraceError($"Network cancellation observer failed: {error.GetType().Name}"); }
        finally { cancellation.Dispose(); }
    }
}
