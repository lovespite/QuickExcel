using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

/// <summary>One immutable job, completed PDF and cancellable platform operations; no UI or browser dependency.</summary>
public sealed class PrintSession : IAsyncDisposable
{
    private readonly SkiaReport _report;
    private readonly string _directory;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly SemaphoreSlim _operation = new(1);
    private readonly object _gate = new();
    private readonly TaskCompletionSource _drained = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private int _users;
    private bool _closing;
    private Task? _dispose;
    public PrintJob Job => _report.Job;
    public IReadOnlyList<string> Warnings => _report.Job.Warnings.Concat(_report.Renderer.Diagnostics.Select(d => d.Message)).Distinct().ToArray();
    public string PdfPath { get; }
    private PrintSession(SkiaReport report, string directory) { _report = report; _directory = directory; PdfPath = Path.Combine(directory, "preview.pdf"); }
    public static async Task<PrintSession> CreateAsync(WorkbookSnapshot snapshot, SheetId sheetId, string fileName, string temporaryRoot,
        Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition>? style = null,
        IEnumerable<string>? fontFiles = null, bool systemFonts = true, CancellationToken cancellationToken = default)
    {
        var root = Path.GetFullPath(temporaryRoot); Directory.CreateDirectory(root);
        if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0) throw new IOException("打印临时目录不能是重解析点。");
        var directory = Path.Combine(root, "print-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        PrintSession? session = null;
        try
        {
            return await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                var report = new SkiaReport(snapshot, snapshot.GetSheet(sheetId), fileName, style, fontFiles, systemFonts, cancellationToken: cancellationToken);
                session = new(report, directory);
                report.WritePdf(session.PdfPath, cancellationToken);
                report.ReleaseDrawingCaches();
                cancellationToken.ThrowIfCancellationRequested(); return session;
            }, cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            if (session is not null) await session.DisposeAsync().ConfigureAwait(false);
            else Directory.Delete(directory); // Only this newly created, non-recursive directory.
            throw;
        }
    }
    /// <summary>The temporary file cannot be deleted until every browser/resource consumer releases its lease.</summary>
    public IAsyncDisposable RetainPdf()
    { lock (_gate) { ObjectDisposedException.ThrowIf(_closing, this); _users++; return new Lease(this); } }
    private sealed class Lease(PrintSession owner) : IAsyncDisposable
    {
        private PrintSession? _owner = owner;
        public ValueTask DisposeAsync() { Interlocked.Exchange(ref _owner, null)?.Release(); return ValueTask.CompletedTask; }
    }
    private void Release() { lock (_gate) { if (--_users == 0 && _closing) _drained.TrySetResult(); } }
    private async Task<T> RunAsync<T>(Func<CancellationToken, T> operation, CancellationToken cancellationToken)
    {
        CancellationTokenSource linked;
        lock (_gate) { ObjectDisposedException.ThrowIf(_closing, this); linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetime.Token); _users++; }
        using (linked)
        try
        {
            await _operation.WaitAsync(linked.Token).ConfigureAwait(false);
            try { return await Task.Run(() => operation(linked.Token), linked.Token).ConfigureAwait(false); }
            finally { _operation.Release(); }
        }
        finally { Release(); }
    }
    public Task ExportAsync(string destination, CancellationToken cancellationToken = default)
    {
        if (Path.GetFullPath(destination).Equals(PdfPath, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new ArgumentException("导出目标不能覆盖当前预览资源。");
        return RunAsync(token => { try { _report.WritePdf(destination, token); return true; } finally { _report.ReleaseDrawingCaches(); } }, cancellationToken);
    }
    public Task<SKBitmap> RenderPageAsync(int page, double dpi, CancellationToken cancellationToken = default)
    {
        if (dpi is <= 0 or > 300 || !double.IsFinite(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi), "系统打印最高 300 DPI。");
        return RunAsync(token => _report.RenderPage(page, dpi, token), cancellationToken);
    }
    public ValueTask DisposeAsync() { lock (_gate) return new(_dispose ??= DisposeCoreAsync()); }
    private async Task DisposeCoreAsync()
    {
        _closing = true; if (_users == 0) _drained.TrySetResult();
        await _lifetime.CancelAsync().ConfigureAwait(false); await _drained.Task.ConfigureAwait(false);
        _report.Dispose(); _operation.Dispose(); _lifetime.Dispose();
        File.Delete(PdfPath); Directory.Delete(_directory);
    }
}
