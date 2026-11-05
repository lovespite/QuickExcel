using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public sealed record OpenedWorkbook(WorkbookSnapshot Snapshot, FileVersion Version);
public interface IWorkbookFileAccess
{
    IDisposable? AcquireExclusive(string path, bool allowMissing = false) => null;
    IDisposable? AcquireReadOnly(string path) => null;
    OpenedWorkbook Open(string path);
    FileVersion ReadVersion(string path);
}

/// <summary>One owned document and checked file identity; UI prompts remain in the host.</summary>
public sealed class WorkbookWorkspace : IDisposable
{
    private readonly WorkbookServices _services;
    private readonly IWorkbookFileAccess _files;
    private bool _disposed;
    private IDisposable? _fileLease;
    private static StringComparison PathComparison => OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    private bool IsCurrentFile(string path) => FilePath is not null &&
        string.Equals(Path.GetFullPath(FilePath), Path.GetFullPath(path), PathComparison);
    public WorkbookDocument Document { get; private set; }
    public EditorController Editor { get; private set; }
    public string? FilePath { get; private set; }
    public FileVersion? Version { get; private set; }
    public bool IsBusy { get; private set; }
    public bool IsReadOnly => Document.IsReadOnly;
    public event Action? Changed;

    public WorkbookWorkspace(WorkbookServices services, IWorkbookFileAccess files)
    {
        _services = services; _files = files;
        Document = NewDocument(); Editor = new(Document);
    }
    private WorkbookDocument NewDocument()
    {
        using var blank = WorkbookDocument.CreateBlank(_services);
        return new(blank.Snapshot.WithScriptModules(new(
            [KeyValuePair.Create("main", new ScriptModule("scripts/main.ts", "export function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host): void {\n    // Initialize shared Actions / Events state here.\n}\n", ScriptLanguage.TypeScript))],
            ["main"], scriptApiVersion: 4, bindings: [new("main", "initialization")])), _services);
    }
    private void RequireIdle()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (IsBusy) throw new InvalidOperationException(UiText.Get("正在处理工作簿，请稍后再试。"));
    }
    public void New(bool demo = false)
    {
        RequireIdle(); Replace(demo ? WorkbookDocument.CreateDemo(_services) : NewDocument(), null, null);
    }
    public void Import(WorkbookSnapshot snapshot)
    {
        RequireIdle(); Replace(new WorkbookDocument(snapshot, _services), null, null);
    }
    public void Recover(WorkbookSnapshot snapshot)
    {
        RequireIdle(); var document = new WorkbookDocument(snapshot, _services);
        document.Session.MarkRecoveredUnsaved(); Replace(document, null, null);
    }
    private void Replace(WorkbookDocument document, string? path, FileVersion? version, IDisposable? lease = null)
    {
        var previous = Document; var previousLease = _fileLease;
        Document = document; Editor = new(document); FilePath = path; Version = version;
        _fileLease = lease;
        try { previous.NotifyClosed(); previous.Dispose(); }
        finally { if (!ReferenceEquals(previousLease, lease)) previousLease?.Dispose(); }
        Changed?.Invoke();
    }
    public async Task OpenAsync(LocalFileSelection selected, bool readOnly = false)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        IDisposable? acquired = null;
        try
        {
            var reuse = IsCurrentFile(selected.Path) && _fileLease is not null;
            if (reuse) readOnly |= IsReadOnly;
            else
            {
                try { acquired = await Task.Run(() => readOnly ? _files.AcquireReadOnly(selected.Path) : _files.AcquireExclusive(selected.Path)); }
                catch (UnauthorizedAccessException) when (!readOnly)
                {
                    readOnly = true;
                    acquired = await Task.Run(() => _files.AcquireReadOnly(selected.Path));
                }
            }
            var opened = await Task.Run(() => _files.Open(selected.Path));
            var document = new WorkbookDocument(opened.Snapshot, _services, readOnly: readOnly);
            var lease = reuse ? _fileLease : acquired; acquired = null;
            Replace(document, selected.Path, opened.Version, lease);
        }
        finally { acquired?.Dispose(); IsBusy = false; Changed?.Invoke(); }
    }
    public async Task SaveAsync(LocalFileSelection selected, bool discardUnsupported = false, bool overwriteOriginal = false)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        IDisposable? acquired = null;
        WorkbookDocument? editableCopy = null;
        try
        {
            var sameFile = IsCurrentFile(selected.Path);
            if (sameFile && IsReadOnly) throw new InvalidOperationException(UiText.Get(DocumentSession.ReadOnlyMessage));
            if (overwriteOriginal && (!sameFile || Version?.Sha256 is null))
                throw new InvalidOperationException(UiText.Get("覆盖只能用于当前已打开且具有版本记录的文件。"));
            if (!sameFile || _fileLease is null) acquired = await Task.Run(() => _files.AcquireExclusive(selected.Path, allowMissing: true));
            var expected = sameFile ? Version ?? new(null) : await Task.Run(() => _files.ReadVersion(selected.Path));
            if (IsReadOnly) editableCopy = new(Document.Snapshot, _services);
            var saved = await Task.Run(() => (editableCopy ?? Document).SaveChecked(selected.Path, expected, discardUnsupported, overwriteOriginal));
            if (editableCopy is not null)
            {
                // Bind the saved package's compatibility/source metadata, rather than
                // carrying the read-only original's migration or loss report forward.
                var reopened = await Task.Run(() => _files.Open(selected.Path));
                if (reopened.Version != saved) throw new WorkbookFileConflictException();
                editableCopy.Dispose();
                editableCopy = new(reopened.Snapshot, _services);
                var copy = editableCopy; editableCopy = null;
                var lease = acquired; acquired = null;
                Replace(copy, selected.Path, saved, lease);
                return;
            }
            FilePath = selected.Path; Version = saved;
            if (!sameFile || acquired is not null)
            {
                var previousLease = _fileLease;
                _fileLease = acquired; acquired = null;
                previousLease?.Dispose();
            }
        }
        finally { editableCopy?.Dispose(); acquired?.Dispose(); IsBusy = false; Changed?.Invoke(); }
    }
    public async Task<bool> CalculateAsync(CancellationToken cancellationToken = default, bool forceFull = false)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        try { return await Task.Run(() => Document.Calculate(cancellationToken, forceFull), cancellationToken); }
        finally { IsBusy = false; Changed?.Invoke(); }
    }
    public async Task<T> InspectAsync<T>(Func<WorkbookSnapshot, T> inspect)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        try { var snapshot = Document.Snapshot; return await Task.Run(() => inspect(snapshot)); }
        finally { IsBusy = false; Changed?.Invoke(); }
    }
    public async Task<PortableExportReport> InspectPortableExportAsync()
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        try { return await Task.Run(Document.InspectPortableExport); }
        finally { IsBusy = false; Changed?.Invoke(); }
    }
    public async Task ExportAsync(LocalFileSelection selected, bool portable, bool allowStaleCaches = false, bool discardUnsupported = false)
    {
        RequireIdle();
        if (IsCurrentFile(selected.Path)) throw new InvalidOperationException(UiText.Get("导出需要使用新副本路径，不能覆盖当前打开的工作簿。"));
        IsBusy = true; Changed?.Invoke();
        try
        {
            await Task.Run(() =>
            {
                if (portable) Document.ExportPortableXlsx(selected.Path, allowStaleCaches, discardUnsupported);
                else Document.ExportXlsx(selected.Path, allowStaleCaches, discardUnsupported);
            });
        }
        finally { IsBusy = false; Changed?.Invoke(); }
    }
    public void DetachFile()
    {
        RequireIdle(); FilePath = null; Version = null; Document.DisableScripts();
        _fileLease?.Dispose(); _fileLease = null; Changed?.Invoke();
    }
    public void Dispose()
    {
        if (_disposed) return;
        RequireIdle(); _disposed = true;
        try { Document.NotifyClosed(); Document.Dispose(); }
        finally { _fileLease?.Dispose(); _fileLease = null; }
    }
}
