using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public sealed record OpenedWorkbook(WorkbookSnapshot Snapshot, FileVersion Version);
public interface IWorkbookFileAccess
{
    OpenedWorkbook Open(string path);
    FileVersion ReadVersion(string path);
}

/// <summary>One owned document and checked file identity; UI prompts remain in the host.</summary>
public sealed class WorkbookWorkspace : IDisposable
{
    private readonly WorkbookServices _services;
    private readonly IWorkbookFileAccess _files;
    private bool _disposed;
    private static StringComparison PathComparison => OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    public WorkbookDocument Document { get; private set; }
    public EditorController Editor { get; private set; }
    public string? FilePath { get; private set; }
    public FileVersion? Version { get; private set; }
    public bool IsBusy { get; private set; }
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
            [KeyValuePair.Create("main", new ScriptModule("scripts/main.ts", "", ScriptLanguage.TypeScript))],
            ["main"], scriptApiVersion: 2, bindings: [new("main", "formula-functions"), new("main", "actions")])), _services);
    }
    private void RequireIdle()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (IsBusy) throw new InvalidOperationException("正在处理工作簿，请稍后再试。");
    }
    public void New(bool demo = false)
    {
        RequireIdle(); Replace(demo ? WorkbookDocument.CreateDemo(_services) : NewDocument(), null, null);
    }
    private void Replace(WorkbookDocument document, string? path, FileVersion? version)
    {
        var previous = Document;
        Document = document; Editor = new(document); FilePath = path; Version = version;
        previous.NotifyClosed(); previous.Dispose(); Changed?.Invoke();
    }
    public async Task OpenAsync(LocalFileSelection selected)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        try
        {
            var opened = await Task.Run(() => _files.Open(selected.Path));
            Replace(new(opened.Snapshot, _services), selected.Path, opened.Version);
        }
        finally { IsBusy = false; Changed?.Invoke(); }
    }
    public async Task SaveAsync(LocalFileSelection selected, bool discardUnsupported = false, bool overwriteOriginal = false)
    {
        RequireIdle(); IsBusy = true; Changed?.Invoke();
        try
        {
            var sameFile = string.Equals(FilePath, selected.Path, PathComparison);
            if (overwriteOriginal && (!sameFile || Version?.Sha256 is null))
                throw new InvalidOperationException("覆盖只能用于当前已打开且具有版本记录的文件。");
            var expected = sameFile ? Version ?? new(null) : await Task.Run(() => _files.ReadVersion(selected.Path));
            var saved = await Task.Run(() => Document.SaveChecked(selected.Path, expected, discardUnsupported, overwriteOriginal));
            FilePath = selected.Path; Version = saved;
        }
        finally { IsBusy = false; Changed?.Invoke(); }
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
        if (string.Equals(FilePath, selected.Path, PathComparison)) throw new InvalidOperationException("导出需要使用新副本路径，不能覆盖当前打开的工作簿。");
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
        RequireIdle(); FilePath = null; Version = null; Document.DisableScripts(); Changed?.Invoke();
    }
    public void Dispose()
    {
        if (_disposed) return;
        RequireIdle(); _disposed = true; Document.NotifyClosed(); Document.Dispose();
    }
}
