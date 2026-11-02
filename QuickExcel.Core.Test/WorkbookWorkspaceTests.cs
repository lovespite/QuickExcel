using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class WorkbookWorkspaceTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "workspace-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private WorkbookWorkspace Create(IWorkbookFileAccess? files = null) => new(new(_storage, new WorkbookCalculator(new NoScripts())), files ?? new Files(_storage));
    private string FilePath(string name) => Path.Combine(_directory, name);
    public WorkbookWorkspaceTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private sealed class Files(TypedWorkbookStorage storage) : IWorkbookFileAccess
    {
        public FileVersion ReadVersion(string path) => FileVersions.Read(path);
        public OpenedWorkbook Open(string path)
        {
            using var lease = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return new(storage.Open(path), FileVersions.Read(lease));
        }
    }
    [Fact]
    public async Task FailedSaveKeepsFileBindingAndCurrentEdits()
    {
        using var workspace = Create(); var original = FilePath("original.xlsq");
        workspace.Editor.SetCell(0, 0, 0, "1"); await workspace.SaveAsync(new(original));
        var version = workspace.Version;
        workspace.Editor.SetCell(0, 0, 0, "2");
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => workspace.SaveAsync(new(FilePath("missing/copy.xlsq"))));
        Assert.Equal(original, workspace.FilePath); Assert.Equal(version, workspace.Version);
        Assert.True(workspace.Document.IsDirty); Assert.False(workspace.IsBusy);
        Assert.Equal("2", workspace.Document.GetCellInput(0, 0, 0));
    }
    [Fact]
    public async Task DeletedOrReplacedFileCannotBeSilentlyRecreatedBySave()
    {
        using var workspace = Create(); var path = FilePath("original.xlsq");
        await workspace.SaveAsync(new(path)); workspace.Editor.SetCell(0, 0, 0, "pending");
        File.Delete(path);
        await Assert.ThrowsAsync<WorkbookFileConflictException>(() => workspace.SaveAsync(new(path)));
        Assert.False(File.Exists(path)); Assert.True(workspace.Document.IsDirty);
        await workspace.SaveAsync(new(FilePath("copy.xlsq"))); Assert.False(workspace.Document.IsDirty);
    }
    [Fact]
    public async Task FailedOpenPreservesDocumentSessionAndAuthorizationState()
    {
        using var workspace = Create(); workspace.Editor.SetCell(0, 0, 0, "pending");
        var session = workspace.Document.Session.SessionId;
        await Assert.ThrowsAsync<FileNotFoundException>(() => workspace.OpenAsync(new(FilePath("missing.xlsq"))));
        Assert.Equal(session, workspace.Document.Session.SessionId); Assert.True(workspace.Document.IsDirty);
        Assert.False(workspace.IsBusy); Assert.False(workspace.Document.ScriptsEnabled);
    }
    [Fact]
    public async Task RecoveryDetachForgetsFileIdentityWithoutDiscardingContent()
    {
        using var workspace = Create(); workspace.Editor.SetCell(0, 0, 0, "copy");
        await workspace.SaveAsync(new(FilePath("recovery.xlsq"))); workspace.DetachFile();
        Assert.Null(workspace.FilePath); Assert.Null(workspace.Version); Assert.False(workspace.Document.ScriptsEnabled);
        Assert.Equal("copy", workspace.Document.GetCellInput(0, 0, 0));
    }
}
