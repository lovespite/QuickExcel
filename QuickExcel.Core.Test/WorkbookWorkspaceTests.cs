using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class WorkbookWorkspaceTests : IDisposable
{
    [Fact]
    public async Task ImportedCopyRequiresSaveBeforeFirstEditAndDoesNotBindSource()
    {
        using var workspace = Create();
        var source = FilePath("import-source.xlsx");
        _storage.Save(source, Book(Sheet("Imported")));
        var bytes = File.ReadAllBytes(source);
        var imported = _storage.Open(source);
        workspace.Import(imported);
        Assert.True(workspace.Document.IsDirty);
        Assert.False(workspace.Document.CanUndo);
        Assert.False(workspace.Document.ScriptsEnabled);
        Assert.Null(workspace.FilePath);
        Assert.Null(workspace.Version);
        await workspace.SaveAsync(new(FilePath("import-copy.xlsx")));
        Assert.False(workspace.Document.IsDirty);
        Assert.Equal(bytes, File.ReadAllBytes(source));
    }

    [Fact]
    public void DefaultMainIsInitializationAndNeverAnAction()
    {
        using var workspace = new WorkbookWorkspace(new(_storage, new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),
            new ScriptWorkerActionExecutor()), new Files(_storage));
        var document = workspace.Document;
        document.ScriptSource = "export function main(workbook: QuickExcel.Workbook) { workbook.getActiveCell().setValue(42); }";
        document.EnableScriptsAndCalculate();
        Assert.Empty(document.ListActions());
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.True(document.Undo());
        document.Calculate();
        Assert.Equal("", document.GetCellInput(0, 0, 0));
    }

    [Fact]
    public async Task NewWorkbooksUseTypeScriptMainWithoutRewritingOpenedJavaScript()
    {
        using var workspace = Create();
        CheckMain();
        var path = FilePath("legacy-js.xlsq");
        var legacy = ScriptModuleSet.Single("// Existing JavaScript source");
        _storage.Save(path, Book(Sheet("S")).WithScriptModules(legacy));
        await workspace.OpenAsync(new(path));
        var opened = workspace.Document.Snapshot.ScriptModules;
        Assert.Equal(3, opened.ScriptApiVersion);
        Assert.Equal(legacy.Modules["main"], opened.Modules["main"]);
        Assert.Equal(legacy.Bindings.ToArray(), opened.Bindings.ToArray());
        workspace.New();
        CheckMain();
        await workspace.SaveAsync(new(FilePath("typescript-main.xlsq")));
        await workspace.OpenAsync(new(FilePath("typescript-main.xlsq")));
        CheckMain();

        void CheckMain()
        {
            var modules = workspace.Document.Snapshot.ScriptModules;
            Assert.Equal(4, modules.ScriptApiVersion);
            Assert.Equal("main", Assert.Single(modules.EntryModules));
            Assert.Equal("scripts/main.ts", Assert.Single(modules.Modules).Value.Path);
            Assert.Contains("export function main", modules.PrimarySource);
            Assert.Equal("main", Assert.Single(modules.EntriesFor("initialization")));
            Assert.Empty(modules.EntriesFor("formula-functions"));
            Assert.Empty(modules.EntriesFor("actions"));
            Assert.False(workspace.Document.IsDirty);
            Assert.False(workspace.Document.ScriptsEnabled);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportPreservesBindingDirtyStateAndHistory(bool portable)
    {
        using var workspace = Create(); var path = FilePath("original.xlsq");
        workspace.Editor.SetCell(0, 0, 0, "1"); await workspace.SaveAsync(new(path));
        var version = workspace.Version;
        workspace.Editor.SetCell(0, 0, 0, "2"); workspace.Editor.SetCell(0, 0, 1, "=A1+1"); await workspace.CalculateAsync();
        var revision = workspace.Document.Snapshot.Revision;
        var copy = FilePath("export.xlsx"); await workspace.ExportAsync(new(copy), portable);
        Assert.True(File.Exists(copy)); Assert.Equal(path, workspace.FilePath); Assert.Equal(version, workspace.Version);
        Assert.True(workspace.Document.IsDirty); Assert.Equal(revision, workspace.Document.Snapshot.Revision);
        Assert.False(workspace.IsBusy); workspace.Editor.Undo(); Assert.Equal("", workspace.Document.GetCellInput(0, 0, 1));
        var output = _storage.Open(copy); Assert.Equal(portable ? CellContentKind.Formula : CellContentKind.Literal, output.Sheets[0].GetCell(0, 1)!.Content.Kind);
    }
    [Fact]
    public async Task FailedExportAndInspectionReleaseBusyAndNeverOverwriteOpenedFile()
    {
        using var workspace = Create(); var path = FilePath("original.xlsq"); await workspace.SaveAsync(new(path));
        var bytes = File.ReadAllBytes(path); workspace.Editor.SetCell(0, 0, 0, "pending");
        await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ExportAsync(new(path), false));
        Assert.Equal(bytes, File.ReadAllBytes(path));
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => workspace.ExportAsync(new(FilePath("missing/export.xlsx")), false));
        await Assert.ThrowsAsync<IOException>(() => workspace.InspectAsync<int>(_ => throw new IOException("injected")));
        Assert.False(workspace.IsBusy); Assert.True(workspace.Document.IsDirty); Assert.Equal(path, workspace.FilePath);
    }
    [Fact]
    public void ShowAllRowsDoesNotMaterializeDefaultRowMetadata()
    {
        using var workspace = Create(); var id = workspace.Document.Snapshot.ActiveSheetId;
        workspace.Editor.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 150_000, 1, hidden: true));
        workspace.Editor.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 0, WorkbookLimits.MaxRows, hidden: false));
        var sheet = workspace.Document.Snapshot.GetSheet(id);
        Assert.Single(sheet.Rows); Assert.False(sheet.Rows[150_000].Hidden);
        workspace.Editor.Undo(); Assert.True(workspace.Document.Snapshot.GetSheet(id).Rows[150_000].Hidden);
    }
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
