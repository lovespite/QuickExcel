using System.Security.AccessControl;
using System.Security.Principal;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ReadOnlyWorkbookTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "read-only-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    public ReadOnlyWorkbookTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string PathFor(string name) => Path.Combine(_directory, name);
    private WorkbookWorkspace Create(TypedWorkbookStorage? storage = null)
    {
        storage ??= _storage;
        return new(new(storage, new WorkbookCalculator(new NoScripts())), new Files(storage));
    }
    private sealed class Files(TypedWorkbookStorage storage) : IWorkbookFileAccess
    {
        public IDisposable? AcquireExclusive(string path, bool allowMissing = false)
        {
            if (!allowMissing && !Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase)) TypedWorkbookStorage.CheckWriteAccess(path);
            return storage.AcquireExclusiveFile(path, allowMissing);
        }
        public IDisposable? AcquireReadOnly(string path) => storage.AcquireReadOnlyFile(path);
        public FileVersion ReadVersion(string path) => storage.ReadVersion(path);
        public OpenedWorkbook Open(string path) => new(storage.Open(path), storage.ReadVersion(path));
    }
    private string Seed(string name = "source.xlsq")
    {
        var path = Path.GetFullPath(PathFor(name));
        if (name.EndsWith(".csv")) File.WriteAllText(path, "42,hello\r\n");
        else _storage.Save(path, Book(Sheet("First", (0, 0, N(42))), Sheet("Second")));
        return path;
    }

    [Theory]
    [InlineData("source.xlsq")]
    [InlineData("source.xlsx")]
    [InlineData("source.csv")]
    public async Task ExplicitReadOnlyAllowsNavigationAndSharedReadsThenSaveAsBecomesEditable(string name)
    {
        var source = Seed(name); var bytes = File.ReadAllBytes(source);
        using var workspace = Create(); await workspace.OpenAsync(new(source), readOnly: true);
        Assert.True(workspace.IsReadOnly); Assert.True(workspace.Document.Session.IsReadOnly);
        Assert.False(workspace.Document.IsDirty); Assert.Equal("42", workspace.Document.GetCellValue(0, 0, 0));
        workspace.Document.Session.UpdateNavigation(workspace.Document.Snapshot.Sheets[^1].Id, new(2, 3));
        Assert.False(workspace.Document.IsDirty);
        using (var second = Create(new())) { await second.OpenAsync(new(source), readOnly: true); Assert.True(second.IsReadOnly); }
        Assert.Equal(bytes, File.ReadAllBytes(source));
        Assert.Empty(Directory.EnumerateFiles(_directory, "*.lock"));
        await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.SaveAsync(new(source)));
        Assert.Throws<InvalidOperationException>(() => workspace.Document.Save(source));
        Assert.Throws<InvalidOperationException>(() => workspace.Document.SaveChecked(source, workspace.Version!));
        if (!name.EndsWith(".csv")) Assert.Throws<InvalidOperationException>(() => _storage.Save(source, workspace.Document.Snapshot));
        await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ExportAsync(new(source), false));

        var copy = PathFor("editable.xlsq");
        await workspace.SaveAsync(new(copy));
        Assert.False(workspace.IsReadOnly); Assert.False(workspace.Document.IsDirty); Assert.False(workspace.Document.ScriptsEnabled);
        Assert.Equal(copy, workspace.Document.Snapshot.FileState!.SourcePath);
        Assert.Equal(copy, workspace.FilePath); Assert.Equal(bytes, File.ReadAllBytes(source));
        workspace.Editor.SetCell(0, 0, 0, "99"); Assert.True(workspace.Document.IsDirty);
        await workspace.SaveAsync(new(copy));
        Assert.Throws<IOException>(() => { using var blocked = File.OpenRead(copy); });
        workspace.New(); Assert.Equal(99, new TypedWorkbookStorage().Open(copy).Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Fact]
    public void SessionRejectsEveryWritePathAndScriptExecutionBeforeAnyStateChange()
    {
        var snapshot = Book(Sheet("First", (0, 0, N(42))));
        using var session = new DocumentSession(snapshot, readOnly: true);
        var cell = new CellAddress(snapshot.ActiveSheetId, 0, 0);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([new(cell, CellInput.Parse("changed"))])));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new AddSheetCommand("Other")));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetAxisMetadataCommand(cell.SheetId, SheetAxis.Row, 0, 1, 50)));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new UpdateScriptSourceCommand("host.functions.register('F',()=>1);")));
        Assert.Throws<InvalidOperationException>(() => session.Undo());
        Assert.Throws<InvalidOperationException>(() => session.Redo());
        Assert.Throws<InvalidOperationException>(() => session.SetScriptsEnabled(true));
        Assert.Throws<InvalidOperationException>(() => session.BeginAction("main"));
        Assert.Throws<InvalidOperationException>(() => session.BeginConsole("workbook.getActiveCell().setValue(1)"));
        Assert.Throws<InvalidOperationException>(() => session.BeginCalculation());
        session.SetScriptsEnabled(false);
        var stamp = new CalculationStamp(session.SessionId, snapshot.Revision, 0, 0);
        Assert.False(session.TryCommitCalculation(new(stamp, [])));
        Assert.False(session.TryFailCalculation(stamp, "not applicable", disableScripts: true));
        Assert.Same(snapshot, session.Snapshot); Assert.False(session.IsDirty); Assert.False(session.ScriptsEnabled);
    }

    [Theory]
    [InlineData("source.xlsq")]
    [InlineData("source.xlsx")]
    public async Task ReadOnlyFileAttributeAutomaticallyOpensReadOnly(string name)
    {
        if (!OperatingSystem.IsWindows()) return;
        var source = Seed(name); File.SetAttributes(source, FileAttributes.ReadOnly);
        try
        {
            using var workspace = Create(); await workspace.OpenAsync(new(source));
            Assert.True(workspace.IsReadOnly); Assert.Equal("42", workspace.Document.GetCellInput(0, 0, 0));
            Assert.Empty(Directory.EnumerateFiles(_directory, "*.lock"));
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.SaveAsync(new(source)));
        }
        finally { File.SetAttributes(source, FileAttributes.Normal); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ActualWindowsWriteDenialFallsBackWithoutCreatingSourceDirectoryFiles(bool denyDirectoryCreation)
    {
        if (!OperatingSystem.IsWindows()) return;
        var folder = PathFor("restricted"); Directory.CreateDirectory(folder);
        var source = Seed("restricted/source.xlsq"); var bytes = File.ReadAllBytes(source);
        var directoryInfo = new DirectoryInfo(folder); var fileInfo = new FileInfo(source);
        var originalDirectory = directoryInfo.GetAccessControl(); var originalFile = fileInfo.GetAccessControl();
        var identity = WindowsIdentity.GetCurrent().User!;
        try
        {
            if (denyDirectoryCreation)
            {
                var denied = directoryInfo.GetAccessControl();
                denied.AddAccessRule(new(identity, FileSystemRights.CreateFiles, AccessControlType.Deny)); directoryInfo.SetAccessControl(denied);
                Assert.Throws<UnauthorizedAccessException>(() => File.WriteAllText(Path.Combine(folder, "probe"), "denied"));
            }
            else
            {
                var denied = fileInfo.GetAccessControl();
                denied.AddAccessRule(new(identity, FileSystemRights.WriteData, AccessControlType.Deny)); fileInfo.SetAccessControl(denied);
                Assert.Throws<UnauthorizedAccessException>(() => { using var probe = File.OpenWrite(source); });
            }
            using var workspace = Create(); await workspace.OpenAsync(new(source));
            Assert.True(workspace.IsReadOnly); Assert.False(workspace.Document.IsDirty);
            Assert.Equal(new[] { source }, Directory.GetFiles(folder)); Assert.Equal(bytes, File.ReadAllBytes(source));
            await workspace.SaveAsync(new(PathFor("copy.xlsq")));
            Assert.False(workspace.IsReadOnly); Assert.Equal(bytes, File.ReadAllBytes(source));
        }
        finally { directoryInfo.SetAccessControl(originalDirectory); fileInfo.SetAccessControl(originalFile); }
    }

    [Fact]
    public async Task SharingConflictDoesNotDowngradeToReadOnlyAndFailedSaveAsKeepsReadOnlySource()
    {
        var source = Seed(); var occupied = Seed("occupied.xlsq");
        using var writer = Create(new()); await writer.OpenAsync(new(occupied));
        using var reader = Create(); await reader.OpenAsync(new(source), readOnly: true);
        var original = reader.Document; var version = reader.Version;
        await Assert.ThrowsAsync<IOException>(() => reader.OpenAsync(new(occupied)));
        await Assert.ThrowsAsync<IOException>(() => reader.OpenAsync(new(occupied), readOnly: true));
        await Assert.ThrowsAsync<IOException>(() => reader.SaveAsync(new(occupied)));
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => reader.SaveAsync(new(PathFor("missing/copy.xlsq"))));
        Assert.Same(original, reader.Document); Assert.True(reader.IsReadOnly); Assert.False(reader.Document.IsDirty);
        Assert.False(reader.IsBusy); Assert.Equal(source, reader.FilePath); Assert.Equal(version, reader.Version);
    }

    [Fact]
    public async Task ReadDenialStillFailsAndPreservesTheCurrentWorkbook()
    {
        if (!OperatingSystem.IsWindows()) return;
        var source = Seed(); var info = new FileInfo(source); var original = info.GetAccessControl();
        using var workspace = Create(); workspace.Editor.SetCell(0, 0, 0, "pending"); var document = workspace.Document;
        try
        {
            var denied = info.GetAccessControl();
            denied.AddAccessRule(new(WindowsIdentity.GetCurrent().User!, FileSystemRights.ReadData, AccessControlType.Deny)); info.SetAccessControl(denied);
            await Assert.ThrowsAsync<UnauthorizedAccessException>(() => workspace.OpenAsync(new(source)));
            Assert.Same(document, workspace.Document); Assert.True(document.IsDirty); Assert.False(workspace.IsReadOnly);
            Assert.False(workspace.IsBusy); Assert.Empty(Directory.GetFiles(_directory, "*.lock"));
        }
        finally { info.SetAccessControl(original); }
    }
}
