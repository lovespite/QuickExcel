using System.Diagnostics;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ExclusiveWorkbookWorkspaceTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "exclusive-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    public ExclusiveWorkbookWorkspaceTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string PathFor(string name) => Path.Combine(_directory, name);
    private WorkbookWorkspace Create(TypedWorkbookStorage? storage = null)
    {
        storage ??= _storage;
        return new(new(storage, new WorkbookCalculator(new NoScripts())), new Files(storage));
    }
    private sealed class Files(TypedWorkbookStorage storage) : IWorkbookFileAccess
    {
        public IDisposable? AcquireExclusive(string path, bool allowMissing = false) => storage.AcquireExclusiveFile(path, allowMissing);
        public IDisposable? AcquireReadOnly(string path) => storage.AcquireReadOnlyFile(path);
        public FileVersion ReadVersion(string path) => storage.ReadVersion(path);
        public OpenedWorkbook Open(string path) => new(storage.Open(path), storage.ReadVersion(path));
    }
    private string Seed(string name = "book.xlsq")
    {
        var path = PathFor(name); _storage.Save(path, Book(Sheet("Sheet1"))); return path;
    }
    private static void AssertLocked(string path)
    {
        Assert.Throws<IOException>(() => { using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete); });
        Assert.Throws<IOException>(() => { using var stream = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete); });
        if (OperatingSystem.IsWindows()) Assert.Throws<IOException>(() => File.Delete(path));
    }
    private static void AssertReleased(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public async Task OpenBlocksAnotherWorkspaceAndExternalAccessUntilClose()
    {
        var path = Seed();
        using var first = Create(); using var second = Create(new());
        await first.OpenAsync(new(path));
        var previous = second.Document;
        AssertLocked(path);
        await Assert.ThrowsAsync<IOException>(() => second.OpenAsync(new(path)));
        Assert.Same(previous, second.Document); Assert.False(second.IsBusy);
        first.New();
        AssertReleased(path);
        await second.OpenAsync(new(path)); AssertLocked(path);
        second.Dispose(); AssertReleased(path);
        Assert.Empty(Directory.EnumerateFiles(_directory, "*.lock"));
    }

    [Fact]
    public async Task SavesAndReopensKeepOwnershipAndSaveAsTransfersIt()
    {
        using var workspace = Create();
        var original = PathFor("original.XLSQ"); var copy = PathFor("copy.xlsq");
        workspace.Editor.SetCell(0, 0, 0, "first");
        await workspace.SaveAsync(new(original)); AssertLocked(original);
        for (var i = 0; i < 3; i++)
        {
            workspace.Editor.SetCell(0, 0, 0, "saved " + i);
            await workspace.SaveAsync(new(original)); AssertLocked(original);
            await workspace.OpenAsync(new(Path.Combine(_directory, ".", "original.XLSQ")));
            Assert.Equal("saved " + i, workspace.Document.GetCellInput(0, 0, 0));
            Assert.False(workspace.Document.IsDirty); AssertLocked(original);
        }
        await workspace.SaveAsync(new(copy)); AssertReleased(original); AssertLocked(copy);
        workspace.DetachFile(); AssertReleased(copy); Assert.Null(workspace.FilePath);
        Assert.Equal("saved 2", workspace.Document.GetCellInput(0, 0, 0));
    }

    [Fact]
    public async Task FailedOpenReleasesCandidateAndRetainsOriginalLeaseAndEdits()
    {
        var original = Seed(); var malformed = PathFor("malformed.xlsq"); File.WriteAllText(malformed, "invalid ZIP");
        using var workspace = Create(); await workspace.OpenAsync(new(original));
        workspace.Editor.SetCell(0, 0, 0, "pending"); var document = workspace.Document;
        await Assert.ThrowsAsync<InvalidDataException>(() => workspace.OpenAsync(new(malformed)));
        Assert.Same(document, workspace.Document); Assert.True(document.IsDirty);
        AssertLocked(original); AssertReleased(malformed);
        await Assert.ThrowsAsync<FileNotFoundException>(() => workspace.OpenAsync(new(PathFor("missing.xlsq"))));
        AssertLocked(original); Assert.False(workspace.IsBusy);
        Assert.Single(Directory.EnumerateFiles(_directory, "*.lock"));
    }

    [Fact]
    public async Task FailedSaveAsToOccupiedFileRetainsBothWorkspacesAndTheirContents()
    {
        var original = Seed(); var target = Seed("target.xlsq");
        using var first = Create(); using var second = Create(new());
        await first.OpenAsync(new(original)); await second.OpenAsync(new(target));
        first.Editor.SetCell(0, 0, 0, "pending"); var version = first.Version;
        await Assert.ThrowsAsync<IOException>(() => first.SaveAsync(new(target)));
        Assert.Equal(original, first.FilePath); Assert.Equal(version, first.Version);
        Assert.True(first.Document.IsDirty); Assert.False(first.IsBusy);
        Assert.Equal("", second.Document.GetCellInput(0, 0, 0));
        AssertLocked(original); AssertLocked(target);
        await first.SaveAsync(new(original)); Assert.False(first.Document.IsDirty);
    }

    [Fact]
    public async Task ValidationFailureKeepsOriginalBytesAndReleasesNewDestination()
    {
        var original = Seed(); var bytes = File.ReadAllBytes(original);
        using var workspace = Create(); await workspace.OpenAsync(new(original));
        // Legacy script versions can be inspected but must never be silently saved as v3.
        var legacy = workspace.Document.Snapshot.WithScriptModules(new(
            [KeyValuePair.Create("main", new ScriptModule("scripts/main.js", ""))], ["main"], scriptApiVersion: 2));
        Assert.Throws<NotSupportedException>(() => _storage.SaveChecked(original, legacy, false, workspace.Version!));
        AssertLocked(original); Assert.Equal(workspace.Version, _storage.ReadVersion(original));
        workspace.New(); Assert.Equal(bytes, File.ReadAllBytes(original));

        // A failed Save As must neither bind nor leave an empty destination behind.
        var missing = PathFor("missing/copy.xlsq");
        workspace.Editor.SetCell(0, 0, 0, "pending");
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => workspace.SaveAsync(new(missing)));
        Assert.Null(workspace.FilePath); Assert.True(workspace.Document.IsDirty); Assert.False(File.Exists(missing));
        Assert.Empty(Directory.EnumerateFiles(_directory, "*.lock"));
    }

    [Fact]
    public async Task SwitchingToXlsxReleasesXlsqAndKeepsXlsxShared()
    {
        var original = Seed(); var xlsx = PathFor("shared.xlsx"); _storage.Save(xlsx, Book(Sheet("Sheet1")));
        using var workspace = Create(); await workspace.OpenAsync(new(original));
        await workspace.OpenAsync(new(xlsx)); AssertReleased(original); AssertReleased(xlsx);
        using var second = Create(new()); await second.OpenAsync(new(xlsx));
    }

    [Fact]
    public async Task FailedAtomicReplaceReacquiresOriginalLockAndKeepsDirtyState()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Seed(); var original = File.ReadAllBytes(path);
        using var workspace = Create(); await workspace.OpenAsync(new(path));
        workspace.Editor.SetCell(0, 0, 0, "pending"); var version = workspace.Version;
        File.SetAttributes(path, FileAttributes.ReadOnly);
        try
        {
            await Assert.ThrowsAsync<UnauthorizedAccessException>(() => workspace.SaveAsync(new(path)));
            Assert.True(workspace.Document.IsDirty); Assert.Equal(version, workspace.Version);
            Assert.False(workspace.IsBusy);
            Assert.Throws<IOException>(() => { using var reader = File.OpenRead(path); });
            Assert.Empty(Directory.EnumerateFiles(_directory, "*.tmp"));
        }
        finally { File.SetAttributes(path, FileAttributes.Normal); }
        AssertLocked(path);
        workspace.New(); Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public async Task WindowsChildProcessCannotReadAnOwnedXlsq()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Seed(); using var workspace = Create(); await workspace.OpenAsync(new(path));
        var start = new ProcessStartInfo("powershell.exe") { UseShellExecute = false, CreateNoWindow = true };
        start.ArgumentList.Add("-NoProfile"); start.ArgumentList.Add("-NonInteractive"); start.ArgumentList.Add("-Command");
        start.ArgumentList.Add("try { $s = [IO.File]::OpenRead($env:QUICKEXCEL_TEST_LOCKED_FILE); $s.Dispose(); exit 1 } catch { if (($_.Exception.InnerException.HResult -band 65535) -eq 32) { exit 0 }; exit 2 }");
        start.Environment["QUICKEXCEL_TEST_LOCKED_FILE"] = path;
        using var child = Process.Start(start)!;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        try { await child.WaitForExitAsync(timeout.Token); Assert.Equal(0, child.ExitCode); }
        finally { if (!child.HasExited) child.Kill(true); }
    }
}
