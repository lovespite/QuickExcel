using System.Diagnostics;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class WorkbookRecoveryTests
{
    private static string Folder() { var path = Path.Combine(AppContext.BaseDirectory, "recovery-evidence", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
    [Fact]
    public void LiveLeaseIsHiddenAndRecoveryPreservesMetadataWithoutAuthorizingScripts()
    {
        var root = Folder(); var store = new WorkbookRecoveryStore(root); var sheet = Sheet("S", (0, 0, T("中文")));
        using var session = new DocumentSession(Book(sheet));
        using (var lease = store.Start(session.SessionId, session.Snapshot.DocumentId))
        {
            var record = lease.Capture(session.Snapshot, "source.xlsq", new("ABC"), DateTimeOffset.UtcNow);
            Assert.Equal(session.SessionId, record.SessionId); Assert.Equal("ABC", record.SourceSha256); Assert.Empty(new WorkbookRecoveryStore(root).Scan().Items);
            Assert.Throws<IOException>(() => store.Start(session.SessionId, session.Snapshot.DocumentId));
        }
        var item = Assert.Single(store.Scan().Items); var recovered = store.Recover(item);
        Assert.Equal("中文", recovered.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        using var workspace = new WorkbookWorkspace(new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())), new Files());
        workspace.Recover(recovered); Assert.True(workspace.Document.IsDirty); Assert.Null(workspace.FilePath); Assert.False(workspace.Document.ScriptsEnabled);
        store.Discard(item); Assert.Empty(store.Scan().Items);
    }
    [Fact]
    public void DamagedLatestGenerationFallsBackAndDiscardDoesNotTouchUnjournaledFiles()
    {
        var root = Folder(); var store = new WorkbookRecoveryStore(root); var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        using (var lease = store.Start(session.SessionId, session.Snapshot.DocumentId))
        {
            session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("first")))]));
            lease.Capture(session.Snapshot, null, null, DateTimeOffset.UtcNow);
            session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("second")))]));
            lease.Capture(session.Snapshot, null, null, DateTimeOffset.UtcNow);
        }
        var folder = Path.Combine(root, session.SessionId.ToString("N")); File.WriteAllText(Path.Combine(folder, "draft1.qer"), "broken");
        var unrelated = Path.Combine(folder, "keep.txt"); File.WriteAllText(unrelated, "keep");
        var scan = store.Scan(); var item = Assert.Single(scan.Items); Assert.True(item.DamagedGeneration); Assert.NotEmpty(scan.Notices);
        Assert.Equal("first", store.Recover(item).Sheets[0].GetCell(0, 0)!.Value.TextValue);
        store.Discard(item); Assert.Equal("keep", File.ReadAllText(unrelated)); Assert.Empty(store.Scan().Items);
    }
    [Fact]
    public void FailedWritesPreserveLastCompleteGenerationAndNormalSaveCanRestartSameSession()
    {
        var root = Folder(); var storage = new FailingStorage(); var store = new WorkbookRecoveryStore(root, storage); var book = Book(Sheet("S", (0, 0, T("last good")))); var id = Guid.NewGuid();
        using (var lease = store.Start(id, book.DocumentId))
        {
            lease.Capture(book, null, null, DateTimeOffset.UtcNow); storage.Fail = true;
            Assert.Throws<IOException>(() => lease.Capture(book.WithRevision(1), null, null, DateTimeOffset.UtcNow));
        }
        var candidate = Assert.Single(store.Scan().Items); Assert.Equal(0, candidate.Metadata.Revision);
        Assert.Equal("last good", store.Recover(candidate).Sheets[0].GetCell(0, 0)!.Value.TextValue);
        store.Discard(candidate); storage.Fail = false;
        using var next = store.Start(id, book.DocumentId); next.Capture(book.WithRevision(2), null, null, DateTimeOffset.UtcNow); next.Discard();
        Assert.Empty(Directory.GetFiles(Path.Combine(root, id.ToString("N")), "*.qer"));
    }
    [Fact]
    public void StaleIdentityOrRevisionDoesNotOverwriteRecovery()
    {
        var store = new WorkbookRecoveryStore(Folder()); var book = Book(Sheet("S")).WithRevision(10); var id = Guid.NewGuid();
        using var lease = store.Start(id, book.DocumentId); lease.Capture(book, null, null, DateTimeOffset.UtcNow);
        Assert.Throws<InvalidOperationException>(() => lease.Capture(book.WithRevision(9), null, null, DateTimeOffset.UtcNow));
        Assert.Throws<InvalidOperationException>(() => lease.Capture(Book(Sheet("Other")), null, null, DateTimeOffset.UtcNow));
    }
    [Fact]
    public void FailedAtomicReplacementKeepsBothPreviousGenerationsAndSourceIsNeverOverwritten()
    {
        if (!OperatingSystem.IsWindows()) return;
        var root = Folder(); var source = Path.Combine(root, "original.xlsq"); File.WriteAllText(source, "changed externally");
        var store = new WorkbookRecoveryStore(root); var book = Book(Sheet("S")); var id = Guid.NewGuid();
        using (var lease = store.Start(id, book.DocumentId))
        {
            lease.Capture(book.WithRevision(1), source, new("old digest"), DateTimeOffset.UtcNow);
            lease.Capture(book.WithRevision(2), source, new("old digest"), DateTimeOffset.UtcNow);
            using var locked = new FileStream(Path.Combine(root, id.ToString("N"), "draft0.qer"), FileMode.Open, FileAccess.Read, FileShare.Read);
            Assert.Throws<IOException>(() => lease.Capture(book.WithRevision(3), source, new("old digest"), DateTimeOffset.UtcNow));
        }
        var item = Assert.Single(store.Scan().Items); Assert.Equal(2, item.Metadata.Revision); store.Recover(item);
        Assert.Equal("changed externally", File.ReadAllText(source));
    }
    [Theory]
    [InlineData("complete")]
    [InlineData("writing")]
    public async Task ProcessCrashReleasesLeaseAndRestoresOnlyCompleteGeneration(string stage)
    {
        var root = Folder(); var start = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "RecoveryFixture", OperatingSystem.IsWindows() ? "RecoveryFixture.exe" : "RecoveryFixture")) { UseShellExecute = false, CreateNoWindow = true };
        start.ArgumentList.Add(root); start.ArgumentList.Add(stage); using var child = Process.Start(start)!;
        try
        {
            var deadline = DateTime.UtcNow.AddSeconds(20);
            while (!File.Exists(Path.Combine(root, "ready"))) { Assert.False(child.HasExited); Assert.True(DateTime.UtcNow < deadline); await Task.Delay(30); }
            var store = new WorkbookRecoveryStore(root); Assert.Empty(store.Scan().Items);
            child.Kill(); await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            var scan = store.Scan();
            for (var retry = 0; scan.Items.Count == 0 && retry < 40; retry++) { await Task.Delay(50); scan = store.Scan(); }
            Assert.True(scan.Items.Count == 1, string.Join("\n", scan.Notices));
            var item = Assert.Single(scan.Items); Assert.Equal("complete", store.Recover(item).Sheets[0].GetCell(0, 0)!.Value.TextValue);
        }
        finally { if (!child.HasExited) { child.Kill(); await child.WaitForExitAsync(); } }
    }
    private sealed class FailingStorage : IWorkbookStorage
    {
        public bool Fail; private readonly TypedWorkbookStorage _storage = new();
        public WorkbookSnapshot Open(string path) => _storage.Open(path);
        public void Export(string path, WorkbookSnapshot snapshot) => _storage.Export(path, snapshot);
        public void Save(string path, WorkbookSnapshot snapshot) { if (Fail) { File.WriteAllText(path, "partial"); throw new IOException("Injected disk full"); } _storage.Save(path, snapshot); }
    }
    private sealed class Files : IWorkbookFileAccess
    {
        public OpenedWorkbook Open(string path) => new(new TypedWorkbookStorage().Open(path), FileVersions.Read(path));
        public FileVersion ReadVersion(string path) => FileVersions.Read(path);
    }
}
