using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

// Test-only host: the parent kills it while the session lease and optional partial write are open.
if (args.Length != 2) return 2;
var root = Path.GetFullPath(args[0]); var stage = args[1];
var sheet = new WorksheetSnapshot(SheetId.New(), "S");
using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
var storage = new FaultStorage(root);
var store = new WorkbookRecoveryStore(root, storage);
using var lease = store.Start(session.SessionId, session.Snapshot.DocumentId);
session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("complete")))]));
lease.Capture(session.Snapshot, null, null, DateTimeOffset.UtcNow);
if (stage == "writing")
{
    storage.Pause = true;
    session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("incomplete")))]));
    lease.Capture(session.Snapshot, null, null, DateTimeOffset.UtcNow);
}
else File.WriteAllText(Path.Combine(root, "ready"), "ready");
await Task.Delay(Timeout.InfiniteTimeSpan); return 0;

sealed class FaultStorage(string root) : IWorkbookStorage
{
    public bool Pause;
    private readonly TypedWorkbookStorage _storage = new();
    public WorkbookSnapshot Open(string path) => _storage.Open(path);
    public void Export(string path, WorkbookSnapshot snapshot) => _storage.Export(path, snapshot);
    public void Save(string path, WorkbookSnapshot snapshot)
    {
        if (!Pause) { _storage.Save(path, snapshot); return; }
        using var file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None); file.Write([1, 2, 3]); file.Flush(true);
        File.WriteAllText(Path.Combine(root, "ready"), "writing"); Thread.Sleep(Timeout.Infinite);
    }
}
