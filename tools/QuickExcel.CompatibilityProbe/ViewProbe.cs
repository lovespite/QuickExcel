using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class ViewProbe
{
    public static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory);
        var original = File.ReadAllBytes(source); var storage = new TypedWorkbookStorage(); var book = storage.Open(source);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.ActiveSheetId, 0, 0), CellContent.FromLiteral(CellValue.FromText("edited view fixture")))]));
        var output = Path.Combine(directory, "edited.diagnostic-copy.xlsx");
        storage.Save(output, session.Snapshot, book.FileState!.CompatibilityIssues.Length > 0);
        var saved = storage.Open(output);
        var viewsEqual = book.Window == saved.Window && book.GetSheet(book.ActiveSheetId).Name == saved.GetSheet(saved.ActiveSheetId).Name &&
            book.Sheets.Zip(saved.Sheets).All(p => p.First.View == p.Second.View && p.First.Visibility == p.Second.Visibility);
        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(output));
        session.Execute(new ShiftAxisCommand(book.ActiveSheetId, SheetAxis.Row, 0, 1));
        var inserted = Path.Combine(directory, "row-inserted.diagnostic-copy.xlsx");
        storage.Save(inserted, session.Snapshot, book.FileState.CompatibilityIssues.Length > 0);
        var insertAudit = OpenXmlAudit.Inspect(File.ReadAllBytes(inserted));
        ProbeRunner.WriteNew(Path.Combine(directory, "views-report.json"), JsonSerializer.SerializeToUtf8Bytes(new
        {
            SourceUnchanged = original.SequenceEqual(File.ReadAllBytes(source)), ViewsEqual = viewsEqual,
            ActiveSheet = saved.GetSheet(saved.ActiveSheetId).Name,
            Sheets = saved.Sheets.Select(s => new { s.Name, s.Visibility, s.View }), Output = audit, Inserted = insertAudit,
            RemainingIssues = book.FileState.CompatibilityIssues
        }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Views equal={viewsEqual}; output/inserted schema issues={audit.Issues.Count}/{insertAudit.Issues.Count}.");
        return viewsEqual && audit.Issues.Count + insertAudit.Issues.Count == 0 && original.SequenceEqual(File.ReadAllBytes(source)) ? 0 : 2;
    }
}
