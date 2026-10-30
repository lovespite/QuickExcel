using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class PrintingProbe
{
    public static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory);
        var original = File.ReadAllBytes(source); var storage = new TypedWorkbookStorage(); var book = storage.Open(source);
        using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 5, 3), CellContent.FromLiteral(CellValue.FromNumber(99)))]));
        session.Execute(new SetCellStyleCommand([new(id, 5, 3)], new(Bold: true)));
        var output = Path.Combine(directory, "edited.diagnostic-copy.xlsx"); storage.Save(output, session.Snapshot, discardUnsupported: true);
        var reopened = storage.Open(output); var same = book.Sheets.Zip(reopened.Sheets).All(p => p.First.Printing.ContentEquals(p.Second.Printing));
        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(output));
        session.Execute(new RenameSheetCommand(id, "改名'打印")); session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1)); session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 0, 1));
        var inserted = Path.Combine(directory, "inserted.diagnostic-copy.xlsx"); storage.Save(inserted, session.Snapshot, discardUnsupported: true);
        var insertAudit = OpenXmlAudit.Inspect(File.ReadAllBytes(inserted));
        using var deleteSession = new DocumentSession(book);
        string? deletionRejection = null;
        try { deleteSession.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 2, delete: true)); }
        catch (NotSupportedException e) { deletionRejection = e.Message; }
        var deletionProtected = deletionRejection is not null && ReferenceEquals(book, deleteSession.Snapshot) && !deleteSession.CanUndo;
        var unchanged = original.SequenceEqual(File.ReadAllBytes(source));
        var report = new { SourceUnchanged = unchanged, PrintingEqual = same, Source = book.Sheets.Select(s => new { s.Name, s.Printing }),
            AfterInsert = session.Snapshot.Sheets.Select(s => new { s.Name, s.Printing }), RemainingIssues = book.FileState!.CompatibilityIssues,
            SchemaIssues = audit.Issues, InsertSchemaIssues = insertAudit.Issues, DeletionProtected = deletionProtected, DeletionRejection = deletionRejection };
        File.WriteAllText(Path.Combine(directory, "printing-report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"PrintingEqual={same}; SourceUnchanged={unchanged}; SchemaIssues={audit.Issues.Count}/{insertAudit.Issues.Count}; DeletionProtected={deletionProtected}");
        return same && unchanged && audit.Issues.Count == 0 && insertAudit.Issues.Count == 0 && deletionProtected ? 0 : 1;
    }
}
