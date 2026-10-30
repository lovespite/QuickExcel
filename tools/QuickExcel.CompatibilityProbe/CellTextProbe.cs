using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class CellTextProbe
{
    public static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); var bytes = File.ReadAllBytes(source); var storage = new TypedWorkbookStorage(); var book = storage.Open(source);
        using var session = new DocumentSession(book); var sheet = book.Sheets.Single(s => s.Name == "内容对照");
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 5), CellContent.FromLiteral(CellValue.FromText("已编辑")))]));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 1), CellContent.FromLiteral(CellValue.FromText("修改后的链接文字"), sheet.GetStyleId(0, 1))),
            new(new(sheet.Id, 1, 0), CellContent.FromLiteral(CellValue.FromText("修改批注所在格"), sheet.GetStyleId(1, 0)))]));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 5)], new(Bold: true)));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Bold: true)));
        var output = Path.Combine(directory, "edited.diagnostic-copy.xlsx"); storage.Save(output, session.Snapshot, discardUnsupported: true);
        var reopened = storage.Open(output); var saved = reopened.Sheets.Single(s => s.Name == sheet.Name);
        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(output));
        var equal = sheet.Annotations.ContentEquals(saved.Annotations) && session.Snapshot.GetSheet(sheet.Id).Cells.Where(p => p.Value.Content.RichText is not null)
            .All(p => saved.GetCell(p.Key.Row, p.Key.Column)?.Content.RichText?.Equals(p.Value.Content.RichText) == true);
        var originalUnchanged = bytes.SequenceEqual(File.ReadAllBytes(source));
        File.WriteAllText(Path.Combine(directory, "cell-text-report.json"), JsonSerializer.Serialize(new
        {
            RichCells = sheet.Cells.Where(p => p.Value.Content.RichText is not null).Select(p => new { Cell = p.Key, p.Value.Content.RichText }).ToArray(),
            Links = sheet.Annotations.Links.Select(p => new { Cell = p.Key, Link = p.Value }).ToArray(),
            Notes = sheet.Annotations.Notes.Select(p => new { Cell = p.Key, Note = p.Value }).ToArray(),
            NoteDrawingBytes = sheet.Annotations.NoteDrawing.Length, ModelEqual = equal, OriginalUnchanged = originalUnchanged,
            SourceDiagnostics = book.FileState!.CompatibilityDiagnostics, OutputDiagnostics = reopened.FileState!.CompatibilityDiagnostics, SchemaIssues = audit.Issues
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"ModelEqual={equal}; OriginalUnchanged={originalUnchanged}; SchemaIssues={audit.Issues.Count}; VML={sheet.Annotations.NoteDrawing.Length}");
        return equal && originalUnchanged && audit.Issues.Count == 0 ? 0 : 1;
    }
}
