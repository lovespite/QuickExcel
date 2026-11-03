using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

internal static class ThreadedCommentsProbe
{
    internal static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); var before = File.ReadAllBytes(source); var storage = new TypedWorkbookStorage();
        using var session = new DocumentSession(storage.Open(source));
        if (session.Snapshot.FileState!.PreservedAttachments.Length != 2) throw new InvalidDataException("Expected two preserved thread anchors.");
        var otherIssues = session.Snapshot.FileState.CompatibilityIssues;
        session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(12)))]));
        var path = Path.Combine(directory, "edited.xlsx"); storage.Save(path, session.Snapshot, discardUnsupported: otherIssues.Length > 0);
        var output = OpenXmlAudit.ReadParts(File.ReadAllBytes(path)); var input = OpenXmlAudit.ReadParts(before);
        var bundles = input.Where(p => p.Key.Contains("threadedComments/", StringComparison.Ordinal) || p.Key.Contains("persons/", StringComparison.Ordinal) || p.Key == "xl/comments1.xml")
            .Select(p => new { Part = p.Key, Preserved = output.TryGetValue(p.Key, out var saved) && p.Value.SequenceEqual(saved), Hash = OpenXmlAudit.Hash(p.Value) }).ToArray();
        var reopened = storage.Open(path); var unchanged = before.SequenceEqual(File.ReadAllBytes(source));
        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(path));
        var passed = bundles.Length == 3 && bundles.All(b => b.Preserved) && reopened.FileState!.PreservedAttachments.Length == 2 &&
            reopened.Sheets[0].GetCell(0, 0)!.Value == CellValue.FromNumber(12) && unchanged && audit.Issues.Count == 0;
        ProbeRunner.WriteNew(Path.Combine(directory, "thread-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { Passed = passed, SourceUnchanged = unchanged, OtherIssues = otherIssues, Bundles = bundles, SchemaIssues = audit.Issues }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Threaded comments: anchors={reopened.FileState!.PreservedAttachments.Length}; bundles preserved={bundles.All(b => b.Preserved)}; passed={passed}");
        return passed ? 0 : 2;
    }
}
