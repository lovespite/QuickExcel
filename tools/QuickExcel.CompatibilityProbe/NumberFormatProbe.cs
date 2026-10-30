using System.Globalization;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class NumberFormatProbe
{
    public static int Run(string reference, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Evidence directory already exists.");
        Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(reference));
        var source = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(reference))!, "excel-source.xlsx");
        var bytes = File.ReadAllBytes(source);
        var audit = OpenXmlAudit.Inspect(bytes);
        var storage = new TypedWorkbookStorage();
        var book = storage.Open(source);
        var results = new List<object>();
        var failures = new List<string>();
        void Check(WorkbookSnapshot snapshot, string mode)
        {
            var sheet = snapshot.Sheets[0];
            foreach (var item in json.RootElement.GetProperty("Cases").EnumerateArray())
            {
                var row = item.GetProperty("Row").GetInt32() - 1;
                var value = sheet.GetCell(row, 0)!.Value;
                var display = CellDisplayFormatter.FormatDisplay(value, snapshot.Styles[sheet.GetStyleId(row, 0)], snapshot.DateSystem, CultureInfo.InvariantCulture);
                var expected = item.GetProperty("Text").GetString()!;
                if (expected.Length > 0 && expected.All(c => c == '#')) expected = "########";
                if (display.FillOffset >= 0 && expected.Length >= display.Text.Length)
                    expected = expected.Remove(display.FillOffset, expected.Length - display.Text.Length);
                var passed = display.Text == expected;
                results.Add(new { Mode = mode, Row = row + 1, Actual = display.Text, Expected = expected, Passed = passed });
                if (!passed) failures.Add($"{mode}: A{row + 1}");
            }
        }
        Check(book, "import");
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 1), CellContent.FromLiteral(CellValue.FromText("edit")))]));
        // Explicit diagnostic copy: the rest of E1 still owns theme/view/property fidelity.
        var copy = Path.Combine(directory, "edited.diagnostic-copy.xlsq");
        storage.Save(copy, session.Snapshot, discardUnsupported: true);
        var output = OpenXmlAudit.ReadParts(File.ReadAllBytes(copy))["workbook.xlsx"];
        ProbeRunner.WriteNew(Path.Combine(directory, "edited.diagnostic-copy.xlsx"), output);
        Check(storage.Open(copy), "save-reopen");
        var outputAudit = OpenXmlAudit.Inspect(output);
        ProbeRunner.WriteNew(Path.Combine(directory, "display-report.json"), JsonSerializer.SerializeToUtf8Bytes(new
        {
            Excel = json.RootElement.Clone(), Source = audit, Output = outputAudit,
            SourceUnchanged = bytes.SequenceEqual(File.ReadAllBytes(source)),
            DeferredCompatibilityIssues = book.FileState!.CompatibilityIssues, Results = results, Failures = failures
        }, ProbeRunner.JsonOptions));
        Console.WriteLine($"{results.Count} display checks; {failures.Count} failures; {outputAudit.Issues.Count} output schema issues.");
        return failures.Count == 0 && outputAudit.Issues.Count == 0 ? 0 : 2;
    }
}
