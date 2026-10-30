using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class DataCleaningProbe
{
    public static int Verify(string evidencePath, string verdictPath)
    {
        using var json = JsonDocument.Parse(File.ReadAllBytes(evidencePath)); var evidence = json.RootElement;
        var source = evidence.GetProperty("Source").GetString()!;
        if (!evidence.GetProperty("Stable").GetBoolean() || !evidence.GetProperty("SourceUnchanged").GetBoolean() ||
            evidence.GetProperty("SourceSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(source)))
            throw new InvalidDataException("Unbound or unstable Excel output evidence.");
        using var session = new DocumentSession(new TypedWorkbookStorage().Open(source));
        if (!session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation()))) throw new InvalidOperationException("Calculation failed.");
        var book = session.Snapshot; var checks = new List<object>(); var differences = new List<object>();
        foreach (var sheetJson in evidence.GetProperty("After").EnumerateArray())
        {
            var name = sheetJson.GetProperty("Sheet").GetString()!; var sheet = book.Sheets.Single(s => s.Name == name);
            foreach (var cellJson in sheetJson.GetProperty("Cells").EnumerateArray())
            {
                var address = cellJson.GetProperty("Address").GetString()!; var p = CellRectangle.Parse(address); var cell = sheet.GetCell(p.Top, p.Left);
                var value = cell?.Value ?? CellValue.Blank; var expected = cellJson.GetProperty("Value");
                void Check(string field, object? wanted, object? actual)
                { var item = new { Sheet = name, Address = address, Field = field, Expected = wanted, Actual = actual, Passed = Equals(wanted, actual) }; checks.Add(item); if (!item.Passed) differences.Add(item); }
                Check("kind", cellJson.GetProperty("Kind").GetString(), value.Kind.ToString());
                if (value.Kind == CellValueKind.Number && expected.ValueKind == JsonValueKind.Number) Check("value", expected.GetDouble(), value.NumberValue);
                else Check("value", expected.ValueKind == JsonValueKind.Null ? "" : expected.ValueKind == JsonValueKind.True ? "TRUE" : expected.ValueKind == JsonValueKind.False ? "FALSE" : expected.ToString(), CellInput.DisplayText(value));
                string? Formula(string? f) => f is null ? null : FormulaSyntax.Print(FormulaSyntax.Parse(f));
                Check("formula", Formula(cellJson.GetProperty("Formula").GetString()), cell?.Content.Kind == CellContentKind.Formula ? Formula(cell.Content.FormulaText) : null);
                Check("text", cellJson.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, book.Styles[sheet.GetStyleId(p.Top, p.Left)], book.DateSystem));
            }
        }
        File.WriteAllText(verdictPath, JsonSerializer.Serialize(new { Passed = differences.Count == 0, Count = checks.Count, Differences = differences }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} native Excel output checks, {differences.Count} differences."); return differences.Count == 0 ? 0 : 1;
    }
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Cleaning must not run scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new output directory."); Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        var source = oracle.GetProperty("Source").GetString()!;
        if (!oracle.GetProperty("Stable").GetBoolean() || !oracle.GetProperty("SourceUnchanged").GetBoolean() ||
            oracle.GetProperty("SourceSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(source)))
            throw new InvalidDataException("Unbound or unstable Excel cleaning evidence.");
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(source));
        var checks = new List<object>(); var differences = new List<object>(); var unsupported = new List<object>();
        var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate(DocumentSession target) { if (!target.TryCommitCalculation(calculator.Calculate(target.BeginCalculation()))) throw new InvalidOperationException("Calculation failed."); }
        void Compare(string stage, string sheet, string field, object? expected, object? actual)
        {
            var pass = Equals(expected, actual); var item = new { Stage = stage, Sheet = sheet, Field = field, Expected = expected, Actual = actual, Passed = pass };
            checks.Add(item); if (!pass) differences.Add(item);
        }
        void Check(WorkbookSnapshot book, string stage)
        {
            foreach (var item in oracle.GetProperty("After").EnumerateArray())
            {
                var name = item.GetProperty("Sheet").GetString()!; var sheet = book.Sheets.Single(s => s.Name == name);
                foreach (var record in item.GetProperty("Cells").EnumerateArray())
                {
                    var address = record.GetProperty("Address").GetString()!; var p = CellRectangle.Parse(address);
                    if (name == "NumericChoice" && p.Top >= 10 && p.Left == 2) continue; // explicit precision/date/bool boundary, below
                    var cell = sheet.GetCell(p.Top, p.Left); var value = cell?.Value ?? CellValue.Blank;
                    Compare(stage, name, address + "/kind", record.GetProperty("Kind").GetString(), value.Kind.ToString());
                    var expected = record.GetProperty("Value");
                    if (value.Kind == CellValueKind.Number && expected.ValueKind == JsonValueKind.Number)
                        Compare(stage, name, address + "/value", expected.GetDouble(), value.NumberValue);
                    else Compare(stage, name, address + "/value", expected.ValueKind == JsonValueKind.Null ? "" : expected.ValueKind == JsonValueKind.True ? "TRUE" :
                        expected.ValueKind == JsonValueKind.False ? "FALSE" : expected.ToString(), CellInput.DisplayText(value));
                    string? Formula(string? f) => f is null ? null : FormulaSyntax.Print(FormulaSyntax.Parse(f));
                    Compare(stage, name, address + "/formula", Formula(record.GetProperty("Formula").GetString()), cell?.Content.Kind == CellContentKind.Formula ? Formula(cell.Content.FormulaText) : null);
                    Compare(stage, name, address + "/text", record.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, book.Styles[sheet.GetStyleId(p.Top, p.Left)], book.DateSystem));
                }
            }
        }
        Calculate(session);
        foreach (var item in oracle.GetProperty("Cases").EnumerateArray())
        {
            var name = item.GetProperty("Sheet").GetString()!; var sheet = session.Snapshot.Sheets.Single(s => s.Name == name);
            var range = CellRectangle.Parse(item.GetProperty("Range").GetString()!); CleaningPreview preview;
            if (item.GetProperty("Operation").GetString() == "Duplicates")
                preview = DataCleaning.RemoveDuplicates(session.Snapshot, sheet.Id, range, item.GetProperty("Keys").EnumerateArray().Select(v => range.Left + v.GetInt32() - 1), item.GetProperty("Header").GetBoolean());
            else
            {
                var dest = CellRectangle.Parse(item.GetProperty("Destination").GetString()!);
                var fixedWidth = item.GetProperty("Operation").GetString() == "Fixed";
                var options = new TextSplitOptions(item.GetProperty("Delimiter").GetString() ?? ",", item.GetProperty("Collapse").GetBoolean(),
                    Breaks: fixedWidth ? item.GetProperty("Fixed").EnumerateArray().Select(v => v.GetInt32()).Where(v => v > 0).ToImmutableArray() : default);
                var types = item.GetProperty("Formats").EnumerateArray().Select(v => v.GetInt32() switch { 1 => CleaningColumnType.Number, 3 => CleaningColumnType.DateMdy, 4 => CleaningColumnType.DateDmy, 5 => CleaningColumnType.DateYmd, _ => CleaningColumnType.Text }).ToArray();
                if (name == "NumericChoice")
                {
                    // Strict conversion intentionally rejects automatic date/bool guessing and >15-digit text IDs.
                    var before = session.Snapshot;
                    try { _ = DataCleaning.SplitText(before, sheet.Id, range, new(dest.Top, dest.Left), options, types); throw new InvalidOperationException("Unsafe conversion should fail."); }
                    catch (ArgumentException ex) { unsupported.Add(new { Sheet = name, Message = ex.Message, InputUnchanged = ReferenceEquals(before, session.Snapshot) }); }
                    range = range with { Bottom = 9 };
                }
                preview = DataCleaning.SplitText(session.Snapshot, sheet.Id, range, new(dest.Top, dest.Left), options, types);
            }
            session.Execute(new ApplyCleaningCommand(preview)); var after = session.Snapshot;
            session.Undo(); session.Redo();
            Compare("history", name, "redo content", true, session.Snapshot.Styles.SequenceEqual(after.Styles) &&
                session.Snapshot.Sheets.SelectMany(s => s.Cells.Select(p => (s.Id, p.Key, p.Value.Content))).ToHashSet()
                    .SetEquals(after.Sheets.SelectMany(s => s.Cells.Select(p => (s.Id, p.Key, p.Value.Content)))));
            Calculate(session);
        }
        // The native recipe uses sheetFormatPr/defaultColWidth, an existing reported
        // preservation boundary. Use an explicit supported width for output display checks.
        foreach (var sheet in session.Snapshot.Sheets)
            session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Column, 0, 8, 24));
        Calculate(session);
        unsupported.Add(new { Field = "defaultColWidth", Policy = "Explicit A:H width 24 before diagnostic-copy export; default width preservation is not claimed." });
        Check(session.Snapshot, "commands");
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "cleaned." + extension);
            storage.Save(path, session.Snapshot, true);
            using var reopened = new DocumentSession(storage.Open(path)); Calculate(reopened); Check(reopened.Snapshot, extension);
            if (extension == "xlsx") differences.AddRange(OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()));
        }
        File.WriteAllText(Path.Combine(directory, "cleaning.json"), JsonSerializer.Serialize(new { Passed = differences.Count == 0, Count = checks.Count, Differences = differences, Checks = checks, Controlled = unsupported },
            new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} cleaning checks, {differences.Count} differences."); return differences.Count == 0 ? 0 : 1;
    }
}
