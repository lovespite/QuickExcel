using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class OutlineSubtotalProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Outline tests must not run scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new output directory."); Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var root = json.RootElement;
        var source = root.GetProperty("Source").GetString()!;
        if (!root.GetProperty("SourceUnchanged").GetBoolean() || root.GetProperty("SourceSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(source))) throw new InvalidDataException("Unbound source evidence.");
        var differences = new List<object>(); var count = 0; var storage = new TypedWorkbookStorage();
        void Check(string stage, string sheet, string field, object? expected, object? actual)
        {
            count++; var pass = expected is double x && actual is double y ? Math.Abs(x - y) <= Math.Max(1, Math.Abs(x)) * 1e-12 : Equals(expected, actual);
            if (!pass) differences.Add(new { Stage = stage, Sheet = sheet, Field = field, Expected = expected, Actual = actual });
        }
        void Compare(WorkbookSnapshot book, JsonElement expectedStage, string stage)
        {
            foreach (var expectedSheet in expectedStage.GetProperty("Sheets").EnumerateArray())
            {
                var name = expectedSheet.GetProperty("Sheet").GetString()!; var sheet = book.Sheets.Single(s => s.Name == name);
                Check(stage, name, "summaryBelow", expectedSheet.GetProperty("SummaryBelow").GetBoolean(), sheet.Data.Outline?.SummaryBelow ?? true);
                Check(stage, name, "summaryRight", expectedSheet.GetProperty("SummaryRight").GetBoolean(), sheet.Data.Outline?.SummaryRight ?? true);
                var hidden = WorksheetFilterEngine.Evaluate(book, sheet.Id);
                foreach (var row in expectedSheet.GetProperty("Rows").EnumerateArray())
                {
                    var r = row.GetProperty("Index").GetInt32() - 1; var meta = sheet.Rows.GetValueOrDefault(r);
                    Check(stage, name, $"row{r + 1}/level", row.GetProperty("Level").GetDouble(), (double)(meta?.OutlineLevel ?? 0) + 1);
                    Check(stage, name, $"row{r + 1}/hidden", row.GetProperty("Hidden").GetBoolean(), meta?.Hidden == true || hidden.Contains(r));
                    if (row.TryGetProperty("Collapsed", out var collapsed)) Check(stage, name, $"row{r + 1}/collapsed", collapsed.GetBoolean(), meta?.Collapsed ?? false);
                }
                foreach (var column in expectedSheet.GetProperty("Columns").EnumerateArray())
                {
                    var c = column.GetProperty("Index").GetInt32() - 1; var meta = sheet.Columns.GetValueOrDefault(c);
                    Check(stage, name, $"col{c + 1}/level", column.GetProperty("Level").GetDouble(), (double)(meta?.OutlineLevel ?? 0) + 1);
                    Check(stage, name, $"col{c + 1}/hidden", column.GetProperty("Hidden").GetBoolean(), meta?.Hidden ?? false);
                    if (column.TryGetProperty("Collapsed", out var collapsed)) Check(stage, name, $"col{c + 1}/collapsed", collapsed.GetBoolean(), meta?.Collapsed ?? false);
                }
                foreach (var expected in expectedSheet.GetProperty("Cells").EnumerateArray())
                {
                    var address = expected.GetProperty("Address").GetString()!; var p = CellRectangle.Parse(address); var cell = sheet.GetCell(p.Top, p.Left); var value = cell?.Value ?? CellValue.Blank;
                    Check(stage, name, address + "/kind", expected.GetProperty("Kind").GetString(), value.Kind.ToString());
                    if (value.Kind == CellValueKind.Number && expected.GetProperty("Value").ValueKind == JsonValueKind.Number)
                        Check(stage, name, address + "/value", expected.GetProperty("Value").GetDouble(), value.NumberValue);
                    else Check(stage, name, address + "/value", expected.GetProperty("Value").ValueKind == JsonValueKind.True ? "TRUE" : expected.GetProperty("Value").ValueKind == JsonValueKind.False ? "FALSE" : expected.GetProperty("Value").ToString(), CellInput.DisplayText(value));
                    string? Formula(string? f) => f is null ? null : FormulaSyntax.Print(FormulaSyntax.Parse(f));
                    Check(stage, name, address + "/formula", Formula(expected.GetProperty("Formula").GetString()), cell?.Content.Kind == CellContentKind.Formula ? Formula(cell.Content.FormulaText) : null);
                    if (cell?.Content.Kind == CellContentKind.Formula) Check(stage, name, address + "/calculated", true, cell.Calculation?.State is CalculationState.Fresh or CalculationState.Error);
                }
            }
        }
        void Calculate(DocumentSession session) { if (!session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation()))) throw new InvalidOperationException("Calculation failed."); }
        foreach (var stage in root.GetProperty("Stages").EnumerateArray())
        {
            var name = stage.GetProperty("Name").GetString()!; var path = stage.GetProperty("Path").GetString()!;
            if (!stage.GetProperty("Stable").GetBoolean() || stage.GetProperty("Sha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(path))) throw new InvalidDataException("Unbound stage.");
            if (name == "output")
            {
                using var input = new DocumentSession(storage.Open(source)); Calculate(input); Compare(input.Snapshot, stage, "Excel-output-vs-QuickExcel-source");
            }
            using var session = new DocumentSession(storage.Open(path)); Calculate(session); Compare(session.Snapshot, stage, name + "/import");
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, name + "." + extension); storage.Save(output, session.Snapshot, true);
                using var reopened = new DocumentSession(storage.Open(output)); Calculate(reopened); Compare(reopened.Snapshot, stage, name + "/" + extension);
                if (extension == "xlsx") differences.AddRange(OpenXmlAudit.Inspect(File.ReadAllBytes(output)).Issues.Select(i => i.ToString()));
            }
        }
        // Apply independently specified operations to the source, then compare with Excel's
        // command results. Reading the already-generated file alone would not verify editing.
        if (root.GetProperty("Stages").EnumerateArray().Any(s => s.GetProperty("Name").GetString() == "initial"))
        {
            using var commands = new DocumentSession(storage.Open(source)); Calculate(commands);
            foreach (var stage in root.GetProperty("Stages").EnumerateArray())
            {
                var name = stage.GetProperty("Name").GetString()!;
                var values = commands.Snapshot.Sheets.Single(s => s.Name == "Subtotals");
                if (name == "manual")
                { commands.Execute(new SetAxisMetadataCommand(values.Id, SheetAxis.Row, 1, 1, hidden: true)); commands.Execute(new SetAxisMetadataCommand(values.Id, SheetAxis.Column, 3, 1, hidden: true)); }
                if (name == "filtered") commands.Execute(new SetAutoFilterCommand(values.Id, new(new(0, 0, 11, 1), [new(1, new(values: ["yes"]))])));
                if (name is "collapsed" or "expanded")
                {
                    foreach (var sheet in commands.Snapshot.Sheets.Where(s => s.Name.StartsWith("Outline")))
                    {
                        if (name == "collapsed") commands.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 2, 4, 2, true));
                        commands.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 1, 11, 1, name == "collapsed"));
                        commands.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Column, 1, 6, 1, name == "collapsed"));
                    }
                }
                if (name == "generated")
                {
                    foreach (var test in root.GetProperty("Cases").EnumerateArray())
                    {
                        var sheet = commands.Snapshot.Sheets.Single(s => s.Name == test.GetProperty("Sheet").GetString());
                        var range = CellRectangle.Parse(test.GetProperty("Range").GetString()!);
                        var preview = SubtotalCommands.Preview(commands.Snapshot, sheet.Id, range, range.Left + test.GetProperty("Group").GetInt32() - 1,
                            test.GetProperty("Totals").EnumerateArray().Select(c => range.Left + c.GetInt32() - 1), test.GetProperty("Function").GetInt32());
                        commands.Execute(new ApplySubtotalsCommand(preview)); commands.Undo(); commands.Redo(); Calculate(commands);
                    }
                }
                Calculate(commands); Compare(commands.Snapshot, stage, name + "/commands");
                var path = Path.Combine(directory, "command-" + name + ".xlsx"); storage.Save(path, commands.Snapshot, true);
                using var reopened = new DocumentSession(storage.Open(path)); Calculate(reopened); Compare(reopened.Snapshot, stage, name + "/command-reopen");
                differences.AddRange(OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()));
            }
        }
        File.WriteAllText(Path.Combine(directory, "outline-subtotal.json"), JsonSerializer.Serialize(new { Passed = differences.Count == 0, Count = count, Differences = differences }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{count} outline/subtotal checks, {differences.Count} differences."); return differences.Count == 0 ? 0 : 1;
    }
}
