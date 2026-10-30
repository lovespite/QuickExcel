using System.Globalization;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class FillProbe
{
    private static readonly string[] Ids = ["down", "right", "handle-down", "handle-left", "clear-content", "clear-format", "numbers", "numbers-negative", "numbers-right",
        "days-1900", "days-1904", "months-end", "months-fictitious", "months-fraction", "months-back", "years-leap", "years-1904"];
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Fill probe cannot execute scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Evidence directory must be new.");
        Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        if (oracle.GetProperty("Origin").GetString() != "independent-excel-oracle" ||
            !oracle.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids)) throw new InvalidDataException("Incomplete independent fill oracle.");
        var paths = new List<object>(); var failures = 0;
        foreach (var expected in oracle.GetProperty("Cases").EnumerateArray())
        {
            var id = expected.GetProperty("Id").GetString()!; var sourcePath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, id + "-source.xlsx");
            var source = File.ReadAllBytes(sourcePath); var hash = OpenXmlAudit.Hash(source);
            if (!hash.Equals(expected.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase) || !expected.GetProperty("Stable").GetBoolean() ||
                !expected.GetProperty("SourceUnchanged").GetBoolean()) throw new InvalidDataException("Oracle/source mismatch.");
            var storage = new TypedWorkbookStorage(); var original = storage.Open(sourcePath); using var session = new DocumentSession(original);
            var sheetId = original.Sheets[0].Id; var definition = expected.GetProperty("Definition");
            DocumentCommand command = id switch
            {
                "down" => new FillCellsCommand(sheetId, new(0, 0, 0, 1), new(0, 0, 5, 1)),
                "right" => new FillCellsCommand(sheetId, new(0, 0, 1, 0), new(0, 0, 1, 4)),
                "handle-down" => new FillCellsCommand(sheetId, new(0, 0, 1, 1), new(0, 0, 5, 1)),
                "handle-left" => new FillCellsCommand(sheetId, new(0, 2, 1, 3), new(0, 0, 1, 3)),
                "clear-content" => new ClearCellsCommand(Enumerable.Range(0, 6).SelectMany(r => Enumerable.Range(0, 2).Select(c => new CellAddress(sheetId, r, c)))),
                "clear-format" => new ClearCellFormatsCommand(Enumerable.Range(0, 6).SelectMany(r => Enumerable.Range(0, 2).Select(c => new CellAddress(sheetId, r, c)))),
                _ => new FillSeriesCommand(sheetId, id == "numbers-right" ? new(0, 0, 0, 5) : new(0, 0, 7, 0),
                    Enum.Parse<CellSeriesKind>(definition.GetProperty("Kind").GetString()!), definition.GetProperty("Step").GetDouble())
            };
            session.Execute(command); var edited = session.Snapshot;
            var undo = session.Undo() && SameContent(original, session.Snapshot) && !session.CanUndo && session.Redo() && SameContent(edited, session.Snapshot);
            var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
            if (result.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)) || !session.TryCommitCalculation(result)) throw new InvalidOperationException("Fill requires actual recalculation.");
            var checks = Compare(session.Snapshot, expected.GetProperty("Reopened"), "calculated").ToList(); var rejections = new List<string>(); string? outputHash = null;
            var issues = new List<string>();
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, id + "." + extension);
                try { storage.Save(output, session.Snapshot); }
                catch (WorkbookCompatibilityException ex) { rejections.Add(extension + ": " + ex.Message); storage.Save(output, session.Snapshot, discardUnsupported: true); }
                checks.AddRange(Compare(storage.Open(output), expected.GetProperty("Reopened"), extension));
                if (extension == "xlsx") { var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(output)); outputHash = audit.Sha256; issues.AddRange(audit.Issues.Select(i => i.ToString()!)); }
            }
            var unchanged = source.SequenceEqual(File.ReadAllBytes(sourcePath)); var passed = undo && unchanged && issues.Count == 0 && checks.All(c => c.Passed);
            if (!passed) failures++;
            paths.Add(new { Id = id, Passed = passed, UndoRedo = undo, SourceUnchanged = unchanged, OutputSha256 = outputHash, SaveRejections = rejections, SchemaIssues = issues, Checks = checks });
            Console.WriteLine($"{id}: {(passed ? "PASS" : "FAIL")}; checks={checks.Count}; failures={checks.Count(c => !c.Passed)}; schema={issues.Count}");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "fill-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), Failures = failures, Paths = paths }, ProbeRunner.JsonOptions));
        return failures == 0 ? 0 : 2;
    }
    private static bool SameContent(WorkbookSnapshot a, WorkbookSnapshot b) => a.Sheets.Zip(b.Sheets).All(pair => pair.First.Cells.Count == pair.Second.Cells.Count &&
        pair.First.Cells.All(p => pair.Second.Cells.TryGetValue(p.Key, out var cell) && p.Value.Content == cell.Content));
    private static IEnumerable<BusinessBaselineProbe.Check> Compare(WorkbookSnapshot actual, JsonElement expected, string stage)
    {
        var system = expected.GetProperty("Date1904").GetBoolean() ? ExcelDateSystem.Mac1904 : ExcelDateSystem.Windows1900;
        yield return new(stage, "DateSystem", system.ToString(), actual.DateSystem.ToString(), actual.DateSystem == system);
        foreach (var item in expected.GetProperty("Cells").EnumerateArray())
        {
            var row = item.GetProperty("Row").GetInt32(); var column = item.GetProperty("Column").GetInt32(); var location = $"{stage}/{row}/{column}";
            var cell = actual.Sheets[0].GetCell(row, column); var value = cell?.Value ?? CellValue.Blank; var kind = item.GetProperty("Kind").GetString();
            yield return new(location, "Kind", kind, value.Kind.ToString(), kind == value.Kind.ToString());
            var raw = item.GetProperty("Value"); var wanted = raw.ValueKind switch { JsonValueKind.Null => "", JsonValueKind.Number => raw.GetDouble().ToString("R", CultureInfo.InvariantCulture),
                JsonValueKind.True => "TRUE", JsonValueKind.False => "FALSE", _ => raw.GetString() };
            var observed = CellInput.DisplayText(value); yield return new(location, "Value", wanted, observed, wanted == observed);
            var ef = BusinessBaselineProbe.CanonicalFormula(item.GetProperty("Formula").GetString()); var af = cell?.Content.Kind == CellContentKind.Formula ? BusinessBaselineProbe.CanonicalFormula(cell.Content.FormulaText) : null;
            yield return new(location, "Formula", ef, af, ef == af);
        }
    }
    public static int Verify(string oraclePath, string actualPath, string modelPath, string outputPath)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); using var actual = JsonDocument.Parse(File.ReadAllBytes(actualPath)); using var model = JsonDocument.Parse(File.ReadAllBytes(modelPath));
        if (actual.RootElement.GetProperty("Origin").GetString() != "excel-output-verification" ||
            model.RootElement.GetProperty("OracleSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)) ||
            !actual.RootElement.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids) ||
            !model.RootElement.GetProperty("Paths").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids)) throw new InvalidDataException("Incomplete or unbound fill evidence.");
        var failures = model.RootElement.GetProperty("Failures").GetInt32(); var paths = new List<object>();
        foreach (var observed in actual.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var id = observed.GetProperty("Id").GetString(); var expected = oracle.RootElement.GetProperty("Cases").EnumerateArray().Single(c => c.GetProperty("Id").GetString() == id);
            var path = model.RootElement.GetProperty("Paths").EnumerateArray().Single(c => c.GetProperty("Id").GetString() == id);
            var differences = EditMatrixVerification.Compare(expected.GetProperty("Reopened"), observed.GetProperty("Reopened"));
            var bound = string.Equals(path.GetProperty("OutputSha256").GetString(), observed.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase);
            var passed = bound && path.GetProperty("Passed").GetBoolean() && observed.GetProperty("SourceUnchanged").GetBoolean() && observed.GetProperty("Stable").GetBoolean() && differences.All(d => d.Classification != "unexpected");
            if (!passed) failures++; paths.Add(new { Id = id, Passed = passed, Bound = bound, Differences = differences });
        }
        ProbeRunner.WriteNew(outputPath, JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures, Paths = paths,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), VerificationSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(actualPath)), ModelSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)) }, ProbeRunner.JsonOptions));
        Console.WriteLine($"17 fill paths; failures={failures}"); return failures == 0 ? 0 : 2;
    }
}
