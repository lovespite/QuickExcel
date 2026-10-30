using System.Globalization;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class ClipboardProbe
{
    public static readonly string[] Modes = ["all", "all-sheet", "values", "formulas", "formats", "transpose", "transpose-mixed", "transpose-qualified",
        "cut-overlap", "cut-sheet", "rich-all", "rich-values", "rich-formulas", "rich-formats"];
    private static readonly string[] Rejected = ["rich-values", "rich-formulas", "rich-formats"];
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Clipboard probe cannot run scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Evidence directory must be new.");
        Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        if (oracle.GetProperty("Origin").GetString() != "independent-excel-oracle" ||
            !oracle.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Mode").GetString()).SequenceEqual(Modes))
            throw new InvalidDataException("Expected complete independent Excel clipboard oracle.");
        var reports = new List<object>(); var failures = 0;
        foreach (var expected in oracle.GetProperty("Cases").EnumerateArray())
        {
            var mode = expected.GetProperty("Mode").GetString()!;
            var sourcePath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, mode + "-source.xlsx");
            var bytes = File.ReadAllBytes(sourcePath); var hash = OpenXmlAudit.Hash(bytes);
            if (!hash.Equals(expected.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase) ||
                !expected.GetProperty("Stable").GetBoolean() || !expected.GetProperty("SourceUnchanged").GetBoolean())
                throw new InvalidDataException("Unstable oracle or source fingerprint mismatch.");
            var storage = new TypedWorkbookStorage(); var original = storage.Open(sourcePath);
            using var session = new DocumentSession(original);
            var payload = WorkbookClipboard.Capture(original, original.Sheets[0].Id, new(0, 0, 1, 2), mode.StartsWith("cut-", StringComparison.Ordinal));
            var target = mode == "cut-overlap" ? new CellRectangle(1, 1, 1, 1) : new(4, 4, 4, 4);
            var pasteMode = mode.EndsWith("values", StringComparison.Ordinal) ? CellPasteMode.Values : mode.EndsWith("formulas", StringComparison.Ordinal)
                ? CellPasteMode.Formulas : mode.EndsWith("formats", StringComparison.Ordinal) ? CellPasteMode.Formats : CellPasteMode.All;
            var command = new PasteCellsCommand(payload, original.Sheets[mode.EndsWith("-sheet", StringComparison.Ordinal) ? 1 : 0].Id,
                target, 0, pasteMode, mode.StartsWith("transpose", StringComparison.Ordinal));
            string? rejection = null; bool undo = false; var checks = new List<BusinessBaselineProbe.Check>(); string? outputHash = null;
            var saveRejections = new List<string>(); var schemaIssues = new List<string>();
            try { session.Execute(command); }
            catch (Exception ex) when (Rejected.Contains(mode) && ex is InvalidOperationException or NotSupportedException) { rejection = ex.Message; }
            bool passed;
            if (Rejected.Contains(mode)) passed = rejection is not null && ReferenceEquals(original, session.Snapshot) && !session.CanUndo;
            else
            {
                var edited = session.Snapshot;
                undo = session.Undo() && SameContent(original, session.Snapshot) && !session.CanUndo && session.Redo() && SameContent(edited, session.Snapshot);
                Calculate(session); checks.AddRange(CompareValues(session.Snapshot, expected.GetProperty("Sheets"), "calculated"));
                foreach (var extension in new[] { "xlsx", "xlsq" })
                {
                    var path = Path.Combine(directory, mode + "." + extension);
                    try { storage.Save(path, session.Snapshot); }
                    catch (WorkbookCompatibilityException ex) { saveRejections.Add(extension + ": " + ex.Message); storage.Save(path, session.Snapshot, discardUnsupported: true); }
                    checks.AddRange(CompareValues(storage.Open(path), expected.GetProperty("Reopened"), extension));
                    if (extension == "xlsx")
                    {
                        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(path)); outputHash = audit.Sha256;
                        schemaIssues.AddRange(audit.Issues.Select(i => i.ToString()!));
                    }
                }
                passed = undo && checks.All(c => c.Passed) && schemaIssues.Count == 0;
            }
            var sourceUnchanged = bytes.SequenceEqual(File.ReadAllBytes(sourcePath)); passed &= sourceUnchanged;
            if (!passed) failures++;
            reports.Add(new { Mode = mode, Passed = passed, Rejection = rejection, UndoRedo = undo, SourceSha256 = hash, SourceUnchanged = sourceUnchanged,
                OutputSha256 = outputHash, SaveRejections = saveRejections, SchemaIssues = schemaIssues, Checks = checks });
            Console.WriteLine($"{mode}: {(passed ? "PASS" : "FAIL")}, failures={checks.Count(c => !c.Passed)}, schema={schemaIssues.Count}, rejected={rejection is not null}");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "clipboard-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)),
            Failures = failures, Paths = reports, Scope = "11 implemented paths and 3 explicit rich-text selective-paste rejections. Excel output verification is separate." }, ProbeRunner.JsonOptions));
        return failures == 0 ? 0 : 2;
    }
    private static bool SameContent(WorkbookSnapshot a, WorkbookSnapshot b) => a.Sheets.Zip(b.Sheets).All(pair =>
        pair.First.Cells.Count == pair.Second.Cells.Count && pair.First.Cells.All(p => pair.Second.Cells.TryGetValue(p.Key, out var cell) && p.Value.Content == cell.Content));
    private static void Calculate(DocumentSession session)
    {
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        if (result.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)) || !session.TryCommitCalculation(result))
            throw new InvalidOperationException("Clipboard comparison requires actual built-in calculation.");
    }
    private static IEnumerable<BusinessBaselineProbe.Check> CompareValues(WorkbookSnapshot book, JsonElement expected, string stage)
    {
        foreach (var source in expected.EnumerateArray())
        foreach (var cell in source.GetProperty("Cells").EnumerateArray())
        {
            var index = source.GetProperty("Index").GetInt32(); var row = cell.GetProperty("Row").GetInt32(); var column = cell.GetProperty("Column").GetInt32();
            var actual = book.Sheets[index].GetCell(row, column); var value = actual?.Value ?? CellValue.Blank;
            var location = $"{stage}/{index}/{row}/{column}";
            var kind = cell.GetProperty("Kind").GetString();
            yield return new(location, "Kind", kind, value.Kind.ToString(), kind == value.Kind.ToString());
            var expectedValue = cell.GetProperty("Value");
            var wanted = expectedValue.ValueKind switch { JsonValueKind.Null => "", JsonValueKind.Number => expectedValue.GetDouble().ToString("R", CultureInfo.InvariantCulture),
                JsonValueKind.True => "True", JsonValueKind.False => "False", _ => expectedValue.GetString() };
            var observed = CellInput.DisplayText(value);
            yield return new(location, "Value", wanted, observed, wanted == observed);
            var ef = BusinessBaselineProbe.CanonicalFormula(cell.GetProperty("Formula").GetString());
            var af = actual?.Content.Kind == CellContentKind.Formula ? BusinessBaselineProbe.CanonicalFormula(actual.Content.FormulaText) : null;
            yield return new(location, "Formula", ef, af, ef == af);
        }
    }
    public static int Verify(string oraclePath, string actualPath, string modelPath, string verdictPath)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); using var actual = JsonDocument.Parse(File.ReadAllBytes(actualPath));
        using var model = JsonDocument.Parse(File.ReadAllBytes(modelPath));
        if (actual.RootElement.GetProperty("Origin").GetString() != "excel-output-verification" ||
            model.RootElement.GetProperty("OracleSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)) ||
            !model.RootElement.GetProperty("Paths").EnumerateArray().Select(c => c.GetProperty("Mode").GetString()).SequenceEqual(Modes) ||
            !actual.RootElement.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Mode").GetString()).SequenceEqual(Modes.Except(Rejected)))
            throw new InvalidDataException("Incomplete or unbound clipboard evidence.");
        var failures = model.RootElement.GetProperty("Failures").GetInt32(); var reports = new List<object>();
        foreach (var observed in actual.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var mode = observed.GetProperty("Mode").GetString();
            var expected = oracle.RootElement.GetProperty("Cases").EnumerateArray().Single(c => c.GetProperty("Mode").GetString() == mode);
            var path = model.RootElement.GetProperty("Paths").EnumerateArray().Single(c => c.GetProperty("Mode").GetString() == mode);
            var differences = EditMatrixVerification.Compare(expected.GetProperty("Reopened"), observed.GetProperty("Reopened"));
            var bound = string.Equals(path.GetProperty("OutputSha256").GetString(), observed.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase);
            var passed = bound && path.GetProperty("Passed").GetBoolean() && observed.GetProperty("Stable").GetBoolean() &&
                observed.GetProperty("SourceUnchanged").GetBoolean() && differences.All(d => d.Classification != "unexpected");
            if (!passed) failures++;
            reports.Add(new { Mode = mode, Passed = passed, Bound = bound, Differences = differences });
        }
        ProbeRunner.WriteNew(verdictPath, JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures, Paths = reports,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), ModelSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)),
            VerificationSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(actualPath)) }, ProbeRunner.JsonOptions));
        Console.WriteLine($"11 applied paths, 3 guarded paths: {failures} failures."); return failures == 0 ? 0 : 2;
    }
}
