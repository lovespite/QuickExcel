using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class ScalarFunctionsProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Scalar formula probes cannot execute scripts."); }
    private static CellPosition Position(string text) { var r = (ReferenceNode)FormulaSyntax.Parse("=" + text); return new(r.Row, r.Column); }
    private static CellValue Value(JsonElement item) => item.GetProperty("Kind").GetString() switch
    {
        "Number" => CellValue.FromNumber(item.GetProperty("Value").GetDouble()), "Text" => CellValue.FromText(item.GetProperty("Value").GetString()!),
        "Boolean" => CellValue.FromBoolean(item.GetProperty("Value").GetBoolean()), "Error" => CellInput.FromErrorText(item.GetProperty("Value").GetString()!), _ => CellValue.Blank
    };
    private static bool Equal(CellValue a, CellValue b) => a.Kind == b.Kind && (a.Kind == CellValueKind.Number
        ? Math.Abs(a.NumberValue - b.NumberValue) <= 1e-12 * Math.Max(1, Math.Abs(b.NumberValue)) : a == b);
    private static CellSnapshot Cell(WorkbookSnapshot snapshot, JsonElement item)
    { var p = Position(item.GetProperty("Address").GetString()!); return snapshot.Sheets.Single(s => s.Name == item.GetProperty("Sheet").GetString()).GetCell(p.Row, p.Column)!; }

    public static int Run(string specPath, string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); using var specJson = JsonDocument.Parse(File.ReadAllBytes(specPath)); using var oracleJson = JsonDocument.Parse(File.ReadAllBytes(oraclePath));
        var spec = specJson.RootElement; var oracle = oracleJson.RootElement;
        if (!oracle.GetProperty("Stable").GetBoolean() || !oracle.GetProperty("EditedStable").GetBoolean() ||
            !string.Equals(oracle.GetProperty("SpecSha256").GetString(), OpenXmlAudit.Hash(File.ReadAllBytes(specPath)), StringComparison.OrdinalIgnoreCase) ||
            !spec.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(oracle.GetProperty("Initial").EnumerateArray().Select(c => c.GetProperty("Id").GetString())))
            throw new InvalidDataException("Incomplete or unbound scalar oracle.");
        var sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(oraclePath))!; var source = Path.Combine(sourceDirectory, "initial.xlsx"); var bytes = File.ReadAllBytes(source);
        if (!OpenXmlAudit.Hash(bytes).Equals(oracle.GetProperty("InitialSha256").GetString(), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Source hash mismatch.");
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(source)); var calculator = new WorkbookCalculator(new NoScripts());
        if (session.Snapshot.FileState!.FormulaReadOnly) throw new InvalidDataException("Supported matrix imported as read-only.");
        var failures = 0; var checks = new List<object>(); var exports = new List<object>(); var guarded = new List<object>();
        void Compare(WorkbookSnapshot snapshot, string state, string stage)
        {
            foreach (var expected in oracle.GetProperty(state).EnumerateArray())
            {
                var cell = Cell(snapshot, expected); var same = Equal(cell.Value, Value(expected)) && cell.Calculation?.State is CalculationState.Fresh or CalculationState.Error &&
                    BusinessBaselineProbe.CanonicalFormula(cell.Content.FormulaText) == BusinessBaselineProbe.CanonicalFormula(expected.GetProperty("Formula").GetString());
                if (!same) failures++; checks.Add(new { Stage = stage, Id = expected.GetProperty("Id").GetString(), Passed = same,
                    Expected = expected.GetProperty("Value").Clone(), Kind = cell.Value.Kind.ToString(), Value = CellInput.DisplayText(cell.Value), Formula = cell.Content.FormulaText });
            }
        }
        void Calculate(string state, string stage)
        {
            var result = calculator.Calculate(session.BeginCalculation());
            if (result.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)) || !session.TryCommitCalculation(result)) throw new InvalidOperationException("Incomplete scalar calculation.");
            Compare(session.Snapshot, state, stage);
        }
        void Save(string state)
        {
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var path = Path.Combine(directory, state.ToLowerInvariant() + "." + extension); string? rejection = null;
                try { storage.Save(path, session.Snapshot); } catch (WorkbookCompatibilityException ex) { rejection = ex.Message; storage.Save(path, session.Snapshot, discardUnsupported: true); }
                using var reopened = new DocumentSession(storage.Open(path));
                if (!reopened.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(reopened.BeginCalculation()))) throw new InvalidOperationException("Reopen calculation failed.");
                Compare(reopened.Snapshot, state, state + "/" + extension);
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()).ToArray() : [];
                if (issues.Length > 0) failures++;
                exports.Add(new { State = state, Extension = extension, Sha256 = OpenXmlAudit.Hash(File.ReadAllBytes(path)), RejectedNormalSave = rejection, SchemaIssues = issues });
            }
        }
        Calculate("Initial", "initial-calculation"); Save("Initial");
        session.Execute(new SetCellsCommand(spec.GetProperty("Edits").EnumerateArray().Select(e =>
        {
            var sheet = session.Snapshot.Sheets.Single(s => s.Name == e.GetProperty("Sheet").GetString()); var p = Position(e.GetProperty("Address").GetString()!);
            var content = e.GetProperty("Kind").GetString() == "Formula" ? CellContent.FromFormula(e.GetProperty("Value").GetString()!) : CellContent.FromLiteral(Value(e));
            return new CellEdit(new(sheet.Id, p.Row, p.Column), content);
        })));
        Calculate("Edited", "edited-calculation");
        if (!session.Undo() || session.CanUndo) throw new InvalidOperationException("Edit is not one transaction."); Calculate("Initial", "undo-calculation");
        if (!session.Redo()) throw new InvalidOperationException("Redo failed."); Calculate("Edited", "redo-calculation"); Save("Edited");
        foreach (var guard in oracle.GetProperty("Guarded").EnumerateArray())
        {
            var path = Path.Combine(sourceDirectory, guard.GetProperty("File").GetString()!); var original = File.ReadAllBytes(path);
            if (!OpenXmlAudit.Hash(original).Equals(guard.GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Guard source mismatch.");
            using var retained = new DocumentSession(storage.Open(path)); var before = retained.Snapshot; string? rejection = null;
            try { retained.Execute(new SetCellsCommand([new(new(before.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(991)))])); }
            catch (InvalidOperationException ex) { rejection = ex.Message; }
            var unchanged = ReferenceEquals(before, retained.Snapshot) && !retained.CanUndo;
            if (!retained.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(retained.BeginCalculation()))) throw new InvalidOperationException("Guard calculation failed.");
            var observed = guard.GetProperty("Observed").EnumerateArray().Single(); var cell = Cell(retained.Snapshot, observed);
            var output = Path.Combine(directory, "guard-" + guard.GetProperty("Id").GetString() + ".xlsx"); storage.Save(output, retained.Snapshot);
            var passed = before.FileState!.FormulaReadOnly && rejection is not null && unchanged && cell.Calculation?.State == CalculationState.Blocked &&
                Equal(cell.Value, Value(observed)) && original.SequenceEqual(File.ReadAllBytes(output)) && original.SequenceEqual(File.ReadAllBytes(path));
            if (!passed) failures++; guarded.Add(new { Id = guard.GetProperty("Id").GetString(), Passed = passed, Rejection = rejection, PreservedPackage = original.SequenceEqual(File.ReadAllBytes(output)) });
        }
        var sourceUnchanged = bytes.SequenceEqual(File.ReadAllBytes(source)); if (!sourceUnchanged) failures++;
        ProbeRunner.WriteNew(Path.Combine(directory, "scalar-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures,
            SpecSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(specPath)), OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), SourceUnchanged = sourceUnchanged,
            Exports = exports, Guarded = guarded, Checks = checks }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Scalar matrix: checks={checks.Count}; guards={guarded.Count}; failures={failures}"); return failures == 0 ? 0 : 2;
    }
    public static int Verify(string oraclePath, string initialPath, string editedPath, string modelPath, string verdictPath)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); using var model = JsonDocument.Parse(File.ReadAllBytes(modelPath)); var failures = model.RootElement.GetProperty("Failures").GetInt32(); var paths = new List<object>();
        if (model.RootElement.GetProperty("OracleSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath))) throw new InvalidDataException("Oracle/model hash mismatch.");
        foreach (var (state, path) in new[] { ("Initial", initialPath), ("Edited", editedPath) })
        {
            using var actual = JsonDocument.Parse(File.ReadAllBytes(path)); var observed = actual.RootElement;
            var output = model.RootElement.GetProperty("Exports").EnumerateArray().Single(e => e.GetProperty("State").GetString() == state && e.GetProperty("Extension").GetString() == "xlsx");
            var bound = string.Equals(observed.GetProperty("SourceSha256").GetString(), output.GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase);
            var differences = EditMatrixVerification.Compare(oracle.RootElement.GetProperty(state), observed.GetProperty("Initial"));
            var passed = bound && observed.GetProperty("Stable").GetBoolean() && observed.GetProperty("SourceUnchanged").GetBoolean() &&
                observed.GetProperty("ReferenceMatches").GetBoolean() && differences.Count == 0;
            if (!passed) failures++; paths.Add(new { State = state, Passed = passed, HashMatches = bound, Differences = differences });
        }
        ProbeRunner.WriteNew(verdictPath, JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures, Paths = paths,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), ModelSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)) }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Excel scalar verification: failures={failures}"); return failures == 0 ? 0 : 2;
    }
}
