using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class ConditionalFunctionsProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Conditional probe cannot execute scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        if (!oracle.GetProperty("Stable").GetBoolean() || !oracle.GetProperty("EditedStable").GetBoolean() || oracle.GetProperty("Cases").GetArrayLength() != 1706)
            throw new InvalidDataException("Incomplete Excel conditional oracle.");
        var source = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, "initial.xlsx"); var original = File.ReadAllBytes(source);
        if (!OpenXmlAudit.Hash(original).Equals(oracle.GetProperty("InitialSha256").GetString(), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Oracle source hash mismatch.");
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(source)); var calculator = new WorkbookCalculator(new NoScripts());
        if (session.Snapshot.FileState!.FormulaReadOnly) throw new InvalidDataException("Supported conditional matrix imported as read-only.");
        var checks = new List<object>(); var failures = 0; var exports = new List<object>();
        void Compare(WorkbookSnapshot snapshot, string state, string stage)
        {
            var results = snapshot.Sheets.Single(s => s.Name == "Results");
            foreach (var expected in oracle.GetProperty(state).EnumerateArray())
            {
                var address = (ReferenceNode)FormulaSyntax.Parse("=" + expected.GetProperty("Address").GetString()); var cell = results.GetCell(address.Row, address.Column)!;
                var same = Matches(cell.Value, expected) && cell.Content.FormulaText == expected.GetProperty("Formula").GetString();
                if (!same) failures++;
                checks.Add(new { Stage = stage, Id = expected.GetProperty("Id").GetString(), Passed = same, Expected = expected.GetProperty("Value").Clone(),
                    Kind = cell.Value.Kind.ToString(), Value = CellInput.DisplayText(cell.Value), Formula = cell.Content.FormulaText });
            }
        }
        void Calculate(string state, string stage)
        {
            var calculated = calculator.Calculate(session.BeginCalculation());
            if (calculated.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)) || !session.TryCommitCalculation(calculated))
                throw new InvalidOperationException("Conditional results must be fully recalculated.");
            Compare(session.Snapshot, state, stage);
        }
        void Save(string state)
        {
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var path = Path.Combine(directory, state.ToLowerInvariant() + "." + extension); string? rejection = null;
                try { storage.Save(path, session.Snapshot); }
                catch (WorkbookCompatibilityException ex) { rejection = ex.Message; storage.Save(path, session.Snapshot, discardUnsupported: true); }
                using var reopened = new DocumentSession(storage.Open(path));
                if (!reopened.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(reopened.BeginCalculation()))) throw new InvalidOperationException("Reopen calculation commit failed.");
                Compare(reopened.Snapshot, state, state + "/" + extension);
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()).ToArray() : [];
                if (issues.Length > 0) failures++;
                exports.Add(new { State = state, Extension = extension, Sha256 = OpenXmlAudit.Hash(File.ReadAllBytes(path)), RejectedNormalSave = rejection, SchemaIssues = issues });
            }
        }
        Calculate("Initial", "initial-calculation"); Save("Initial");
        var data = session.Snapshot.Sheets.Single(s => s.Name == "Data");
        session.Execute(new SetCellsCommand([new(new(data.Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(7))), new(new(data.Id, 2, 1), CellContent.FromLiteral(CellValue.FromNumber(300)))]));
        Calculate("Edited", "edited-calculation");
        if (!session.Undo() || session.CanUndo) throw new InvalidOperationException("Edit was not one undo transaction."); Calculate("Initial", "undo-calculation");
        if (!session.Redo()) throw new InvalidOperationException("Redo failed."); Calculate("Edited", "redo-calculation"); Save("Edited");
        var unchanged = original.SequenceEqual(File.ReadAllBytes(source)); if (!unchanged) failures++;
        ProbeRunner.WriteNew(Path.Combine(directory, "conditional-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), SourceUnchanged = unchanged, Exports = exports, Checks = checks }, ProbeRunner.JsonOptions));
        Console.WriteLine($"1706 conditional formulas; checks={checks.Count}; failures={failures}; source unchanged={unchanged}"); return failures == 0 ? 0 : 2;
    }
    private static bool Matches(CellValue value, JsonElement expected)
    {
        if (value.Kind.ToString() != expected.GetProperty("Kind").GetString()) return false;
        var raw = expected.GetProperty("Value");
        return value.Kind switch
        {
            CellValueKind.Number => Math.Abs(value.NumberValue - raw.GetDouble()) <= 1e-12 * Math.Max(1, Math.Abs(raw.GetDouble())),
            CellValueKind.Boolean => value.BooleanValue == raw.GetBoolean(), CellValueKind.Blank => true,
            _ => CellInput.DisplayText(value) == raw.GetString()
        };
    }
    public static int Verify(string oraclePath, string initialPath, string editedPath, string modelPath, string verdictPath)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); using var model = JsonDocument.Parse(File.ReadAllBytes(modelPath));
        var failures = model.RootElement.GetProperty("Failures").GetInt32(); var paths = new List<object>();
        if (model.RootElement.GetProperty("OracleSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath))) throw new InvalidDataException("Oracle/model hash mismatch.");
        foreach (var (state, path) in new[] { ("Initial", initialPath), ("Edited", editedPath) })
        {
            using var actual = JsonDocument.Parse(File.ReadAllBytes(path)); var observed = actual.RootElement;
            var output = model.RootElement.GetProperty("Exports").EnumerateArray().Single(e => e.GetProperty("State").GetString() == state && e.GetProperty("Extension").GetString() == "xlsx");
            var hashMatches = string.Equals(observed.GetProperty("SourceSha256").GetString(), output.GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase);
            var differences = EditMatrixVerification.Compare(oracle.RootElement.GetProperty(state), observed.GetProperty("Initial"));
            var passed = hashMatches && observed.GetProperty("Stable").GetBoolean() && observed.GetProperty("SourceUnchanged").GetBoolean() &&
                observed.GetProperty("ReferenceMatches").GetBoolean() && differences.Count == 0;
            if (!passed) failures++; paths.Add(new { State = state, Passed = passed, HashMatches = hashMatches, Differences = differences });
        }
        ProbeRunner.WriteNew(verdictPath, JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures, Paths = paths,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), ModelSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)) }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Excel conditional verification: failures={failures}"); return failures == 0 ? 0 : 2;
    }
}
