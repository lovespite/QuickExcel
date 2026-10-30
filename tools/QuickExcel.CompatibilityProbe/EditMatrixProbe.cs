using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

/// <summary>Expected edits come from independent Excel operations, never from serializing our own command result.</summary>
public static class EditMatrixProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Edit matrix must not execute scripts."); }
    public static int Run(string inputDirectory, string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Evidence directory must be new.");
        Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath));
        var oracle = json.RootElement;
        if (oracle.GetProperty("Origin").GetString() != "real-excel-edit-oracle" || oracle.GetProperty("Failures").GetArrayLength() != 0)
            throw new InvalidDataException("Expected successful independent Excel edit evidence.");
        var reports = new List<object>(); var failures = 0;
        foreach (var fixture in oracle.GetProperty("Cases").EnumerateArray())
        {
            var id = fixture.GetProperty("Id").GetString()!;
            if (Path.GetFileName(id) != id || id.IndexOfAny(['/', '\\', ':']) >= 0) throw new InvalidDataException("Invalid fixture ID.");
            var sourcePath = Path.Combine(inputDirectory, id + ".xlsx"); var original = File.ReadAllBytes(sourcePath); var source = OpenXmlAudit.Inspect(original);
            var subdirectory = Path.Combine(directory, id); Directory.CreateDirectory(subdirectory);
            var paths = fixture.GetProperty("Paths").EnumerateArray().ToArray();
            if (!paths.Select(p => p.GetProperty("Mode").GetString()).SequenceEqual(ProbeRunner.Modes)) throw new InvalidDataException("Expected all eight ordered edit paths.");
            var originalExpected = paths[0].GetProperty("AfterReopen");
            foreach (var path in paths)
            {
                if (!string.Equals(path.GetProperty("SourceSha256").GetString(), source.Sha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Oracle/source fingerprint mismatch.");
                var mode = path.GetProperty("Mode").GetString()!; var expected = path.GetProperty("AfterReopen");
                var storage = new TypedWorkbookStorage(); var book = storage.Open(sourcePath); using var session = new DocumentSession(book);
                ProbeRunner.Apply(session, mode);
                IReadOnlyList<BusinessBaselineProbe.Check> undo = [];
                if (mode != "unchanged")
                {
                    if (!session.Undo()) throw new InvalidOperationException("Expected undoable edit.");
                    Calculate(session); undo = BusinessBaselineProbe.Compare(session.Snapshot, originalExpected, equivalentFormulas: true);
                    if (!session.Redo()) throw new InvalidOperationException("Expected redoable edit.");
                }
                Calculate(session);
                var calculated = BusinessBaselineProbe.Compare(session.Snapshot, expected, equivalentFormulas: true);
                var outputPath = Path.Combine(subdirectory, mode + ".xlsx"); string? saveRejection = null;
                // Unchanged route proves preservation. Recalculation is checked above independently, without claiming cached originals are fresh.
                var toSave = mode == "unchanged" ? book : session.Snapshot;
                try { storage.Save(outputPath, toSave); }
                catch (WorkbookCompatibilityException ex)
                {
                    saveRejection = ex.Message; outputPath = Path.Combine(subdirectory, mode + ".diagnostic-copy.xlsx");
                    storage.Save(outputPath, toSave, discardUnsupported: true);
                }
                var output = OpenXmlAudit.Inspect(File.ReadAllBytes(outputPath));
                var reopened = BusinessBaselineProbe.Compare(storage.Open(outputPath), expected, equivalentFormulas: true);
                var rawDifferences = OpenXmlAudit.Compare(source, output);
                var checks = undo.Concat(calculated).Concat(reopened).ToArray();
                var sourceUnchanged = original.SequenceEqual(File.ReadAllBytes(sourcePath));
                var passed = checks.All(c => c.Passed) && output.Issues.Count == 0 && sourceUnchanged && (mode != "unchanged" || source.Sha256 == output.Sha256);
                if (!passed) failures++;
                reports.Add(new { Id = id, Mode = mode, Operation = ProbeRunner.Describe(mode), Passed = passed,
                    Status = saveRejection is null ? "normal-save" : "normal-save-blocked;explicit-diagnostic-copy", SaveRejection = saveRejection,
                    SourceUnchanged = sourceUnchanged, Source = source, Output = output, OutputFile = Path.GetRelativePath(directory, outputPath),
                    ExpectedExcel = expected.Clone(), Undo = undo, Calculated = calculated, Reopened = reopened,
                    RawPackageDifferences = rawDifferences, Diagnostics = TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics });
                Console.WriteLine($"{id}/{mode}: {(passed ? "PASS" : "FAIL")}, {checks.Count(c => !c.Passed)} semantic failures, schema={output.Issues.Count}");
            }
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "edit-matrix-report.json"), JsonSerializer.SerializeToUtf8Bytes(new
        { Oracle = oracle.Clone(), Paths = reports, Failures = failures,
            Scope = "Model value/formula/display/font against independent Excel edits; full captured printing/format/visibility fields require Excel VerifyDirectory step. Raw differences are evidence, not an equivalence claim." }, ProbeRunner.JsonOptions));
        return failures == 0 ? 0 : 2;
    }
    private static void Calculate(DocumentSession session)
    {
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        if (result.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)))
            throw new InvalidOperationException("Edit matrix requires actual supported-formula calculation, not retained caches.");
        if (!session.TryCommitCalculation(result)) throw new InvalidOperationException("Edit matrix calculation commit failed.");
    }
}
