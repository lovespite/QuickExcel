using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class FilterPersistenceProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Filter probe cannot run scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        var sourcePath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, "filters.xlsx"); var source = File.ReadAllBytes(sourcePath);
        if (OpenXmlAudit.Hash(source) != oracle.GetProperty("WorkbookSha256").GetString() || !oracle.GetProperty("Stable").GetBoolean() || !oracle.GetProperty("EditedStable").GetBoolean())
            throw new InvalidDataException("Unbound or unstable Excel oracle.");
        var storage = new TypedWorkbookStorage(); var opened = storage.Open(sourcePath); using var session = new DocumentSession(opened);
        var checks = new List<object>(); var failures = 0; var paths = new List<object>(); var normalSaveRejections = new List<string>();
        void Check(WorkbookSnapshot snapshot, JsonElement records, string stage)
        {
            foreach (var expected in records.EnumerateArray())
            {
                var id = expected.GetProperty("Id").GetString()!; var sheet = snapshot.Sheets.Single(s => s.Name == id);
                var filtered = WorksheetFilterEngine.Evaluate(snapshot, sheet.Id, allowSourceCaches: true);
                var hidden = filtered.Union(sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key)).Select(r => r + 1).Order().ToArray();
                var wanted = expected.GetProperty("HiddenRows").EnumerateArray().Select(v => v.GetInt32()).ToArray();
                var expectedOrder = expected.GetProperty("Order").EnumerateArray().Select(v => v.GetDouble()).ToArray();
                var order = Enumerable.Range(1, expectedOrder.Length).Select(r => sheet.GetCell(r, 2)!.Value.NumberValue).ToArray();
                var passed = hidden.SequenceEqual(wanted) && order.SequenceEqual(expectedOrder) && sheet.Data.Filter is not null && (!id.StartsWith("sort-", StringComparison.Ordinal) || sheet.Data.Sort is not null);
                if (!passed) failures++;
                checks.Add(new { Stage = stage, Id = id, Passed = passed, Hidden = hidden, ExpectedHidden = wanted, FilterModeled = sheet.Data.Filter is not null, SortModeled = sheet.Data.Sort is not null,
                    Issues = snapshot.FileState?.CompatibilityDiagnostics.Where(d => d.SheetName == id && d.Feature.Contains("Filter", StringComparison.OrdinalIgnoreCase) || d.SheetName == id && d.Feature.Contains("sort", StringComparison.OrdinalIgnoreCase)).ToArray() });
            }
        }
        void Calculate()
        {
            var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
            if (!session.TryCommitCalculation(result) || result.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error))) throw new InvalidOperationException("Filter probe requires fresh formula caches.");
        }
        Check(opened, oracle.GetProperty("Cases"), "import");
        var first = opened.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(first, 30, 3), CellContent.FromLiteral(CellValue.FromText("outside edit")))])); Calculate();
        void SaveAndCheck(string stage, JsonElement expected)
        {
            Check(session.Snapshot, expected, stage);
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var path = Path.Combine(directory, "filters-" + stage + "." + extension);
                try { storage.Save(path, session.Snapshot); }
                catch (WorkbookCompatibilityException ex) { normalSaveRejections.Add(stage + "/" + extension + ": " + ex.Message); storage.Save(path, session.Snapshot, true); }
                var read = storage.Open(path); Check(read, expected, stage + "/" + extension);
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()).ToArray() : [];
                if (issues.Length > 0) failures++;
                paths.Add(new { Stage = stage, Extension = extension, Sha256 = OpenXmlAudit.Hash(File.ReadAllBytes(path)), Issues = issues });
            }
        }
        SaveAndCheck("initial", oracle.GetProperty("Cases"));
        var multi = session.Snapshot.Sheets.Single(s => s.Name == "multi").Id;
        session.Execute(new SetCellsCommand([new(new(multi, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(6)))])); Calculate();
        SaveAndCheck("edited", oracle.GetProperty("Edited"));
        session.Undo(); Calculate(); Check(session.Snapshot, oracle.GetProperty("Cases"), "undo"); session.Redo(); Calculate(); Check(session.Snapshot, oracle.GetProperty("Edited"), "redo");
        var unchanged = source.SequenceEqual(File.ReadAllBytes(sourcePath)); if (!unchanged) failures++;
        ProbeRunner.WriteNew(Path.Combine(directory, "filter-model.json"), JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures,
            SourceUnchanged = unchanged, OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), SourceSha256 = OpenXmlAudit.Hash(source),
            Checks = checks, Paths = paths, NormalSaveRejections = normalSaveRejections }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Filter checks={checks.Count}; failures={failures}"); return failures == 0 ? 0 : 2;
    }
}
