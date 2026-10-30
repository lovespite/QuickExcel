using System.Globalization;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class SheetNamesProbe
{
    private static readonly string[] Ids = ["copy", "reorder", "rename-global", "rename-local", "edit-global", "delete-global", "delete-local", "create-shadow"];
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Sheet/name probe cannot execute scripts."); }
    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Evidence directory must be new.");
        Directory.CreateDirectory(directory); using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        if (oracle.GetProperty("Origin").GetString() != "independent-excel-oracle" || !oracle.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids))
            throw new InvalidDataException("Incomplete independent sheet/name oracle.");
        var paths = new List<object>(); var failures = 0;
        foreach (var expected in oracle.GetProperty("Cases").EnumerateArray())
        {
            var id = expected.GetProperty("Id").GetString()!; var sourcePath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, id + "-source.xlsx");
            var source = File.ReadAllBytes(sourcePath); var hash = OpenXmlAudit.Hash(source); var guarded = id == "delete-local";
            if (!hash.Equals(expected.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase) || !expected.GetProperty("SourceUnchanged").GetBoolean() ||
                !guarded && !expected.GetProperty("Stable").GetBoolean()) throw new InvalidDataException("Oracle/source mismatch.");
            var storage = new TypedWorkbookStorage(); var original = storage.Open(sourcePath); using var session = new DocumentSession(original); var sheet = original.Sheets[0];
            DocumentCommand command = id switch
            {
                "copy" => new CopySheetCommand(sheet.Id, "Copy", 2), "reorder" => new ReorderSheetCommand(sheet.Id, 2),
                "rename-global" => new RenameDefinedNameCommand("GlobalSource", null, "Amount"),
                "rename-local" => new RenameDefinedNameCommand("Rate", sheet.Id, "LocalRate"),
                "edit-global" => new DefineNameCommand(new("GlobalSource", "=Other!$A$1")),
                "delete-global" => new DeleteDefinedNameCommand("GlobalOther", null),
                "delete-local" => new DeleteDefinedNameCommand("Rate", sheet.Id),
                _ => new DefineNameCommand(new("GlobalOther", "=Source!$A$3", sheet.Id))
            };
            if (guarded)
            {
                string? rejection = null;
                try { session.Execute(command); } catch (NotSupportedException ex) { rejection = ex.Message; }
                var ok = rejection is not null && ReferenceEquals(original, session.Snapshot) && !session.CanUndo;
                if (!ok) failures++;
                paths.Add(new { Id = id, Passed = ok, Guarded = true, Rejection = rejection, SourceSha256 = hash });
                Console.WriteLine(id + ": " + (ok ? "GUARDED" : "FAIL")); continue;
            }
            session.Execute(command); var edited = session.Snapshot;
            var undo = session.Undo() && SameInput(original, session.Snapshot) && !session.CanUndo && session.Redo() && SameInput(edited, session.Snapshot);
            var calculated = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
            if (calculated.Cells.Values.Any(c => c.State is not (CalculationState.Fresh or CalculationState.Error)) || !session.TryCommitCalculation(calculated)) throw new InvalidOperationException("Sheet/name comparison requires recalculation.");
            var checks = Compare(session.Snapshot, expected.GetProperty("Reopened"), "calculated").ToList(); var rejections = new List<string>(); string? outputHash = null; var issues = new List<string>();
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
            paths.Add(new { Id = id, Passed = passed, Guarded = false, UndoRedo = undo, SourceUnchanged = unchanged, OutputSha256 = outputHash, SaveRejections = rejections, SchemaIssues = issues, Checks = checks });
            Console.WriteLine($"{id}: {(passed ? "PASS" : "FAIL")}; checks={checks.Count}; failures={checks.Count(c => !c.Passed)}; schema={issues.Count}");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "sheet-names-report.json"), JsonSerializer.SerializeToUtf8Bytes(new { OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), Failures = failures, Paths = paths }, ProbeRunner.JsonOptions));
        return failures == 0 ? 0 : 2;
    }
    private static bool SameInput(WorkbookSnapshot a, WorkbookSnapshot b) => a.ActiveSheetId == b.ActiveSheetId && a.DefinedNames.SequenceEqual(b.DefinedNames) &&
        a.Sheets.Select(s => s.Id).SequenceEqual(b.Sheets.Select(s => s.Id)) && a.Sheets.Zip(b.Sheets).All(p => p.First.Name == p.Second.Name && p.First.Cells.Count == p.Second.Cells.Count &&
            p.First.Cells.All(c => p.Second.Cells.TryGetValue(c.Key, out var cell) && c.Value.Content == cell.Content));
    private static IEnumerable<BusinessBaselineProbe.Check> Compare(WorkbookSnapshot actual, JsonElement expected, string stage)
    {
        var sheets = expected.GetProperty("Sheets").EnumerateArray().ToArray();
        yield return new(stage, "SheetCount", sheets.Length.ToString(), actual.Sheets.Length.ToString(), sheets.Length == actual.Sheets.Length);
        for (var index = 0; index < Math.Min(sheets.Length, actual.Sheets.Length); index++)
        {
            var wanted = sheets[index]; var sheet = actual.Sheets[index]; var prefix = stage + "/" + sheet.Name;
            yield return new(prefix, "Name", wanted.GetProperty("Name").GetString(), sheet.Name, wanted.GetProperty("Name").GetString() == sheet.Name);
            yield return new(prefix, "Images", wanted.GetProperty("Shapes").GetArrayLength().ToString(), sheet.Images.Length.ToString(), wanted.GetProperty("Shapes").GetArrayLength() == sheet.Images.Length);
            foreach (var item in wanted.GetProperty("Cells").EnumerateArray())
            {
                var row = item.GetProperty("Row").GetInt32(); var column = item.GetProperty("Column").GetInt32(); var location = $"{prefix}/{row}/{column}";
                var cell = sheet.GetCell(row, column); var value = cell?.Value ?? CellValue.Blank; var kind = item.GetProperty("Kind").GetString();
                yield return new(location, "Kind", kind, value.Kind.ToString(), kind == value.Kind.ToString());
                var raw = item.GetProperty("Value"); var ev = raw.ValueKind switch { JsonValueKind.Null => "", JsonValueKind.Number => raw.GetDouble().ToString("R", CultureInfo.InvariantCulture), JsonValueKind.True => "TRUE", JsonValueKind.False => "FALSE", _ => raw.GetString() };
                var av = CellInput.DisplayText(value); yield return new(location, "Value", ev, av, ev == av);
                var ef = BusinessBaselineProbe.CanonicalFormula(item.GetProperty("Formula").GetString()); var af = cell?.Content.Kind == CellContentKind.Formula ? BusinessBaselineProbe.CanonicalFormula(cell.Content.FormulaText) : null;
                yield return new(location, "Formula", ef, af, ef == af);
            }
        }
        var names = actual.DefinedNames.ToDictionary(n => n.Scope is null ? n.Name : actual.GetSheet(n.Scope.Value).Name + "!" + n.Name, StringComparer.OrdinalIgnoreCase);
        var expectedNames = expected.GetProperty("Names").EnumerateArray().ToArray();
        yield return new(stage, "NameCount", expectedNames.Length.ToString(), names.Count.ToString(), expectedNames.Length == names.Count);
        foreach (var name in expectedNames)
        {
            var key = name.GetProperty("Name").GetString()!; var ef = BusinessBaselineProbe.CanonicalFormula(name.GetProperty("Formula").GetString());
            var af = names.TryGetValue(key, out var definition) ? BusinessBaselineProbe.CanonicalFormula(definition.Formula) : null;
            yield return new(stage + "/names/" + key, "Formula", ef, af, ef == af);
        }
    }
    public static int Verify(string oraclePath, string actualPath, string modelPath, string outputPath)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); using var actual = JsonDocument.Parse(File.ReadAllBytes(actualPath)); using var model = JsonDocument.Parse(File.ReadAllBytes(modelPath));
        if (actual.RootElement.GetProperty("Origin").GetString() != "excel-output-verification" || model.RootElement.GetProperty("OracleSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)) ||
            !actual.RootElement.GetProperty("Cases").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids.Where(i => i != "delete-local")) ||
            !model.RootElement.GetProperty("Paths").EnumerateArray().Select(c => c.GetProperty("Id").GetString()).SequenceEqual(Ids)) throw new InvalidDataException("Incomplete or unbound sheet/name evidence.");
        var failures = model.RootElement.GetProperty("Failures").GetInt32(); var paths = new List<object>();
        foreach (var observed in actual.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var id = observed.GetProperty("Id").GetString(); var expected = oracle.RootElement.GetProperty("Cases").EnumerateArray().Single(c => c.GetProperty("Id").GetString() == id);
            var path = model.RootElement.GetProperty("Paths").EnumerateArray().Single(p => p.GetProperty("Id").GetString() == id);
            var differences = EditMatrixVerification.Compare(expected.GetProperty("Reopened"), observed.GetProperty("Reopened"));
            var bound = string.Equals(path.GetProperty("OutputSha256").GetString(), observed.GetProperty("SourceSha256").GetString(), StringComparison.OrdinalIgnoreCase);
            var passed = bound && path.GetProperty("Passed").GetBoolean() && observed.GetProperty("SourceUnchanged").GetBoolean() && observed.GetProperty("Stable").GetBoolean() && differences.All(d => d.Classification != "unexpected");
            if (!passed) failures++; paths.Add(new { Id = id, Passed = passed, Bound = bound, Differences = differences });
        }
        ProbeRunner.WriteNew(outputPath, JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures, Paths = paths,
            OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), VerificationSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(actualPath)), ModelSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)) }, ProbeRunner.JsonOptions));
        Console.WriteLine($"7 sheet/name paths; failures={failures}; qualified local-name deletion remains guarded"); return failures == 0 ? 0 : 2;
    }
}
