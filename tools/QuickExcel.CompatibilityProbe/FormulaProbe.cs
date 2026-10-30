using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class FormulaProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("This probe must not execute scripts."); }
    public static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory);
        var original = File.ReadAllBytes(source); var storage = new TypedWorkbookStorage(); var book = storage.Open(source);
        using var session = new DocumentSession(book); var id = book.Sheets[0].Id; var readOnly = book.FileState!.FormulaReadOnly;
        string? rejectedEdit = null;
        if (readOnly)
        {
            try { session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(16)))])); }
            catch (InvalidOperationException e) { rejectedEdit = e.Message; }
            session.UpdateNavigation(id, new(4, 4));
        }
        else session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(16)))]));
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        if (!session.TryCommitCalculation(result)) throw new InvalidOperationException("Calculation commit failed.");
        var output = Path.Combine(directory, readOnly ? "preserved.xlsx" : "edited.diagnostic-copy.xlsx");
        storage.Save(output, session.Snapshot, discardUnsupported: !readOnly);
        var reopened = storage.Open(output); var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(output));
        var sourceAudit = OpenXmlAudit.Inspect(original);
        var introducedIssues = audit.Issues.Except(sourceAudit.Issues).ToArray();
        var identical = original.SequenceEqual(File.ReadAllBytes(output));
        var formulaValues = book.Sheets.SelectMany(s => s.Cells.Where(p => p.Value.Content.Kind == CellContentKind.Formula).Select(p =>
        {
            var address = new CellAddress(s.Id, p.Key.Row, p.Key.Column); var calculated = result.Cells[address];
            var savedSheet = reopened.Sheets.Single(r => r.Name == s.Name); var saved = savedSheet.GetCell(p.Key.Row, p.Key.Column)!;
            return new { Sheet = s.Name, Cell = p.Key, Formula = p.Value.Content.FormulaText, SourceCache = CellInput.DisplayText(p.Value.Value),
                Result = calculated.Value is { } value ? CellInput.DisplayText(value) : null, State = calculated.State.ToString(), SavedFormula = saved.Content.FormulaText, SavedCache = CellInput.DisplayText(saved.Value) };
        })).ToArray();
        var sourceUnchanged = original.SequenceEqual(File.ReadAllBytes(source));
        var valid = sourceUnchanged && (readOnly ? identical && introducedIssues.Length == 0 : audit.Issues.Count == 0) && formulaValues.All(f => f.Formula == f.SavedFormula) &&
            (readOnly ? rejectedEdit is not null && identical && formulaValues.All(f => f.State == "Blocked" && f.SourceCache == f.Result && f.Result == f.SavedCache)
                : formulaValues.Any(f => f.State == "Blocked" && f.SourceCache == f.Result) && formulaValues.Any(f => f.State == "Fresh" && f.Result == "18"));
        var report = new { Valid = valid, ReadOnly = readOnly, RejectedEdit = rejectedEdit, PreservedPackage = identical, SourceUnchanged = sourceUnchanged,
            book.FileState.PreservedFormulas, Formulas = formulaValues, Diagnostics = TypedWorkbookStorage.InspectCompatibility(session.Snapshot),
            SourceSchemaIssues = sourceAudit.Issues, SchemaIssues = audit.Issues, IntroducedSchemaIssues = introducedIssues };
        File.WriteAllText(Path.Combine(directory, "formulas-report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"ReadOnly={readOnly}; SourceUnchanged={sourceUnchanged}; PreservedPackage={identical}; SchemaIssues={audit.Issues.Count}; Valid={valid}");
        return valid ? 0 : 1;
    }
}
