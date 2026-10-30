using QuickExcel.Domain;

namespace QuickExcel.Formulas;

/// <summary>Visibility and filter inputs are dependencies even when outside the value range.</summary>
public static class SubtotalFunctions
{
    public static bool HasUnmodeledFilters(WorkbookSnapshot book) => book.FileState?.CompatibilityDiagnostics.Any(d =>
        d.Feature.Contains("autoFilter", StringComparison.OrdinalIgnoreCase) || d.Feature.Contains("table", StringComparison.OrdinalIgnoreCase)) == true;
    public static bool Contains(FormulaNode expression) => FormulaSyntax.Walk(expression).OfType<FunctionNode>().Any(f => f.Name == "SUBTOTAL");
    public static bool Present(WorkbookSnapshot book) => book.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula)
        .Select(c => c.Content.FormulaText).Concat(book.DefinedNames.Select(n => n.Formula)).Any(f =>
        { try { return Contains(FormulaSyntax.Parse(f)); } catch (FormatException) { return false; } });
    public static IEnumerable<ReferenceNode> FilterDependencies(FormulaNode expression, WorkbookSnapshot book)
    {
        foreach (var function in FormulaSyntax.Walk(expression).OfType<FunctionNode>().Where(f => f.Name == "SUBTOTAL"))
        foreach (var reference in function.Arguments.Skip(1).SelectMany(FormulaSyntax.Walk).OfType<ReferenceNode>())
        {
            var sheet = book.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
            foreach (var filter in sheet.Data.AllFilters)
            {
                var first = Math.Max(reference.Row, filter.Range.Top + 1); var last = Math.Min(reference.EndRow, filter.Range.Bottom);
                if (first > last) continue;
                foreach (var column in filter.Columns.Keys) yield return new(sheet.Name, first, column, last, column, IsRange: true);
            }
        }
    }
}
