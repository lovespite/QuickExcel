namespace QuickExcel.Formulas;

public static class ReferenceFunctions
{
    public static bool IsPosition(string name) => name is "ROW" or "COLUMN";
    public static bool IsDimension(string name) => name is "ROWS" or "COLUMNS";
    public static bool ReturnsArray(FunctionNode function) => IsPosition(function.Name) && function.Arguments.Count == 1 &&
        function.Arguments[0] is ReferenceNode reference &&
        (function.Name == "ROW" ? ConditionalRanges.Shape(reference).Rows : ConditionalRanges.Shape(reference).Columns) > 1;

    // Coordinates and dimensions of a static reference do not read its values.
    // Keeping the original function in the bound AST preserves portable export semantics.
    public static FormulaNode ValueDependencies(FormulaNode node) => node is FunctionNode f &&
        (IsPosition(f.Name) || IsDimension(f.Name)) && (!StandardFunctions.AcceptsArguments(f.Name, f.Arguments.Count) || f.Arguments.Count == 1 && f.Arguments[0] is ReferenceNode)
        ? f with { Arguments = [] } : node;
}
