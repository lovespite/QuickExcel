namespace QuickExcel.Formulas;

/// <summary>Array-producing functions share one result model and spill allocator.</summary>
public static class DynamicArrayFunctions
{
    public static IReadOnlyDictionary<string, (int Min, int Max)> Signatures { get; } =
        new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["SEQUENCE"] = (1, 4), ["FILTER"] = (2, 3), ["SORT"] = (1, 4), ["UNIQUE"] = (1, 3)
        };

    public static bool IsFunction(string name) => Signatures.ContainsKey(name);
    public static bool UsesArrays(FormulaNode expression) => expression is FunctionNode { Name: "INDIRECT" }
        || expression is LetNode && DateMathFunctions.ArrayResult(expression)
        || expression is ReferenceNode { SourceExpression: FunctionNode } reference && ConditionalRanges.Shape(reference) != (1, 1)
        || FormulaSyntax.Walk(expression).Any(n =>
        n is UnaryNode { Operator: "#" } || n is LetNode let && DateMathFunctions.ArrayResult(let)
        || n is FunctionNode function && (function.Name is "SUMPRODUCT" || function.Name is "IFS" or "SWITCH" && DateMathFunctions.ArrayResult(function)
            || IsFunction(function.Name) || ReferenceFunctions.ReturnsArray(function) || ReturnsLookupArray(function) || UsesOrderStatisticArray(function)));

    private static bool UsesOrderStatisticArray(FunctionNode function) => OrderStatisticFunctions.IsFunction(function.Name)
        && StandardFunctions.AcceptsArguments(function.Name,function.Arguments.Count)
        && function.Arguments[function.Name=="RANK.EQ" ? 1 : 0] is var data && data is not ReferenceNode && DateMathFunctions.ArrayResult(data);

    private static bool ReturnsLookupArray(FunctionNode function)
    {
        if (function.Name != "XLOOKUP" || function.Arguments.Count < 3 || function.Arguments[1] is not ReferenceNode lookup || function.Arguments[2] is not ReferenceNode output) return false;
        var l = ConditionalRanges.Shape(lookup); var r = ConditionalRanges.Shape(output);
        return l.Columns == 1 && l.Rows == r.Rows && r.Columns > 1 || l.Rows == 1 && l.Columns == r.Columns && r.Rows > 1;
    }
}
