namespace QuickExcel.Formulas;

/// <summary>Reference semantics shared by dependency indexing and conditional aggregation.</summary>
public static class ConditionalRanges
{
    // Logical Excel grid size; counting absent cells does not expand the product's storage budget.
    public const int ExcelRows = 1_048_576;
    public const int ExcelColumns = 16_384;
    public static bool IsFunction(string name) => name is "SUMIF" or "SUMIFS" or "COUNTIF" or "COUNTIFS" or "AVERAGEIF" or "AVERAGEIFS" or "COUNTBLANK";
    public static (int Rows, int Columns) Shape(ReferenceNode range) =>
        (range.WholeColumns ? ExcelRows : range.EndRow - range.Row + 1, range.EndColumn - range.Column + 1);

    public static ReferenceNode Project(ReferenceNode criteria, ReferenceNode result)
    {
        var (rows, columns) = Shape(criteria);
        return result with { EndRow = Math.Min(ExcelRows - 1, result.Row + rows - 1),
            EndColumn = Math.Min(ExcelColumns - 1, result.Column + columns - 1), WholeColumns = false, IsRange = true };
    }

    public static FormulaNode ForDependencies(FormulaNode expression) => FormulaSyntax.Map(expression, node =>
        node is FunctionNode { Name: "SUMIF" or "AVERAGEIF", Arguments.Count: 3 } function &&
        function.Arguments[0] is ReferenceNode criteria && function.Arguments[2] is ReferenceNode result
            ? function with { Arguments = [criteria, function.Arguments[1], Project(criteria, result)] } : node);

    public static void ValidateModes(FormulaNode expression)
    {
        foreach (var function in FormulaSyntax.Walk(expression).OfType<FunctionNode>())
        {
            if (!IsFunction(function.Name)) continue;
            var args = function.Arguments;
            var paired = function.Name.EndsWith("IFS", StringComparison.Ordinal);
            for (var i = 0; i < args.Count; i++)
            {
                var criterion = function.Name != "COUNTBLANK" && (paired
                    ? i % 2 == (function.Name == "COUNTIFS" ? 1 : 0) && (function.Name == "COUNTIFS" || i > 0)
                    : i == 1);
                if (criterion && args[i] is ReferenceNode range && Shape(range) != (1, 1))
                    throw new FormatException("条件统计的数组/区域条件尚未实现；只支持标量条件。");
                if (!criterion && args[i] is not (ReferenceNode or ValueNode))
                    throw new FormatException("条件统计只支持直接或名称绑定的矩形区域，尚不支持计算得到的引用。");
            }
        }
    }
}
