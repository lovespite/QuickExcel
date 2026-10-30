namespace QuickExcel.Formulas;

/// <summary>Reject unimplemented array results before import or portable export can reinterpret them as scalars.</summary>
public static class ScalarFormulaModes
{
    public static void Validate(FormulaNode expression)
    {
        if (DateMathFunctions.ArrayResult(expression)) throw new FormatException("多格公式结果尚未实现；保留来源公式和缓存。");
        foreach (var node in FormulaSyntax.Walk(expression))
        {
            if (node is UnaryNode or BinaryNode && DateMathFunctions.ArrayResult(node)) throw new FormatException("数组运算尚未实现。");
            if (node is not FunctionNode function) continue;
            var args = function.Arguments;
            if (function.Name == "INDEX" && args.Count == 4) throw new FormatException("INDEX 多区域引用形式尚未实现；保留来源公式和缓存。");
            if (function.Name is "ABS" or "NOT" or "ROUND" or "LEFT" or "RIGHT" or "MID" or "LEN" or "TRIM" or "UPPER" or "LOWER" or "IF" or "IFERROR" or "CONCATENATE")
            {
                if (args.Any(DateMathFunctions.ArrayResult)) throw new FormatException("此函数的数组参数/返回尚未实现。");
            }
            else if (function.Name is "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT" or "COUNTA" or "PRODUCT" or "AND" or "OR")
            {
                if (args.Any(a => a is not ReferenceNode && DateMathFunctions.ArrayResult(a))) throw new FormatException("聚合函数暂不支持计算得到的多格数组。");
            }
            else if (function.Name == "INDEX" && args.Count is 2 or 3)
            {
                if (args.Skip(1).Any(DateMathFunctions.ArrayResult) || args[0] is not ReferenceNode && DateMathFunctions.ArrayResult(args[0]))
                    throw new FormatException("INDEX 目前只支持标量索引和直接/名称区域。");
                var shape = args[0] is ReferenceNode reference ? ConditionalRanges.Shape(reference) : (Rows: 1, Columns: 1);
                if (shape.Rows > 1 && shape.Columns > 1 && args.Count == 2) throw new FormatException("二维 INDEX 省略列号会返回数组，尚未实现。");
                void RequireScalarIndex(FormulaNode index, int size)
                {
                    if (size <= 1) return;
                    if (LookupModes.TryConstant(index, out var number))
                    { if (Math.Floor(number) == 0) throw new FormatException("INDEX 零索引的多格结果尚未实现。"); return; }
                    if (index is FunctionNode { Name: "MATCH" or "XMATCH" }) return; // Successful matches are strictly positive.
                    if (index is ValueNode { Value.Kind: QuickExcel.Domain.CellValueKind.Error }) return;
                    throw new FormatException("INDEX 多格维度的索引目前须为常量或 MATCH/XMATCH 结果，以避免未实现的零索引数组。");
                }
                RequireScalarIndex(args[1], shape.Rows == 1 && args.Count == 2 ? shape.Columns : shape.Rows);
                if (args.Count == 3) RequireScalarIndex(args[2], shape.Columns);
            }
        }
    }
}
