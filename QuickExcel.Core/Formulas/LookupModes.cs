using System.Globalization;
using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public static class LookupModes
{
    public static bool IsFunction(string name) => name is "VLOOKUP" or "HLOOKUP" or "MATCH" or "XLOOKUP" or "XMATCH";
    public static bool TryConstant(FormulaNode node, out double value)
    {
        if (node is ValueNode v)
        {
            if (v.Value.Kind == CellValueKind.Number) { value = v.Value.NumberValue; return true; }
            if (v.Value.Kind == CellValueKind.Boolean) { value = v.Value.BooleanValue ? 1 : 0; return true; }
            if (v.Value.Kind == CellValueKind.Text && double.TryParse(v.Value.TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value)) return true;
        }
        if (node is UnaryNode { Operator: "+" or "-" } u && TryConstant(u.Operand, out value)) { if (u.Operator == "-") value = -value; return true; }
        value = 0; return false;
    }
    public static void Validate(FormulaNode expression, bool arrayContext = false)
    {
        foreach (var f in FormulaSyntax.Walk(expression).OfType<FunctionNode>())
        {
            if (!IsFunction(f.Name) || f.Arguments.Count < 2) continue;
            var args = f.Arguments;
            if (args[0] is ReferenceNode needle && ConditionalRanges.Shape(needle) != (1, 1))
                throw new FormatException("查找值的数组模式尚未实现；只支持单个查找值。");
            if (args[1] is not (ReferenceNode or ValueNode)) throw new FormatException("查找范围须为直接或名称绑定的引用。");
            if (f.Name is not ("XLOOKUP" or "XMATCH")) continue;
            if (f.Name == "XLOOKUP" && args.Count >= 3)
            {
                if (args[2] is not (ReferenceNode or ValueNode)) throw new FormatException("返回范围须为直接或名称绑定的引用。");
                if (args[1] is ReferenceNode lookup && args[2] is ReferenceNode output)
                {
                    var l = ConditionalRanges.Shape(lookup); var r = ConditionalRanges.Shape(output);
                    if (!arrayContext && (l.Columns == 1 && l.Rows == r.Rows && r.Columns > 1 || l.Rows == 1 && l.Columns == r.Columns && r.Rows > 1))
                        throw new FormatException("XLOOKUP 多格返回尚未实现；导入公式保留原包和来源缓存。");
                }
                if (!arrayContext && args.Count >= 4 && args[3] is ReferenceNode fallback && ConditionalRanges.Shape(fallback) != (1, 1))
                    throw new FormatException("XLOOKUP 尚不支持多格未找到返回值。");
            }
            var matchIndex = f.Name == "XLOOKUP" ? 4 : 2;
            for (var i = matchIndex; i < Math.Min(args.Count, matchIndex + 2); i++)
            {
                if (args[i] is MissingNode) continue;
                if (!TryConstant(args[i], out var mode))
                {
                    if (ConstantSyntax(args[i])) continue; // Invalid literal types/errors are ordinary Excel argument errors.
                    throw new FormatException("现代查找模式目前须为常量；计算得到的模式尚未实现。");
                }
                mode = Math.Floor(mode);
                if (i == matchIndex && mode == 3) throw new FormatException("现代查找的正则表达式模式尚未实现。");
                if (i > matchIndex && mode is 2 or -2) throw new FormatException("现代查找目前支持从首项或末项线性搜索，二分搜索模式尚未实现。");
            }
        }
    }
    private static bool ConstantSyntax(FormulaNode node) => node is ValueNode || node is UnaryNode { Operator: "+" or "-" } u && ConstantSyntax(u.Operand);
}
