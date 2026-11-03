using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public static class TextFunctions
{
    public static bool IsFunction(string name) => name is "TEXT" or "VALUE" or "SUBSTITUTE" or "REPLACE" or "FIND" or "SEARCH" or "TEXTJOIN" or "CONCAT"
        or "ISBLANK" or "ISNUMBER" or "ISTEXT" or "ISERROR" or "IFNA";
    public static void Validate(FormulaNode expression, WorkbookSnapshot snapshot, bool arrayContext = false)
    {
        foreach (var f in FormulaSyntax.Walk(expression).OfType<FunctionNode>().Where(f => IsFunction(f.Name)))
        {
            if (arrayContext && f.Name == "IFNA") continue;
            for (var i = 0; i < f.Arguments.Count; i++)
            {
                var rangeAllowed = f.Name == "CONCAT" || f.Name == "TEXTJOIN" && i >= 2;
                if (DateMathFunctions.ArrayResult(f.Arguments[i]) && (!rangeAllowed || f.Arguments[i] is not ReferenceNode))
                    throw new FormatException("文本函数的数组返回/计算区域尚未实现；连接函数支持直接或名称绑定的文本区域。");
            }
            if (f.Name != "TEXT" || f.Arguments.Count != 2) continue;
            CellValue format;
            if (f.Arguments[1] is ValueNode constant) format = constant.Value;
            else if (f.Arguments[1] is ReferenceNode reference && ConditionalRanges.Shape(reference) == (1, 1))
            {
                var cell = snapshot.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase)).GetCell(reference.Row, reference.Column);
                if (cell?.Content.Kind == CellContentKind.Formula) throw new FormatException("TEXT 格式目前须为常量或文字/数值单元格；计算得到的格式尚未实现。");
                format = cell?.Value ?? CellValue.Blank;
            }
            else throw new FormatException("TEXT 格式目前须为常量或文字/数值单元格。");
            if (format.Kind == CellValueKind.Text && CellDisplayFormatter.TextSupport(format.TextValue) == TextFormatSupport.Unsupported)
                throw new FormatException("TEXT 格式超出已支持的数字/日期格式范围；保留来源公式和缓存。");
        }
    }
}
