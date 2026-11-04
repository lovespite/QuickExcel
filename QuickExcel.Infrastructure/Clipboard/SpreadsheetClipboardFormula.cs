using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Clipboard;

internal static partial class SpreadsheetClipboardFormula
{
    internal static readonly CellPosition Origin = new(WorkbookLimits.MaxRows / 2, WorkbookLimits.MaxColumns / 2);
    [GeneratedRegex(@"\GR(?<r>\[-?\d+\]|\d*)C(?<c>\[-?\d+\]|\d*)(?![\w.\[])", RegexOptions.CultureInvariant, 100)]
    private static partial Regex Reference();
    internal static string Read(string text)
    {
        if (text.Length > 8192) throw new FormatException("公式过长。");
        var result = new StringBuilder(); var quoted = false;
        for (var i = 0; i < text.Length;)
        {
            var ch = text[i];
            if (ch == '"') { quoted = !quoted; result.Append(ch); i++; continue; }
            if (!quoted && ch == 'R' && (i == 0 || !char.IsLetterOrDigit(text[i - 1]) && text[i - 1] != '_'))
            {
                var match = Reference().Match(text, i);
                if (match.Success)
                {
                    int Coordinate(string s, int origin)
                    {
                        if (s.Length == 0) return origin;
                        if (s[0] != '[') throw new NotSupportedException("Excel XML 未提供原选区位置，绝对/混合引用无法安全还原");
                        return checked(origin + int.Parse(s.AsSpan(1, s.Length - 2), CultureInfo.InvariantCulture));
                    }
                    var row = Coordinate(match.Groups["r"].Value, Origin.Row); var col = Coordinate(match.Groups["c"].Value, Origin.Column);
                    WorkbookLimits.ValidatePosition(row, col);
                    result.Append(FormulaSyntax.Print(new ReferenceNode(null, row, col, row, col))); i += match.Length; continue;
                }
            }
            if (!quoted && ch is '[' or ']' or '!' or '\'') throw new NotSupportedException("外部公式依赖工作表、外部工作簿或名称，不能安全绑定");
            result.Append(ch); i++;
        }
        var a1 = result.ToString(); ExternalClipboardFormula.Validate(a1); return a1;
    }
    internal static string Write(string formula, CellPosition origin)
    {
        var expression = ExternalClipboardFormula.Validate(formula);
        string Ref(int row, int col) => "R" + (row == origin.Row ? "" : "[" + (row - origin.Row).ToString(CultureInfo.InvariantCulture) + "]") +
            "C" + (col == origin.Column ? "" : "[" + (col - origin.Column).ToString(CultureInfo.InvariantCulture) + "]");
        string Print(FormulaNode n) => n switch
        {
            ReferenceNode r => Ref(r.Row, r.Column) + (r.IsRange ? ":" + Ref(r.EndRow, r.EndColumn) : ""),
            UnaryNode u => u.Operator == "%" ? "(" + Print(u.Operand) + ")%" : u.Operator + "(" + Print(u.Operand) + ")",
            BinaryNode b => "(" + Print(b.Left) + b.Operator + Print(b.Right) + ")",
            FunctionNode f => f.Name + "(" + string.Join(",", f.Arguments.Select(Print)) + ")",
            _ => FormulaSyntax.Print(n)
        };
        return "=" + Print(expression);
    }
}
