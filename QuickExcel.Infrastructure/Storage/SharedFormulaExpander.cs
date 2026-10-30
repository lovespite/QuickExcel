using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedWorkbookXml;
using static QuickExcel.Storage.TypedStyles;

namespace QuickExcel.Storage;

internal sealed class SharedFormulaExpander
{
    private sealed record Master(string Formula, QuickExcel.Domain.CellPosition Position, QuickExcel.Domain.CellPosition Start, QuickExcel.Domain.CellPosition End);
    private readonly Dictionary<int, Master> _masters = [];
    private static readonly Regex ReferencePattern = new(@"(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})(?![A-Za-z0-9_.(])", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public SharedFormulaExpander(XElement sheet)
    {
        foreach (var cell in sheet.Element(Main + "sheetData")?.Elements(Main + "row").SelectMany(r => r.Elements(Main + "c")) ?? [])
        {
            var formula = cell.Element(Main + "f");
            if ((string?)formula?.Attribute("t") != "shared" || formula.Value.Length == 0) continue;
            var position = Position(XlsxPackage.Attr(cell, "r"));
            var range = XlsxPackage.Attr(formula, "ref").Split(':');
            if (range.Length is < 1 or > 2) throw new InvalidDataException("共享公式范围无效。");
            var start = Position(range[0]); var end = Position(range[^1]);
            if (start.Row > position.Row || start.Column > position.Column || end.Row < position.Row || end.Column < position.Column ||
                !_masters.TryAdd(Int(XlsxPackage.Attr(formula, "si"), int.MaxValue), new(formula.Value, position, start, end)))
                throw new InvalidDataException("共享公式主单元格或编号无效。");
        }
    }

    public string Expand(XElement formula, QuickExcel.Domain.CellPosition position)
    {
        if ((string?)formula.Attribute("t") is null or "normal") return formula.Value;
        if ((string?)formula.Attribute("t") != "shared") throw new NotSupportedException("数组/动态数组公式尚不支持可编辑导入。");
        if (!_masters.TryGetValue(Int(XlsxPackage.Attr(formula, "si"), int.MaxValue), out var master) ||
            position.Row < master.Start.Row || position.Column < master.Start.Column || position.Row > master.End.Row || position.Column > master.End.Column)
            throw new InvalidDataException("共享公式缺少主公式或从属地址越界。");
        if (position == master.Position) return master.Formula;
        if (formula.Value.Length != 0) throw new InvalidDataException("共享公式含重复的主表达式。");
        return Translate(master.Formula, position.Row - master.Position.Row, position.Column - master.Position.Column);
    }

    public string Source(XElement formula, QuickExcel.Domain.CellPosition position)
    {
        if (!_masters.TryGetValue(Int(XlsxPackage.Attr(formula, "si"), int.MaxValue), out var master) ||
            position.Row < master.Start.Row || position.Column < master.Start.Column || position.Row > master.End.Row || position.Column > master.End.Column)
            throw new InvalidDataException("共享公式缺少主公式或从属地址越界。");
        return master.Formula;
    }

    private static string Translate(string formula, int rowOffset, int columnOffset)
    {
        var output = new StringBuilder();
        var start = 0;
        void References(string segment)
        {
            if (segment.Contains('[') || segment.Contains(']')) throw new NotSupportedException("外部/结构化共享公式引用尚不支持。");
            output.Append(ReferencePattern.Replace(segment, match =>
            {
                var column = 0;
                foreach (var ch in match.Groups[2].Value.ToUpperInvariant()) column = column * 26 + ch - 'A' + 1;
                var row = int.Parse(match.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
                if (column > 16384 || row > 1048576) return match.Value;
                if (match.Groups[1].Value.Length == 0) column += columnOffset;
                if (match.Groups[3].Value.Length == 0) row += rowOffset;
                if (row is < 1 or > 1048576 || column is < 1 or > 16384) return "#REF!";
                var letters = "";
                for (var current = column; current > 0; current = (current - 1) / 26) letters = (char)('A' + (current - 1) % 26) + letters;
                return match.Groups[1].Value + letters + match.Groups[3].Value + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }));
        }
        for (var index = 0; index < formula.Length; index++)
        {
            if (formula[index] is not ('\'' or '"')) continue;
            References(formula[start..index]);
            var quote = formula[index]; var quotedStart = index++;
            while (index < formula.Length)
            {
                if (formula[index] == quote)
                {
                    if (index + 1 < formula.Length && formula[index + 1] == quote) { index += 2; continue; }
                    break;
                }
                index++;
            }
            if (index >= formula.Length) throw new InvalidDataException("共享公式引号未闭合。");
            output.Append(formula[quotedStart..(index + 1)]);
            start = index + 1;
        }
        References(formula[start..]);
        return output.ToString();
    }
}
