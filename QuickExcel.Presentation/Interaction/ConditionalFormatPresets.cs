using System.Globalization;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Presentation;

public enum ConditionalPreset
{
    GreaterThan, LessThan, Between, Equal, ContainsText, DuplicateValues, UniqueValues,
    TopItems, TopPercent, BottomItems, BottomPercent, AboveAverage, BelowAverage
}

/// <summary>Quick rules use the existing portable rule model, including live ranking formulas.</summary>
public static class ConditionalFormatPresets
{
    public static ConditionalFormatRule Create(CellRectangle range, ConditionalPreset preset, DifferentialStyle style,
        string first = "0", string second = "0", int rank = 10)
    {
        if (!Enum.IsDefined(preset)) throw new ArgumentOutOfRangeException(nameof(preset));
        var anchor = FormulaSyntax.ColumnName(range.Left) + (range.Top + 1).ToString(CultureInfo.InvariantCulture);
        if (preset is ConditionalPreset.DuplicateValues or ConditionalPreset.UniqueValues)
            return new([range], preset == ConditionalPreset.DuplicateValues ? ConditionalFormatKind.DuplicateValues : ConditionalFormatKind.UniqueValues, 1, style);
        if (preset == ConditionalPreset.ContainsText)
        {
            if (string.IsNullOrEmpty(first)) throw new ArgumentException("请输入要查找的文本。");
            return new([range], ConditionalFormatKind.ContainsText, 1, style,
                $"=NOT(ISERROR(SEARCH({Quote(first)},{anchor})))", text: first);
        }
        if (preset <= ConditionalPreset.Equal)
            return new([range], ConditionalFormatKind.CellIs, 1, style, Operand(first),
                preset == ConditionalPreset.Between ? Operand(second) : null,
                preset switch { ConditionalPreset.GreaterThan => ConditionalComparison.GreaterThan,
                    ConditionalPreset.LessThan => ConditionalComparison.LessThan, ConditionalPreset.Between => ConditionalComparison.Between, _ => ConditionalComparison.Equal });

        var absolute = FormulaSyntax.Print(new ReferenceNode(null, range.Top, range.Left, range.Bottom, range.Right,
            ColumnAbsolute: true, RowAbsolute: true, EndColumnAbsolute: true, EndRowAbsolute: true, IsRange: true));
        string condition;
        if (preset is ConditionalPreset.AboveAverage or ConditionalPreset.BelowAverage)
            condition = $"{anchor}{(preset == ConditionalPreset.AboveAverage ? ">" : "<")}AVERAGE({absolute})";
        else
        {
            var percent = preset is ConditionalPreset.TopPercent or ConditionalPreset.BottomPercent;
            if (rank < 1 || rank > (percent ? 100 : 1000)) throw new ArgumentException(percent ? "百分比须为 1 到 100 的整数。" : "项数须为 1 到 1000 的整数。");
            var bottom = preset is ConditionalPreset.BottomItems or ConditionalPreset.BottomPercent;
            var count = $"COUNT({absolute})";
            var amount = percent ? $"ROUNDUP({count}*{rank}/100,0)" : rank.ToString(CultureInfo.InvariantCulture);
            condition = $"{anchor}{(bottom ? "<=" : ">=")}{(bottom ? "SMALL" : "LARGE")}({absolute},MIN({count},{amount}))";
        }
        return new([range], ConditionalFormatKind.Expression, 1, style, $"=IFERROR(AND(ISNUMBER({anchor}),{condition}),FALSE)");
    }

    private static string Quote(string text) => "\"" + text.Replace("\"", "\"\"") + "\"";
    private static string Operand(string input)
    {
        var text = input.Trim();
        if (text.Length == 0) throw new ArgumentException("请输入比较值或公式。");
        if (text.StartsWith('=')) return text;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out var value) && double.IsFinite(value))
            return "=" + value.ToString("R", CultureInfo.InvariantCulture);
        return "=" + Quote(input);
    }
}
