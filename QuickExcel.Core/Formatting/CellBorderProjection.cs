using QuickExcel.Domain;

namespace QuickExcel.Core;

public static class CellBorderProjection
{
    /// <summary>Merged cells keep perimeter formatting on their physical edge cells in XLSX.</summary>
    public static StyleDefinition ForRange(CellRectangle range, Func<int, int, StyleDefinition> styleAt)
    {
        var style = styleAt(range.Top, range.Left);
        if (range.Count == 1) return style;
        BorderStroke Strongest(IEnumerable<BorderStroke> lines) => lines.MaxBy(b => CellBorders.Weight(b.Style))!;
        var left = Strongest(Enumerable.Range(range.Top, range.Rows).Select(r => { var s = styleAt(r, range.Left); return new BorderStroke(s.LeftBorder, s.LeftBorderColor); }));
        var right = Strongest(Enumerable.Range(range.Top, range.Rows).Select(r => { var s = styleAt(r, range.Right); return new BorderStroke(s.RightBorder, s.RightBorderColor); }));
        var top = Strongest(Enumerable.Range(range.Left, range.Columns).Select(c => { var s = styleAt(range.Top, c); return new BorderStroke(s.TopBorder, s.TopBorderColor); }));
        var bottom = Strongest(Enumerable.Range(range.Left, range.Columns).Select(c => { var s = styleAt(range.Bottom, c); return new BorderStroke(s.BottomBorder, s.BottomBorderColor); }));
        return style with { LeftBorder = left.Style, LeftBorderColor = left.Color, RightBorder = right.Style, RightBorderColor = right.Color,
            TopBorder = top.Style, TopBorderColor = top.Color, BottomBorder = bottom.Style, BottomBorderColor = bottom.Color };
    }
}
