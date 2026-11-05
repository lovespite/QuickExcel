using QuickExcel.Domain;

namespace QuickExcel.Core;

public static class CellBorderProjection
{
    /// <summary>Merged cells keep perimeter formatting on their physical edge cells in XLSX.</summary>
    public static StyleDefinition ForRange(CellRectangle range, Func<int, int, StyleDefinition> styleAt, WorksheetSnapshot? sheet = null)
    {
        var style = styleAt(range.Top, range.Left);
        if (range.Count == 1) return style;
        // Imported merges may span an entire sheet. Inspect sparse formatting locations in that case.
        var rows = sheet is not null && (long)range.Rows + range.Columns > 4096
            ? sheet.Cells.Keys.Where(p => p.Column == range.Left || p.Column == range.Right).Select(p => p.Row).Concat(sheet.Rows.Keys).Append(range.Top).Append(range.Bottom).Where(r => r >= range.Top && r <= range.Bottom).Distinct()
            : Enumerable.Range(range.Top, range.Rows);
        var columns = sheet is not null && (long)range.Rows + range.Columns > 4096
            ? sheet.Cells.Keys.Where(p => p.Row == range.Top || p.Row == range.Bottom).Select(p => p.Column).Concat(sheet.Columns.Keys).Append(range.Left).Append(range.Right).Where(c => c >= range.Left && c <= range.Right).Distinct()
            : Enumerable.Range(range.Left, range.Columns);
        BorderStroke Strongest(IEnumerable<BorderStroke> lines) => lines.MaxBy(b => CellBorders.Weight(b.Style))!;
        var left = Strongest(rows.Select(r => { var s = styleAt(r, range.Left); return new BorderStroke(s.LeftBorder, s.LeftBorderColor); }));
        var right = Strongest(rows.Select(r => { var s = styleAt(r, range.Right); return new BorderStroke(s.RightBorder, s.RightBorderColor); }));
        var top = Strongest(columns.Select(c => { var s = styleAt(range.Top, c); return new BorderStroke(s.TopBorder, s.TopBorderColor); }));
        var bottom = Strongest(columns.Select(c => { var s = styleAt(range.Bottom, c); return new BorderStroke(s.BottomBorder, s.BottomBorderColor); }));
        return style with { LeftBorder = left.Style, LeftBorderColor = left.Color, RightBorder = right.Style, RightBorderColor = right.Color,
            TopBorder = top.Style, TopBorderColor = top.Color, BottomBorder = bottom.Style, BottomBorderColor = bottom.Color };
    }
}
