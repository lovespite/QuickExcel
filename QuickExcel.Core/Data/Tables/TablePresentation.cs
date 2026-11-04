using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Named table styling is a view projection; it never flattens banding into cell style records.</summary>
public static class TablePresentation
{
    public static StyleDefinition NumberStyle(WorksheetSnapshot sheet, IReadOnlyList<StyleDefinition> styles, int row, int column)
    {
        var direct = styles[sheet.GetStyleId(row, column)]; var normal = styles[0];
        if (direct.NumberFormatId != normal.NumberFormatId || direct.NumberFormatCode != normal.NumberFormatCode) return direct;
        var table = sheet.Data.Tables.FirstOrDefault(t => t.Range.Contains(row, column)); if (table is null) return direct;
        var definition = table.Columns[column - table.Range.Left];
        var format = table.HeaderRow && row == table.Range.Top ? definition.HeaderStyle : table.TotalsRow && row == table.Range.Bottom ? definition.TotalsStyle : definition.DataStyle;
        return format?.NumberFormatId is { } id ? direct with { NumberFormatId = id, NumberFormatCode = format.NumberFormatCode } : direct;
    }
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<WorkbookSnapshot, System.Collections.Concurrent.ConcurrentDictionary<SheetId, int[]>> HiddenRows = new();
    private static int VisibleOffset(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int start, int row)
    {
        var hidden = HiddenRows.GetOrCreateValue(snapshot).GetOrAdd(sheet.Id, _ =>
        {
            var manual = sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key);
            try { return WorksheetFilterEngine.Evaluate(snapshot, sheet.Id, WorksheetFilterEngine.CanUseSourceCaches(snapshot)).Union(manual).Order().ToArray(); }
            catch (InvalidOperationException) { return manual.Order().ToArray(); } // The view reports the unavailable filter separately.
        });
        int Before(int position) { var i = Array.BinarySearch(hidden, position); return i < 0 ? ~i : i; }
        if (Array.BinarySearch(hidden, row) >= 0) return row - start; // Excel retains the original band for invisible rows.
        return row - start - (Before(row) - Before(start));
    }
    public static StyleDefinition GetStyle(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
    {
        var style = RawStyle(snapshot, sheet, row, column, out var priority);
        (int Kind, string? Color) Edge(int nr, int nc, int side, int opposite, int ownKind, string? ownColor)
        {
            if (nr < 0 || nc < 0 || nr >= WorkbookLimits.MaxRows || nc >= WorkbookLimits.MaxColumns) return (ownKind, ownColor);
            var other = RawStyle(snapshot, sheet, nr, nc, out var rank);
            var kind = opposite switch { 0 => other.LeftBorder, 1 => other.RightBorder, 2 => other.TopBorder, _ => other.BottomBorder };
            var color = opposite switch { 0 => other.LeftBorderColor, 1 => other.RightBorderColor, 2 => other.TopBorderColor, _ => other.BottomBorderColor };
            int Weight(int k) => CellBorders.Weight(k);
            return rank[opposite] > priority[side] || rank[opposite] == priority[side] && (Weight(kind) > Weight(ownKind) || Weight(kind) == Weight(ownKind) && side is 0 or 2) ? (kind, color) : (ownKind, ownColor);
        }
        var left = Edge(row, column - 1, 0, 1, style.LeftBorder, style.LeftBorderColor); var right = Edge(row, column + 1, 1, 0, style.RightBorder, style.RightBorderColor);
        var top = Edge(row - 1, column, 2, 3, style.TopBorder, style.TopBorderColor); var bottom = Edge(row + 1, column, 3, 2, style.BottomBorder, style.BottomBorderColor);
        return style with { LeftBorder = left.Kind, LeftBorderColor = left.Color, RightBorder = right.Kind, RightBorderColor = right.Color,
            TopBorder = top.Kind, TopBorderColor = top.Color, BottomBorder = bottom.Kind, BottomBorderColor = bottom.Color };
    }
    private static StyleDefinition RawStyle(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column, out int[] priorities)
    {
        var ranks = new int[4]; priorities = ranks;
        var direct = snapshot.Styles[sheet.GetStyleId(row, column)]; var table = sheet.Data.Tables.FirstOrDefault(t => t.Range.Contains(row, column));
        if (table is null)
        { ranks[0] = direct.LeftBorder == 0 ? 0 : 100; ranks[1] = direct.RightBorder == 0 ? 0 : 100; ranks[2] = direct.TopBorder == 0 ? 0 : 100; ranks[3] = direct.BottomBorder == 0 ? 0 : 100; return direct; }
        var result = direct; var range = table.Range; var firstData = range.Top + (table.HeaderRow ? 1 : 0); var lastData = range.Bottom - (table.TotalsRow ? 1 : 0);
        void Apply(TableStyleElement? style, CellRectangle scope, int priority)
        {
            if (style is null || !scope.Contains(row, column)) return;
            var left = column == scope.Left ? style.Left : style.Vertical; var right = column == scope.Right ? style.Right : style.Vertical;
            var top = row == scope.Top ? style.Top : style.Horizontal; var bottom = row == scope.Bottom ? style.Bottom : style.Horizontal;
            if (left is not null) ranks[0] = priority; if (right is not null) ranks[1] = priority; if (top is not null) ranks[2] = priority; if (bottom is not null) ranks[3] = priority;
            result = result with
            {
                Bold = style.Bold ?? result.Bold, FontColor = style.FontColor ?? result.FontColor, FillColor = style.FillColor ?? result.FillColor,
                LeftBorder = left?.Style ?? result.LeftBorder, LeftBorderColor = left?.Color ?? result.LeftBorderColor,
                RightBorder = right?.Style ?? result.RightBorder, RightBorderColor = right?.Color ?? result.RightBorderColor,
                TopBorder = top?.Style ?? result.TopBorder, TopBorderColor = top?.Color ?? result.TopBorderColor,
                BottomBorder = bottom?.Style ?? result.BottomBorder, BottomBorderColor = bottom?.Color ?? result.BottomBorderColor,
                NumberFormatId = style.NumberFormatId ?? result.NumberFormatId, NumberFormatCode = style.NumberFormatId.HasValue ? style.NumberFormatCode : result.NumberFormatCode
            };
        }
        TableStyleElement? Part(string type) => table.Style.Elements.FirstOrDefault(e => e.Type == type);
        void Band(bool rows)
        {
            if (row < firstData || row > lastData) return;
            var first = Part(rows ? "firstRowStripe" : "firstColumnStripe"); var second = Part(rows ? "secondRowStripe" : "secondColumnStripe");
            var size1 = first?.Size ?? 1; var size2 = second?.Size ?? 1; var offset = (rows ? VisibleOffset(snapshot, sheet, firstData, row) : column - range.Left) % (size1 + size2);
            var band = offset < size1 ? first : second; var start = (rows ? row : column) - (offset < size1 ? offset : offset - size1); var size = offset < size1 ? size1 : size2;
            Apply(band, rows ? new(start, range.Left, Math.Min(lastData, start + size - 1), range.Right) : new(firstData, start, lastData, Math.Min(range.Right, start + size - 1)), rows ? 3 : 2);
        }
        Apply(Part("wholeTable"), range, 1);
        if (table.Style.ColumnStripes) Band(false); if (table.Style.RowStripes) Band(true);
        if (table.Style.FirstColumn) Apply(Part("firstColumn"), range with { Right = range.Left }, 4);
        if (table.Style.LastColumn) Apply(Part("lastColumn"), range with { Left = range.Right }, 5);
        if (table.HeaderRow)
        {
            Apply(Part("headerRow"), range with { Bottom = range.Top }, 6);
            Apply(Part("firstHeaderCell"), new(range.Top, range.Left, range.Top, range.Left), 7);
            Apply(Part("lastHeaderCell"), new(range.Top, range.Right, range.Top, range.Right), 8);
        }
        if (table.TotalsRow)
        {
            Apply(Part("totalRow"), range with { Top = range.Bottom }, 9);
            Apply(Part("firstTotalCell"), new(range.Bottom, range.Left, range.Bottom, range.Left), 10);
            Apply(Part("lastTotalCell"), new(range.Bottom, range.Right, range.Bottom, range.Right), 11);
        }
        var columnInfo = table.Columns[column - range.Left];
        Apply(table.HeaderRow && row == range.Top ? columnInfo.HeaderStyle : table.TotalsRow && row == range.Bottom ? columnInfo.TotalsStyle : columnInfo.DataStyle,
            new(row, column, row, column), 12);
        var normal = snapshot.Styles[0];
        if (direct.LeftBorder != normal.LeftBorder || direct.LeftBorderColor != normal.LeftBorderColor) ranks[0] = 100;
        if (direct.RightBorder != normal.RightBorder || direct.RightBorderColor != normal.RightBorderColor) ranks[1] = 100;
        if (direct.TopBorder != normal.TopBorder || direct.TopBorderColor != normal.TopBorderColor) ranks[2] = 100;
        if (direct.BottomBorder != normal.BottomBorder || direct.BottomBorderColor != normal.BottomBorderColor) ranks[3] = 100;
        // Cell formatting overrides only the properties it changes from the workbook's Normal style.
        return result with
        {
            Bold = direct.Bold != normal.Bold ? direct.Bold : result.Bold,
            FontColor = direct.FontColor != normal.FontColor ? direct.FontColor : result.FontColor,
            FillColor = direct.FillColor != normal.FillColor ? direct.FillColor : result.FillColor,
            NumberFormatId = direct.NumberFormatId != normal.NumberFormatId || direct.NumberFormatCode != normal.NumberFormatCode ? direct.NumberFormatId : result.NumberFormatId,
            NumberFormatCode = direct.NumberFormatId != normal.NumberFormatId || direct.NumberFormatCode != normal.NumberFormatCode ? direct.NumberFormatCode : result.NumberFormatCode,
            LeftBorder = direct.LeftBorder != normal.LeftBorder ? direct.LeftBorder : result.LeftBorder,
            LeftBorderColor = direct.LeftBorderColor != normal.LeftBorderColor ? direct.LeftBorderColor : result.LeftBorderColor,
            RightBorder = direct.RightBorder != normal.RightBorder ? direct.RightBorder : result.RightBorder,
            RightBorderColor = direct.RightBorderColor != normal.RightBorderColor ? direct.RightBorderColor : result.RightBorderColor,
            TopBorder = direct.TopBorder != normal.TopBorder ? direct.TopBorder : result.TopBorder,
            TopBorderColor = direct.TopBorderColor != normal.TopBorderColor ? direct.TopBorderColor : result.TopBorderColor,
            BottomBorder = direct.BottomBorder != normal.BottomBorder ? direct.BottomBorder : result.BottomBorder,
            BottomBorderColor = direct.BottomBorderColor != normal.BottomBorderColor ? direct.BottomBorderColor : result.BottomBorderColor
        };
    }
}
