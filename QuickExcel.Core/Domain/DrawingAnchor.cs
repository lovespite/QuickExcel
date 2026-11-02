namespace QuickExcel.Domain;

public enum DrawingAnchorKind { TwoCell, OneCell, Absolute }
public readonly record struct DrawingBounds(double X, double Y, double Width, double Height);

/// <summary>DrawingML coordinates are integer EMUs (9525 per 96-DPI pixel), independent of monitor DPI.</summary>
public sealed record DrawingAnchor
{
    public const long EmusPerPixel = 9525;
    public DrawingAnchorKind Kind { get; init; }
    public int Row { get; init; }
    public int Column { get; init; }
    public long RowOffset { get; init; }
    public long ColumnOffset { get; init; }
    public int EndRow { get; init; }
    public int EndColumn { get; init; }
    public long EndRowOffset { get; init; }
    public long EndColumnOffset { get; init; }
    public long X { get; init; }
    public long Y { get; init; }
    public long Width { get; init; }
    public long Height { get; init; }

    public void Validate()
    {
        if (!Enum.IsDefined(Kind)) throw new ArgumentException("图片锚点类型无效。");
        WorkbookLimits.ValidatePosition(Row, Column);
        if (new[] { RowOffset, ColumnOffset, EndRowOffset, EndColumnOffset, X, Y, Width, Height }.Any(v => v < 0 || v > 10_000_000_000_000L))
            throw new ArgumentException("绘图坐标超出预算。");
        if (Kind == DrawingAnchorKind.TwoCell && (EndRow < Row || EndColumn < Column || EndRow > WorkbookLimits.MaxRows || EndColumn > WorkbookLimits.MaxColumns ||
            EndRow == Row && EndRowOffset <= RowOffset || EndColumn == Column && EndColumnOffset <= ColumnOffset))
            throw new ArgumentException("绘图坐标反向或零尺寸：终点必须在起点右下方。");
        if (Kind != DrawingAnchorKind.TwoCell && (Width <= 0 || Height <= 0)) throw new ArgumentException("绘图宽高必须为正数。");
    }

    public DrawingBounds Bounds(WorksheetSnapshot sheet)
    {
        var x = Kind == DrawingAnchorKind.Absolute ? X / (double)EmusPerPixel : DrawingGeometry.ColumnStart(sheet, Column) + ColumnOffset / (double)EmusPerPixel;
        var y = Kind == DrawingAnchorKind.Absolute ? Y / (double)EmusPerPixel : DrawingGeometry.RowStart(sheet, Row) + RowOffset / (double)EmusPerPixel;
        return new(x, y, Kind == DrawingAnchorKind.TwoCell ? Math.Max(0, DrawingGeometry.ColumnStart(sheet, EndColumn) + EndColumnOffset / (double)EmusPerPixel - x) : Width / (double)EmusPerPixel,
            Kind == DrawingAnchorKind.TwoCell ? Math.Max(0, DrawingGeometry.RowStart(sheet, EndRow) + EndRowOffset / (double)EmusPerPixel - y) : Height / (double)EmusPerPixel);
    }
}

public static class DrawingGeometry
{
    public static double RowSize(WorksheetSnapshot s, int r) => s.Rows.GetValueOrDefault(r)?.Hidden == true ? 0 : s.Rows.GetValueOrDefault(r)?.Height ?? s.DefaultRowHeight ?? 31;
    public static double ColumnSize(WorksheetSnapshot s, int c) => s.Columns.GetValueOrDefault(c)?.Hidden == true ? 0 : Math.Floor((s.Columns.GetValueOrDefault(c)?.Width ?? s.Data.DefaultColumnWidth ?? WorkbookLimits.DefaultColumnWidth(c)) * 7 + 5);
    public static double RowStart(WorksheetSnapshot s, int r) => r * (s.DefaultRowHeight ?? 31) + s.Rows.Where(p => p.Key < r).Sum(p => RowSize(s, p.Key) - (s.DefaultRowHeight ?? 31));
    public static double ColumnStart(WorksheetSnapshot s, int c)
    {
        double Default(int index) => Math.Floor((s.Data.DefaultColumnWidth ?? WorkbookLimits.DefaultColumnWidth(index)) * 7 + 5);
        return (c == 0 ? 0 : Default(0) + (c - 1) * Default(1)) + s.Columns.Where(p => p.Key < c).Sum(p => ColumnSize(s, p.Key) - Default(p.Key));
    }
    public static long Emus(double pixels) => checked((long)Math.Round(pixels * DrawingAnchor.EmusPerPixel));

    public static DrawingAnchor FromBounds(WorksheetSnapshot sheet, DrawingBounds bounds, DrawingAnchorKind kind)
    {
        if (!new[] { bounds.X, bounds.Y, bounds.Width, bounds.Height }.All(double.IsFinite) || bounds.X < 0 || bounds.Y < 0 || bounds.Width <= 0 || bounds.Height <= 0)
            throw new ArgumentException("绘图位置和尺寸无效。");
        (int Index, long Offset) Locate(double value, bool row)
        {
            var limit = row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
            for (var i = 0; i < limit; i++)
            {
                var size = row ? RowSize(sheet, i) : ColumnSize(sheet, i);
                if (size > value) return (i, Emus(value));
                value -= size;
            }
            if (value == 0) return (limit, 0);
            throw new ArgumentException("绘图超出工作表坐标边界。");
        }
        DrawingAnchor anchor;
        if (kind == DrawingAnchorKind.Absolute) anchor = new() { Kind = kind, X = Emus(bounds.X), Y = Emus(bounds.Y), Width = Emus(bounds.Width), Height = Emus(bounds.Height) };
        else
        {
            var row = Locate(bounds.Y, true); var column = Locate(bounds.X, false);
            anchor = new() { Kind = kind, Row = row.Index, Column = column.Index, RowOffset = row.Offset, ColumnOffset = column.Offset };
            if (kind == DrawingAnchorKind.OneCell) anchor = anchor with { Width = Emus(bounds.Width), Height = Emus(bounds.Height) };
            else
            {
                var endRow = Locate(bounds.Y + bounds.Height, true); var endColumn = Locate(bounds.X + bounds.Width, false);
                anchor = anchor with { EndRow = endRow.Index, EndColumn = endColumn.Index, EndRowOffset = endRow.Offset, EndColumnOffset = endColumn.Offset };
            }
        }
        anchor.Validate(); return anchor;
    }
}
