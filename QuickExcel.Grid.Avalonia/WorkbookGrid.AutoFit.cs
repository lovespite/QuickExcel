using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    /// <summary>Measure displayed, unmerged content on the entire visible axis, including offscreen cells.</summary>
    public double MeasureAutoFit(SheetAxis axis, int index)
    {
        var sheet = Sheet;
        var isRow = axis == SheetAxis.Row;
        int Other(CellPosition position) => isRow ? position.Column : position.Row;
        // Sort only stored cells in the requested axis. Walk merge intervals alongside
        // them, avoiding a scan of every empty cell or every merge for each value.
        var cells = sheet.Cells.Where(p => (isRow ? p.Key.Row : p.Key.Column) == index &&
            _rows.Cell(p.Key.Row) is not null && _columns.Cell(p.Key.Column) is not null).OrderBy(p => Other(p.Key));
        var merges = sheet.MergedCells.Select(CellRectangle.Parse)
            .Where(m => isRow ? index >= m.Top && index <= m.Bottom : index >= m.Left && index <= m.Right)
            .Select(m => (First: isRow ? m.Left : m.Top, Last: isRow ? m.Right : m.Bottom)).OrderBy(m => m.First).ToArray();
        var mergeIndex = 0;
        var maximum = 0d;
        var scaleX = (TopLevel.GetTopLevel(this)?.RenderScaling ?? 1) * _zoom;
        var scaleY = scaleX;
        var cache = new Dictionary<(string Text, StyleDefinition Style, double Width, RichCellText? Rich), double>();
        foreach (var (position, _) in cells)
        {
            var other = Other(position);
            while (mergeIndex < merges.Length && merges[mergeIndex].Last < other) mergeIndex++;
            if (mergeIndex < merges.Length && merges[mergeIndex].First <= other) continue;
            var style = ResolveStyle?.Invoke(_snapshot!, sheet, position.Row, position.Column) ?? _snapshot!.Styles[sheet.GetStyleId(position.Row, position.Column)];
            var value = CellDisplayFormatter.Format(sheet.GetCell(position.Row, position.Column)!.Value, style, _snapshot!.DateSystem);
            if (value.Length == 0) continue;
            style = style with { FontName = style.FontName ?? "Microsoft YaHei UI" };
            var width = isRow ? DrawingGeometry.ColumnSize(sheet, position.Column) : 0;
            var rich = sheet.GetCell(position.Row, position.Column)?.Content.RichText;
            var key = (value, style, width, rich);
            if (!cache.TryGetValue(key, out var size))
            {
                using var text = Renderer.Layout(new(value, style,
                    isRow ? Math.Max(1, Math.Round(width * scaleX) / scaleX - 4) : double.PositiveInfinity,
                    rich, 14));
                size = isRow ? text.Metrics.RotatedHeight + 3 : text.Metrics.RotatedWidth + text.Indent + 6;
                if (cache.Count < 4096) cache[key] = size;
            }
            maximum = Math.Max(maximum, size);
            if (maximum >= 10000) break;
        }
        if (maximum == 0) return Math.Clamp(isRow ? sheet.DefaultRowHeight ?? 31 : DrawingGeometry.ColumnSize(sheet, index), isRow ? 1 : 12, 10000);
        var scale = isRow ? scaleY : scaleX;
        return Math.Clamp(Math.Ceiling(maximum * scale) / scale, isRow ? 1 : 12, 10000);
    }
}
