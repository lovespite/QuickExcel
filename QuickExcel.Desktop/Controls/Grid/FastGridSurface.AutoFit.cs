using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed partial class FastGridSurface
{
    /// <summary>Measure displayed, unmerged content on the entire visible axis, including offscreen cells.</summary>
    internal double MeasureAutoFit(SheetAxis axis, int index)
    {
        var sheet = Sheet;
        var isRow = axis == SheetAxis.Row;
        int Other(CellPosition position) => isRow ? position.Column : position.Row;
        // Sort only stored cells in the requested axis. Walk merge intervals alongside
        // them, avoiding a scan of every empty cell or every merge for each value.
        var cells = sheet.Cells.Where(p => (isRow ? p.Key.Row : p.Key.Column) == index &&
            !_model.HiddenRows.Contains(p.Key.Row) && !_model.HiddenColumns.Contains(p.Key.Column)).OrderBy(p => Other(p.Key));
        var merges = sheet.MergedCells.Select(CellRectangle.Parse)
            .Where(m => isRow ? index >= m.Top && index <= m.Bottom : index >= m.Left && index <= m.Right)
            .Select(m => (First: isRow ? m.Left : m.Top, Last: isRow ? m.Right : m.Bottom)).OrderBy(m => m.First).ToArray();
        var mergeIndex = 0;
        var maximum = 0d;
        var scaleX = Control.RenderScaleX;
        var scaleY = Control.RenderScaleY;
        var cache = new Dictionary<(string Text, StyleDefinition Style, double Width, RichCellText? Rich), double>();
        foreach (var (position, _) in cells)
        {
            var other = Other(position);
            while (mergeIndex < merges.Length && merges[mergeIndex].Last < other) mergeIndex++;
            if (mergeIndex < merges.Length && merges[mergeIndex].First <= other) continue;
            var value = DisplayText(position.Row, position.Column);
            if (value.Length == 0) continue;
            var style = ConditionalPresentation.GetStyle(_snapshot, sheet, position.Row, position.Column);
            style = style with { FontName = style.FontName ?? Control.CellFontName };
            var width = isRow ? sheet.Columns.GetValueOrDefault(position.Column)?.Width is { } units
                ? Math.Clamp(units * 7 + 5, 1, 10000) : position.Column == 0 ? 160 : 112 : 0;
            var rich = sheet.GetCell(position.Row, position.Column)?.Content.RichText;
            var key = (value, style, width, rich);
            if (!cache.TryGetValue(key, out var size))
            {
                using var text = Artwork.Renderer.Layout(new(value, style,
                    isRow ? Math.Max(1, Math.Round(width * scaleX) / scaleX - 4) : double.PositiveInfinity,
                    rich, Control.CellFontSize));
                size = isRow ? text.Metrics.RotatedHeight + 3 : text.Metrics.RotatedWidth + text.Indent + 6;
                if (cache.Count < 4096) cache[key] = size;
            }
            maximum = Math.Max(maximum, size);
            if (maximum >= 10000) break;
        }
        if (maximum == 0) return Math.Clamp(isRow ? sheet.DefaultRowHeight ?? 31 : index == 0 ? 160 : 112, isRow ? 1 : 12, 10000);
        var scale = isRow ? scaleY : scaleX;
        return Math.Clamp(Math.Ceiling(maximum * scale) / scale, isRow ? 1 : 12, 10000);
    }
}
