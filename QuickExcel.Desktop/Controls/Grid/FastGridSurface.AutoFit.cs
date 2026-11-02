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
        var defaultStyle = new StyleDefinition();
        var pixelsPerDip = Math.Max(Control.RenderScaleY, DpiDetector.DpiYKoef);
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
            var styleId = sheet.GetStyleId(position.Row, position.Column);
            var style = ConditionalPresentation.GetStyle(_snapshot, sheet, position.Row, position.Column);
            var width = isRow ? sheet.Columns.GetValueOrDefault(position.Column)?.Width is { } units
                ? Math.Clamp(units * 7 + 5, 1, 10000) : position.Column == 0 ? 160 : 112 : 0;
            var rich = sheet.GetCell(position.Row, position.Column)?.Content.RichText;
            var key = (value, style, width, rich);
            if (!cache.TryGetValue(key, out var size))
            {
                if (style == defaultStyle && rich is null)
                {
                    // Default cells use the bitmap glyph renderer, whose rounded
                    // advances differ from FormattedText for an entire string.
                    var font = Control.GetFont(false, false);
                    size = isRow ? (font.GetTextHeight(value) + 2 * Control.CellPaddingVertical + 2) / scaleY :
                        (font.GetTextWidth(value, (int)(10000 * scaleX)) + 2 * Control.CellPaddingHorizontal + 2) / scaleX;
                }
                else
                {
                    var text = FastGridArtwork.CreateText(value, style, pixelsPerDip, rich, Control.CellFontSize);
                    if (isRow && style.TextRotation == 0)
                    {
                        text.MaxTextWidth = Math.Max(1, Math.Round(width * scaleX) / scaleX - 4);
                        text.MaxLineCount = style.WrapText ? int.MaxValue : 1;
                        text.Trimming = System.Windows.TextTrimming.CharacterEllipsis;
                    }
                    var bounds = new System.Windows.Rect(0, 0, text.WidthIncludingTrailingWhitespace, text.Height);
                    if (style.TextRotation is > 0 and <= 180)
                        bounds = new RotateTransform(style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90).TransformBounds(bounds);
                    size = isRow ? bounds.Height + 3 : bounds.Width + FastGridArtwork.IndentWidth(style, pixelsPerDip, Control.CellFontSize) + 6;
                }
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
