using System.Windows;
using System.Windows.Media;

namespace QuickExcel.Desktop;

internal sealed partial class FastGridSurface
{
    private ReferenceHighlight? _referenceHighlight;
    private GridRange? _referenceRange;

    private void InstallReferenceHighlight()
    {
        _referenceHighlight = new ReferenceHighlight(this) { IsHitTestVisible = false };
        _overlay.Children.Insert(1, _referenceHighlight);
    }

    internal GridRange? ReferenceRange => _referenceRange;

    /// <summary>Highlights the cell/range a formula currently refers to without touching document selection.</summary>
    internal void ShowReference(GridRange? range)
    {
        if (_referenceRange == range) return;
        _referenceRange = range;
        _referenceHighlight?.InvalidateVisual();
    }

    private sealed class ReferenceHighlight(FastGridSurface surface) : FrameworkElement
    {
        private static readonly Brush Fill = new SolidColorBrush(Color.FromArgb(46, 24, 117, 83));
        private static readonly Pen Outline = new(new SolidColorBrush(Color.FromArgb(255, 24, 117, 83)), 2);

        static ReferenceHighlight()
        {
            Fill.Freeze(); Outline.Freeze();
        }

        protected override void OnRender(DrawingContext drawing)
        {
            if (surface._referenceRange is not { } range || surface.RowCount <= 0 || surface.ColumnCount <= 0) return;
            // A concrete reference outside the displayed model is not evidence for a cell at the
            // edge; draw nothing rather than moving =Z1 onto column L.
            if (range.Bottom < 0 || range.Right < 0 || range.Top >= surface.RowCount || range.Left >= surface.ColumnCount) return;
            var rows = surface.Control.GetVisibleModelRows();
            var columns = surface.Control.GetVisibleModelColumns();
            var top = Math.Max(range.Top, 0); var bottom = Math.Min(range.Bottom, surface.RowCount - 1);
            var left = Math.Max(range.Left, 0); var right = Math.Min(range.Right, surface.ColumnCount - 1);
            // Whole-column/row references cover the entire model; clip them to the current
            // viewport so the outline remains visible instead of collapsing at the top-left.
            if (rows.Count > 0) { top = Math.Max(top, rows.Min()); bottom = Math.Min(bottom, rows.Max()); }
            if (columns.Count > 0) { left = Math.Max(left, columns.Min()); right = Math.Min(right, columns.Max()); }
            if (top > bottom || left > right) return;
            foreach (var fragment in surface.Control.GetModelRangeFragments(new(top, left, bottom, right)))
            {
                var visible = Rect.Intersect(fragment.Bounds, fragment.Clip);
                if (visible.Width <= 0 || visible.Height <= 0) continue;
                drawing.PushClip(new RectangleGeometry(fragment.Clip));
                drawing.DrawRectangle(Fill, null, fragment.Bounds);
                drawing.DrawRectangle(null, Outline, fragment.Bounds);
                drawing.Pop();
            }
        }
    }
}
