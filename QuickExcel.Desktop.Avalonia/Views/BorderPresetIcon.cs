using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Rendering;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Compact border samples with dotted cell guides and solid selected edges.</summary>
internal sealed class BorderPresetIcon : Control
{
    static BorderPresetIcon() => AffectsRender<BorderPresetIcon>(TextElement.ForegroundProperty);

    private CellBorderPatch _patch;
    public CellBorderPatch Patch
    {
        get => _patch;
        set { _patch = value; InvalidateVisual(); }
    }

    public BorderPresetIcon(CellBorderPatch patch)
    {
        _patch = patch; Width = Height = 20;
        VerticalAlignment = VerticalAlignment.Center; IsHitTestVisible = false; Focusable = false;
    }

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        var r = new Rect(1.5, 1.5, 16, 16);
        var lines = new (Point A, Point B, BorderStroke? Stroke)[]
        {
            (r.TopLeft, r.BottomLeft, Patch.Left), (r.TopRight, r.BottomRight, Patch.Right),
            (r.TopLeft, r.TopRight, Patch.Top), (r.BottomLeft, r.BottomRight, Patch.Bottom),
            (new(r.Left, r.Center.Y), new(r.Right, r.Center.Y), Patch.InsideHorizontal),
            (new(r.Center.X, r.Top), new(r.Center.X, r.Bottom), Patch.InsideVertical)
        };
        var foreground = TextElement.GetForeground(this);
        var guide = new SolidColorBrush(Color.Parse("#888888"));
        foreach (var (a, b, stroke) in lines)
        {
            if (stroke is { Style: > 0 }) continue;
            var step = (b - a) / 8;
            for (var i = 0; i <= 8; i++)
            {
                var p = a + step * i;
                context.DrawRectangle(guide, null, new Rect(p.X - .5, p.Y - .5, 1, 1));
            }
        }
        void Draw(Point a, Point b, BorderStroke? stroke)
        {
            if (stroke is null || stroke.Style == 0) return;
            var width = BorderMetrics.Width(stroke.Style);
            var dashes = BorderMetrics.Dashes(stroke.Style);
            var pen = new Pen(foreground, width, dashes is null ? null : new DashStyle(dashes.Select(d => (double)d / width), 0));
            if (stroke.Style == 6)
            {
                var direction = new Vector(b.X - a.X, b.Y - a.Y); var offset = new Vector(-direction.Y, direction.X) / direction.Length;
                context.DrawLine(pen, a + offset, b + offset); context.DrawLine(pen, a - offset, b - offset);
            }
            else context.DrawLine(pen, a, b);
        }
        foreach (var (a, b, stroke) in lines) Draw(a, b, stroke);
        if (Patch.DiagonalUp == true) Draw(r.BottomLeft, r.TopRight, Patch.Diagonal);
        if (Patch.DiagonalDown == true) Draw(r.TopLeft, r.BottomRight, Patch.Diagonal);
    }
}
