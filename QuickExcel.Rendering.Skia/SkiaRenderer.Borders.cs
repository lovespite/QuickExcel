using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

public sealed partial class SkiaRenderer
{
    /// <summary>Draw after all cell backgrounds; centered strokes share a single worksheet edge.</summary>
    public static void DrawBorders(SKCanvas canvas, RenderRect bounds, StyleDefinition style, bool monochrome = false)
    {
        var r = Rect(bounds);
        DrawBorderLine(canvas, style.LeftBorder, style.LeftBorderColor, new(r.Left, r.Top), new(r.Left, r.Bottom), monochrome);
        DrawBorderLine(canvas, style.RightBorder, style.RightBorderColor, new(r.Right, r.Top), new(r.Right, r.Bottom), monochrome);
        DrawBorderLine(canvas, style.TopBorder, style.TopBorderColor, new(r.Left, r.Top), new(r.Right, r.Top), monochrome);
        DrawBorderLine(canvas, style.BottomBorder, style.BottomBorderColor, new(r.Left, r.Bottom), new(r.Right, r.Bottom), monochrome);
        if (style.DiagonalDown) DrawBorderLine(canvas, style.DiagonalBorder, style.DiagonalBorderColor, new(r.Left, r.Top), new(r.Right, r.Bottom), monochrome);
        if (style.DiagonalUp) DrawBorderLine(canvas, style.DiagonalBorder, style.DiagonalBorderColor, new(r.Left, r.Bottom), new(r.Right, r.Top), monochrome);
    }

    public static void DrawBorderLine(SKCanvas canvas, int kind, string? color, SKPoint from, SKPoint to, bool monochrome = false)
    {
        if (kind == 0 || from == to) return;
        using var paint = new SKPaint { Color = monochrome ? SKColors.Black : Color(color, SKColors.Black), IsAntialias = true,
            Style = SKPaintStyle.Stroke, StrokeWidth = BorderMetrics.Width(kind) };
        // Align odd-width orthogonal strokes to pixel centers; double lines retain a clear gap.
        if (kind is 1 or 3 or 4 or 5 or 6 or 9 or 11)
        {
            if (from.X == to.X) { from.X += .5f; to.X += .5f; }
            else if (from.Y == to.Y) { from.Y += .5f; to.Y += .5f; }
        }
        var dx = to.X - from.X; var dy = to.Y - from.Y; var length = MathF.Sqrt(dx * dx + dy * dy);
        var normal = new SKPoint(-dy / length, dx / length);
        if (kind == 6)
        {
            canvas.DrawLine(from + normal, to + normal, paint); canvas.DrawLine(from - normal, to - normal, paint); return;
        }
        if (kind == 13)
        {
            // A skewed dash/dot pattern, with the same orientation on either side of a shared edge.
            SKPoint Along(float distance) => new(dx / length * distance, dy / length * distance);
            using var path = new SKPathBuilder();
            for (float p = 0; p < length; p += 13)
            {
                void Segment(float start, float end)
                {
                    if (start >= length) return;
                    var a = from + Along(start); var b = from + Along(Math.Min(end, length));
                    path.MoveTo(a - normal); path.LineTo(b - normal); path.LineTo(b + normal + Along(-2)); path.LineTo(a + normal + Along(-2)); path.Close();
                }
                Segment(p, p + 7); Segment(p + 10, p + 12);
            }
            paint.Style = SKPaintStyle.Fill; using var geometry = path.Detach(); canvas.DrawPath(geometry, paint); return;
        }
        using var dash = BorderMetrics.Dashes(kind) is { } pattern ? SKPathEffect.CreateDash(pattern, 0) : null;
        paint.PathEffect = dash; canvas.DrawLine(from, to, paint);
    }
}
