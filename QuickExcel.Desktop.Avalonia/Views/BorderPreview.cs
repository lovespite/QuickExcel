using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Rendering;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class BorderPreview : Control
{
    public BorderStroke?[] Lines { get; set; } = new BorderStroke?[8];
    public bool SampleText { get; init; }
    public Action<int>? Toggle { get; set; }
    public BorderPreview()
    {
        PointerPressed += (_, e) =>
        {
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
            var p = e.GetPosition(this); var r = DrawingBounds; var x = (p.X - r.Left) / r.Width; var y = (p.Y - r.Top) / r.Height;
            var candidates = new[] { Math.Abs(x), Math.Abs(x - 1), Math.Abs(y), Math.Abs(y - 1), Math.Abs(y - .5), Math.Abs(x - .5), Math.Abs(x + y - 1) / 1.4, Math.Abs(x - y) / 1.4 };
            Toggle?.Invoke(Array.IndexOf(candidates, candidates.Min())); e.Handled = true;
        };
    }
    private Rect DrawingBounds => new Rect(Bounds.Size).Deflate(SampleText ? 18 : 3);
    public override void Render(DrawingContext context)
    {
        base.Render(context); var r = DrawingBounds;
        if (SampleText) context.DrawRectangle(Brushes.White, new Pen(Brushes.LightGray), new Rect(Bounds.Size));
        var lines = new (Point A, Point B)[] { (r.TopLeft, r.BottomLeft), (r.TopRight, r.BottomRight), (r.TopLeft, r.TopRight), (r.BottomLeft, r.BottomRight),
            (new(r.Left, r.Center.Y), new(r.Right, r.Center.Y)), (new(r.Center.X, r.Top), new(r.Center.X, r.Bottom)), (r.BottomLeft, r.TopRight), (r.TopLeft, r.BottomRight) };
        for (var i = 0; i < lines.Length; i++)
        {
            var stroke = Lines[i];
            if (stroke?.Style == 0 && i >= 6) continue;
            if (stroke is null || stroke.Style == 0)
            { if (i < 6) context.DrawLine(new Pen(stroke is null ? Brushes.Gray : Brushes.LightGray, .7, new DashStyle([2d, 3d], 0)), lines[i].A, lines[i].B); }
            else DrawLine(context, stroke, lines[i].A, lines[i].B);
        }
        if (SampleText)
        {
            var text = new FormattedText("文本", System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Microsoft YaHei UI"), 14, Brushes.Black);
            foreach (var x in new[] { .25, .75 }) foreach (var y in new[] { .25, .75 })
                context.DrawText(text, new(r.Left + r.Width * x - text.Width / 2, r.Top + r.Height * y - text.Height / 2));
        }
    }
    internal static void DrawLine(DrawingContext context, BorderStroke stroke, Point a, Point b)
    {
        if (stroke.Style is 1 or 3 or 4 or 5 or 6 or 9 or 11)
        {
            if (a.X == b.X) { a = a.WithX(a.X + .5); b = b.WithX(b.X + .5); }
            else if (a.Y == b.Y) { a = a.WithY(a.Y + .5); b = b.WithY(b.Y + .5); }
        }
        var width = BorderMetrics.Width(stroke.Style); var color = new SolidColorBrush(Color.Parse("#" + (stroke.Color ?? "FF000000")));
        var dashes = BorderMetrics.Dashes(stroke.Style);
        var pen = new Pen(color, width, dashes is null ? null : new DashStyle(dashes.Select(d => (double)d / width), 0));
        if (stroke.Style == 6)
        {
            var v = new Vector(b.X - a.X, b.Y - a.Y); var offset = new Vector(-v.Y, v.X) / v.Length;
            context.DrawLine(pen, a + offset, b + offset); context.DrawLine(pen, a - offset, b - offset);
        }
        else if (stroke.Style != 0) context.DrawLine(pen, a, b);
    }
}

internal sealed class BorderLineSample(int style) : Control
{
    public override void Render(DrawingContext context) => BorderPreview.DrawLine(context, new(style), new(3, Bounds.Height / 2), new(Bounds.Width - 3, Bounds.Height / 2));
}
