using QuickExcel.Core;
using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

public sealed partial class SkiaRenderer
{
    public void DrawChart(SKCanvas canvas, RenderRect rectangle, WorkbookSnapshot book, WorksheetChart chart, bool monochrome = false)
    {
        var b = Rect(rectangle); var series = ChartData.Read(book, chart); var count = series.Max(s => s.Values.Length);
        using var paint = new SKPaint { IsAntialias = true };
        SKColor SeriesColor(string value)
        {
            var color = Color(value, SKColors.SteelBlue);
            if (!monochrome) return color;
            var gray = (byte)((color.Red * 299 + color.Green * 587 + color.Blue * 114) / 1000); return new(gray, gray, gray);
        }
        void Fill(SKRect rect, SKColor color) { paint.Color = color; paint.Style = SKPaintStyle.Fill; canvas.DrawRect(rect, paint); }
        void Line(float x1, float y1, float x2, float y2, SKColor color, float width = 1)
        { paint.Color = color; paint.StrokeWidth = width; canvas.DrawLine(x1, y1, x2, y2, paint); }
        void Label(string text, float x, float y, float width, float size = 11, int align = 1) =>
            DrawText(canvas, new(text, new() { FontSize = size * .75, Horizontal = align, Vertical = 0 }), SKRect.Create(x, y, Math.Max(1, width), size * 2), monochrome, $"chart:{chart.Title}");
        canvas.Save(); canvas.ClipRect(b);
        try
        {
            Fill(b, SKColors.White); paint.Style = SKPaintStyle.Stroke; paint.Color = SKColors.LightGray; canvas.DrawRect(b, paint); paint.Style = SKPaintStyle.Fill;
            Label(chart.Title, b.Left + 12, b.Top + 10, b.Width - 24, 16, 2);
            var lw = chart.Legend is "r" or "l" ? Math.Min(130, b.Width * .25f) : 0;
            var lh = chart.Legend is "t" or "b" ? 32 : 0;
            var p = SKRect.Create(b.Left + 54 + (chart.Legend == "l" ? lw : 0), b.Top + 42 + (chart.Legend == "t" ? lh : 0), Math.Max(1, b.Width - 74 - lw), Math.Max(1, b.Height - 78 - lh));
            var labels = chart.Kind == WorksheetChartKind.Pie ? series[0].Categories : series.Select(s => s.Name).ToArray();
            var colors = chart.Kind == WorksheetChartKind.Pie ? series[0].PointColors : series.Select(s => s.Color).ToArray();
            if (chart.Legend is not null)
                for (var i = 0; i < labels.Length; i++)
                {
                    var horizontal = lh > 0; var w = horizontal ? (b.Width - 20) / Math.Max(1, labels.Length) : Math.Max(20, lw - 12);
                    var x = horizontal ? b.Left + 10 + i * w : chart.Legend == "l" ? b.Left + 8 : b.Right - lw + 6;
                    var y = horizontal ? chart.Legend == "t" ? b.Top + 35 : b.Bottom - 23 : p.Top + i * 20;
                    if (y + 16 > b.Bottom - 4) break;
                    Fill(SKRect.Create(x, y + 3, 9, 9), SeriesColor(colors[i])); Label(labels[i], x + 14, y, w - 14);
                }
            if (count == 0) { Label("没有可显示的数据", p.Left, p.Top, p.Width); return; }
            if (chart.Kind == WorksheetChartKind.Pie)
            {
                var values = series[0].Values.Select(v => Math.Abs(v ?? 0)).ToArray(); var largest = values.DefaultIfEmpty().Max();
                var total = largest == 0 ? 0 : values.Sum(v => v / largest); var angle = -90f; var radius = Math.Min(p.Width, p.Height) / 2;
                var oval = SKRect.Create(p.MidX - radius, p.MidY - radius, radius * 2, radius * 2);
                for (var i = 0; i < values.Length && total > 0; i++)
                {
                    var sweep = (float)(values[i] / largest / total * 360); if (sweep == 0) continue;
                    paint.Color = SeriesColor(series[0].PointColors[i]); canvas.DrawArc(oval, angle, sweep, true, paint); angle += sweep;
                }
                return;
            }
            var all = series.SelectMany(s => s.Values).OfType<double>().ToArray();
            var minimum = chart.Minimum ?? Math.Min(0, all.DefaultIfEmpty().Min()); var maximum = chart.Maximum ?? Math.Max(0, all.DefaultIfEmpty().Max());
            if (minimum >= maximum) { if (chart.Maximum is null) maximum = minimum + Math.Max(1, Math.Abs(minimum) * .1); else minimum = maximum - Math.Max(1, Math.Abs(maximum) * .1); }
            if (!double.IsFinite(minimum) || !double.IsFinite(maximum) || minimum >= maximum) throw new InvalidOperationException("自动坐标轴超出数值范围，请明确设置可显示的上下限。");
            var magnitude = Math.Max(Math.Abs(minimum), Math.Abs(maximum)); var range = maximum / magnitude - minimum / magnitude;
            float Y(double value) => (float)(p.Bottom - (value / magnitude - minimum / magnitude) / range * p.Height);
            var baseline = Y(Math.Clamp(0, minimum, maximum));
            for (var i = 0; i <= 5; i++)
            {
                var value = minimum * (1 - i / 5d) + maximum * (i / 5d); var y = Y(value);
                Line(p.Left, y, p.Right, y, SKColors.LightGray, .5f);
                Label(CellDisplayFormatter.Format(CellValue.FromNumber(value), new() { NumberFormatCode = chart.NumberFormat }, book.DateSystem), b.Left + (chart.Legend == "l" ? lw : 0), y - 7, 48, 10, 3);
            }
            Line(p.Left, p.Top, p.Left, p.Bottom, SKColors.Gray); Line(p.Left, baseline, p.Right, baseline, SKColors.Gray);
            var step = p.Width / count; var labelStep = Math.Max(1, (int)Math.Ceiling(60 / step));
            for (var i = 0; i < series[0].Categories.Length; i += labelStep) Label(series[0].Categories[i], p.Left + i * step, p.Bottom + 6, step * labelStep, 10, 2);
            canvas.Save(); canvas.ClipRect(p);
            for (var s = 0; s < series.Count; s++)
            {
                SKPoint? previous = null;
                for (var i = 0; i < series[s].Values.Length; i++)
                {
                    if (series[s].Values[i] is not { } value) { previous = null; continue; }
                    var x = p.Left + (i + .5f) * step; var y = Y(value); var color = SeriesColor(series[s].Color);
                    if (chart.Kind == WorksheetChartKind.Column)
                    { var width = step / (series.Count + 1.5f); Fill(SKRect.Create(x - series.Count * width / 2 + s * width, Math.Min(y, baseline), width, Math.Abs(baseline - y)), color); }
                    else { if (previous is { } old) Line(old.X, old.Y, x, y, color, 2); previous = new(x, y); }
                }
            }
            canvas.Restore();
        }
        finally { canvas.Restore(); }
    }
}
