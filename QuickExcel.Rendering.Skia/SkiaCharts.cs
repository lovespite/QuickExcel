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
        SKColor ChartColor(string value)
        {
            var color = Color(value, SKColors.SteelBlue);
            if (!monochrome) return color;
            var gray = (byte)((color.Red * 299 + color.Green * 587 + color.Blue * 114) / 1000); return new(gray, gray, gray, color.Alpha);
        }
        void Fill(SKRect rect, SKColor color) { paint.Color = color; paint.Style = SKPaintStyle.Fill; canvas.DrawRect(rect, paint); }
        void Line(float x1, float y1, float x2, float y2, SKColor color, float width = 1)
        { paint.Color = color; paint.StrokeWidth = width; paint.Style = SKPaintStyle.Stroke; canvas.DrawLine(x1, y1, x2, y2, paint); paint.Style = SKPaintStyle.Fill; }
        void Label(string text, float x, float y, float width, ChartTextStyle? style = null, int align = 1)
        {
            style ??= chart.Appearance.Axis;
            DrawText(canvas, new(text, new() { FontSize = style.Size, FontName = style.FontFamily, FontColor = style.Color, Bold = style.Bold, Italic = style.Italic, Horizontal = align, Vertical = 0 }),
                SKRect.Create(x, y, Math.Max(1, width), (float)style.Size * 3), monochrome, $"chart:{chart.Title}");
        }
        string Number(double value) => CellDisplayFormatter.Format(CellValue.FromNumber(value), new() { NumberFormatCode = chart.NumberFormat }, book.DateSystem);
        void DataLabel(int s, int i, float x, float y, float width)
        {
            var options = chart.Series[s]; var data = series[s]; var parts = new List<string>();
            if (options.ShowSeriesName) parts.Add(data.Name); if (options.ShowCategories) parts.Add(data.Categories[i]);
            if (options.ShowValues && data.Values[i] is { } value) parts.Add(Number(value));
            if (parts.Count > 0) Label(string.Join(" · ", parts), x, y, width, align: 2);
        }
        void Marker(SKPoint point, ChartSeries options, SKColor color)
        {
            if (options.Marker == "none") return;
            var r = (float)options.MarkerSize * 2 / 3; var x = point.X; var y = point.Y;
            paint.Color = color; paint.Style = SKPaintStyle.Fill;
            if (options.Marker is "circle" or "auto" or "dot") { canvas.DrawCircle(x, y, options.Marker == "dot" ? r / 2 : r, paint); return; }
            if (options.Marker == "square") { canvas.DrawRect(SKRect.Create(x - r, y - r, r * 2, r * 2), paint); return; }
            if (options.Marker is "diamond" or "triangle")
            { using var path = new SKPathBuilder(); path.MoveTo(x, y - r); path.LineTo(x + r, y + (options.Marker == "triangle" ? r : 0)); if (options.Marker == "diamond") path.LineTo(x, y + r); path.LineTo(x - r, y + (options.Marker == "triangle" ? r : 0)); path.Close(); using var built = path.Detach(); canvas.DrawPath(built, paint); return; }
            if (options.Marker is "x" or "star") { Line(x - r, y - r, x + r, y + r, color, 1.5f); Line(x - r, y + r, x + r, y - r, color, 1.5f); }
            if (options.Marker is "plus" or "star" or "dash") { Line(x - r, y, x + r, y, color, 1.5f); if (options.Marker != "dash") Line(x, y - r, x, y + r, color, 1.5f); }
        }
        canvas.Save(); canvas.RotateDegrees((float)chart.Anchor.Rotation, b.MidX, b.MidY); canvas.ClipRect(b);
        try
        {
            Fill(b, ChartColor(chart.Appearance.Background)); paint.Style = SKPaintStyle.Stroke; paint.Color = ChartColor(chart.Appearance.BorderColor); paint.StrokeWidth = (float)chart.Appearance.BorderWidth;
            if (paint.StrokeWidth > 0) canvas.DrawRect(b, paint); paint.Style = SKPaintStyle.Fill;
            Label(chart.Title, b.Left + 12, b.Top + 10, b.Width - 24, chart.Appearance.Title, 2);
            var lw = chart.Legend is "r" or "l" ? Math.Min(130, b.Width * .25f) : 0; var lh = chart.Legend is "t" or "b" ? 32 : 0;
            var p = SKRect.Create(b.Left + 62 + (chart.Legend == "l" ? lw : 0), b.Top + 42 + (chart.Legend == "t" ? lh : 0), Math.Max(1, b.Width - 82 - lw), Math.Max(1, b.Height - 82 - lh));
            Fill(p, ChartColor(chart.Appearance.PlotBackground));
            var labels = chart.Kind is WorksheetChartKind.Pie or WorksheetChartKind.Funnel ? series[0].Categories : chart.Kind == WorksheetChartKind.Waterfall ? new[] { "增加", "减少", "总计" } : series.Select(s => s.Name).ToArray();
            var colors = chart.Kind is WorksheetChartKind.Pie or WorksheetChartKind.Funnel ? series[0].PointColors : chart.Kind == WorksheetChartKind.Waterfall ? new[] { chart.IncreaseColor, chart.DecreaseColor, chart.TotalColor } : series.Select(s => s.Color).ToArray();
            if (chart.Legend is not null)
                for (var i = 0; i < labels.Length; i++)
                {
                    var horizontal = lh > 0; var w = horizontal ? (b.Width - 20) / Math.Max(1, labels.Length) : Math.Max(20, lw - 12);
                    var x = horizontal ? b.Left + 10 + i * w : chart.Legend == "l" ? b.Left + 8 : b.Right - lw + 6;
                    var y = horizontal ? chart.Legend == "t" ? b.Top + 35 : b.Bottom - 25 : p.Top + i * 20;
                    if (y + 16 > b.Bottom - 4) break;
                    Fill(SKRect.Create(x, y + 3, 9, 9), ChartColor(colors[i])); Label(labels[i], x + 14, y, w - 14, chart.Appearance.Legend);
                }
            if (count == 0 || !series.Any(s => s.Values.Any(v => v.HasValue))) { Label("没有可显示的数据", p.Left, p.Top, p.Width); return; }
            if (chart.Kind == WorksheetChartKind.Pie)
            {
                var values = series[0].Values.Select(v => Math.Abs(v ?? 0)).ToArray(); var largest = values.DefaultIfEmpty().Max();
                var total = largest == 0 ? 0 : values.Sum(v => v / largest); var angle = -90f; var radius = Math.Min(p.Width, p.Height) / 2;
                var oval = SKRect.Create(p.MidX - radius, p.MidY - radius, radius * 2, radius * 2);
                for (var i = 0; i < values.Length && total > 0; i++)
                {
                    var sweep = (float)(values[i] / largest / total * 360); if (sweep == 0) continue;
                    paint.Color = ChartColor(series[0].PointColors[i]); canvas.DrawArc(oval, angle, sweep, true, paint);
                    var mid = (angle + sweep / 2) * Math.PI / 180; DataLabel(0, i, p.MidX + (float)Math.Cos(mid) * radius * .65f - 40, p.MidY + (float)Math.Sin(mid) * radius * .65f - 7, 80); angle += sweep;
                }
                return;
            }
            if (chart.Kind == WorksheetChartKind.Funnel)
            {
                if (series[0].Values.Any(v => v < 0)) throw new InvalidOperationException("漏斗图数值须非负。");
                var max = series[0].Values.OfType<double>().DefaultIfEmpty().Max(); var height = p.Height / count;
                for (var i = 0; i < count; i++)
                {
                    var width = max == 0 ? 0 : (float)((series[0].Values[i] ?? 0) / max * p.Width);
                    Fill(SKRect.Create(p.MidX - width / 2, p.Top + i * height + 1, width, Math.Max(0, height - 2)), ChartColor(series[0].PointColors[i]));
                    Label(series[0].Categories[i], b.Left + (chart.Legend == "l" ? lw : 0), p.Top + (i + .5f) * height - 7, 58, align: 3);
                    DataLabel(0, i, p.Left, p.Top + (i + .5f) * height - 7, p.Width);
                }
                return;
            }
            var waterfall = chart.Kind == WorksheetChartKind.Waterfall ? ChartData.Waterfall(series[0]) : [];
            var all = waterfall.Length > 0 ? waterfall.SelectMany(v => new[] { v.Start, v.End }).ToArray() : series.SelectMany(s => s.Values).OfType<double>().ToArray();
            static (double Min, double Max) Limits(double[] values, double? requestedMin, double? requestedMax, bool zero)
            {
                var min = requestedMin ?? (zero ? Math.Min(0, values.DefaultIfEmpty().Min()) : values.DefaultIfEmpty().Min());
                var max = requestedMax ?? (zero ? Math.Max(0, values.DefaultIfEmpty().Max()) : values.DefaultIfEmpty().Max());
                if (min >= max) { if (requestedMax is null) max = min + Math.Max(1, Math.Abs(min) * .1); else min = max - Math.Max(1, Math.Abs(max) * .1); }
                if (!double.IsFinite(min) || !double.IsFinite(max) || min >= max) throw new InvalidOperationException("坐标轴超出数值范围，请设置上下限。");
                return (min, max);
            }
            var (minimum, maximum) = Limits(all, chart.Minimum, chart.Maximum, true);
            static double Fraction(double v, double min, double max) { var magnitude = Math.Max(Math.Abs(min), Math.Abs(max)); return (v / magnitude - min / magnitude) / (max / magnitude - min / magnitude); }
            float Y(double v) => (float)(p.Bottom - Fraction(v, minimum, maximum) * p.Height);
            var baseline = Y(Math.Clamp(0, minimum, maximum));
            var yTicks = Math.Clamp((int)(p.Height / Math.Max(24, chart.Appearance.Axis.Size * 2.5)), 1, 5);
            for (var i = 0; i <= yTicks; i++)
            {
                var value = minimum * (1 - i / (double)yTicks) + maximum * (i / (double)yTicks); var y = Y(value);
                Line(p.Left, y, p.Right, y, ChartColor(chart.Appearance.GridColor), .5f);
                Label(Number(value), b.Left + (chart.Legend == "l" ? lw : 0), y - 7, 56, align: 3);
            }
            Line(p.Left, p.Top, p.Left, p.Bottom, SKColors.Gray); Line(p.Left, baseline, p.Right, baseline, SKColors.Gray);
            var step = p.Width / count; var labelStep = Math.Max(1, (int)Math.Ceiling(60 / step));
            var (xmin, xmax) = Limits(series.SelectMany(s => s.XValues).OfType<double>().ToArray(), chart.XMinimum, chart.XMaximum, false);
            float X(double v) => (float)(p.Left + Fraction(v, xmin, xmax) * p.Width);
            if (chart.Kind == WorksheetChartKind.Scatter)
                for (var i = 0; i <= 5; i++)
                { var value = xmin * (1 - i / 5d) + xmax * (i / 5d); var text = CellDisplayFormatter.Format(CellValue.FromNumber(value), new() { NumberFormatCode = chart.XNumberFormat }, book.DateSystem); Label(text, X(value) - 30, p.Bottom + 6, 60, align: 2); }
            else for (var i = 0; i < series[0].Categories.Length; i += labelStep) Label(series[0].Categories[i], p.Left + i * step, p.Bottom + 6, step * labelStep, align: 2);
            canvas.Save(); canvas.ClipRect(p);
            try
            {
                for (var s = 0; s < series.Count; s++)
                {
                    var options = chart.Series[s]; var data = series[s]; var segment = new List<SKPoint>();
                    void Flush()
                    {
                        if (segment.Count > 1 && options.ConnectPoints)
                        {
                            using var path = new SKPathBuilder(); path.MoveTo(segment[0]);
                            for (var j = 1; j < segment.Count; j++)
                            {
                                if (!options.Smooth && chart.Kind != WorksheetChartKind.SmoothLine) { path.LineTo(segment[j]); continue; }
                                var a = segment[Math.Max(0, j - 2)]; var from = segment[j - 1]; var to = segment[j]; var d = segment[Math.Min(segment.Count - 1, j + 1)];
                                path.CubicTo(from.X + (to.X - a.X) / 6, from.Y + (to.Y - a.Y) / 6, to.X - (d.X - from.X) / 6, to.Y - (d.Y - from.Y) / 6, to.X, to.Y);
                            }
                            paint.Style = SKPaintStyle.Stroke; paint.Color = ChartColor(data.Color); paint.StrokeWidth = (float)options.LineWidth; if (paint.StrokeWidth > 0) { using var built = path.Detach(); canvas.DrawPath(built, paint); } paint.Style = SKPaintStyle.Fill;
                        }
                        segment.Clear();
                    }
                    for (var i = 0; i < data.Values.Length; i++)
                    {
                        if (data.Values[i] is not { } value || chart.Kind == WorksheetChartKind.Scatter && data.XValues[i] is null) { Flush(); continue; }
                        var x = chart.Kind == WorksheetChartKind.Scatter ? X(data.XValues[i]!.Value) : p.Left + (i + .5f) * step; var y = Y(value); var color = ChartColor(data.PointColors[i]);
                        if (chart.Kind == WorksheetChartKind.Column)
                        {
                            var overlap = chart.Overlap / 100f; var width = step / (series.Count - (series.Count - 1) * overlap + chart.GapWidth / 100f);
                            var group = width * (series.Count - (series.Count - 1) * overlap);
                            Fill(SKRect.Create(x - group / 2 + s * width * (1 - overlap), Math.Min(y, baseline), width, Math.Abs(baseline - y)), color);
                        }
                        else if (chart.Kind == WorksheetChartKind.Waterfall)
                        {
                            var pair = waterfall[i]; var width = step / (1 + chart.GapWidth / 100f); var start = Y(pair.Start); y = Y(pair.End);
                            if (options.PointColors.IsEmpty || data.PointColors[i].Length == 0) color = ChartColor(data.Totals[i] ? chart.TotalColor : value >= 0 ? chart.IncreaseColor : chart.DecreaseColor);
                            Fill(SKRect.Create(x - width / 2, Math.Min(start, y), width, Math.Abs(start - y)), color);
                            if (chart.ShowConnectors && i + 1 < count) Line(x + width / 2, y, x + step - width / 2, y, SKColors.Gray);
                        }
                        else { segment.Add(new(x, y)); Marker(new(x, y), options, color); }
                        DataLabel(s, i, x - step / 2, Math.Clamp(y - 18, p.Top, p.Bottom - 16), step);
                    }
                    Flush();
                }
            }
            finally { canvas.Restore(); }
        }
        finally { canvas.Restore(); }
    }
}
