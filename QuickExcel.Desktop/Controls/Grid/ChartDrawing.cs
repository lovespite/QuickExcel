using System.Globalization;
using System.Windows;
using System.Windows.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>The same vector renderer is used by the grid, print preview, and printer.</summary>
internal static class ChartDrawing
{
    internal static void Paint(DrawingContext dc, Rect bounds, WorkbookSnapshot book, WorksheetChart chart, bool failIfStale = false, bool monochrome = false)
    {
        Brush Color(string value)
        {
            if (!monochrome) return FastGridArtwork.Brush(value, Brushes.SteelBlue);
            var gray = (byte)((Convert.ToByte(value.Substring(2, 2), 16) * 299 + Convert.ToByte(value.Substring(4, 2), 16) * 587 + Convert.ToByte(value.Substring(6, 2), 16) * 114) / 1000);
            return new SolidColorBrush(System.Windows.Media.Color.FromRgb(gray, gray, gray));
        }
        dc.PushClip(new RectangleGeometry(bounds)); dc.DrawRectangle(Brushes.White, new Pen(Brushes.LightGray, 1), bounds);
        void Label(string text, double x, double y, double width, double size = 11, TextAlignment align = TextAlignment.Left)
        {
            if (width <= 0) return;
            var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), size, Brushes.Black, 1)
            { MaxTextWidth = width, MaxLineCount = 1, Trimming = TextTrimming.CharacterEllipsis, TextAlignment = align };
            dc.DrawText(formatted, new Point(x, y));
        }
        try
        {
            var series = ChartData.Read(book, chart); var count = series.Max(s => s.Values.Length);
            Label(chart.Title, bounds.Left + 12, bounds.Top + 10, bounds.Width - 24, 16, TextAlignment.Center);
            var legendWidth = chart.Legend is "r" or "l" ? Math.Min(130, bounds.Width * .25) : 0;
            var legendHeight = chart.Legend is "t" or "b" ? 32 : 0;
            var plot = new Rect(bounds.Left + 54 + (chart.Legend == "l" ? legendWidth : 0), bounds.Top + 42 + (chart.Legend == "t" ? legendHeight : 0),
                Math.Max(1, bounds.Width - 74 - legendWidth), Math.Max(1, bounds.Height - 78 - legendHeight));
            var legendLabels = chart.Kind == WorksheetChartKind.Pie ? series[0].Categories : series.Select(s => s.Name).ToArray();
            var legendColors = chart.Kind == WorksheetChartKind.Pie ? series[0].PointColors : series.Select(s => s.Color).ToArray();
            if (chart.Legend is not null)
            {
                var horizontal = legendHeight > 0; var itemWidth = horizontal ? (bounds.Width - 20) / Math.Max(1, legendLabels.Length) : Math.Max(20, legendWidth - 12);
                for (var i = 0; i < legendLabels.Length; i++)
                {
                    var x = horizontal ? bounds.Left + 10 + i * itemWidth : chart.Legend == "l" ? bounds.Left + 8 : bounds.Right - legendWidth + 6;
                    var y = horizontal ? chart.Legend == "t" ? bounds.Top + 35 : bounds.Bottom - 23 : plot.Top + i * 20;
                    if (y + 16 > bounds.Bottom - 4) break;
                    dc.DrawRectangle(Color(legendColors[i]), null, new Rect(x, y + 3, 9, 9)); Label(legendLabels[i], x + 14, y, itemWidth - 14);
                }
            }
            if (count == 0) { Label("没有可显示的数据", plot.Left, plot.Top, plot.Width); return; }
            if (chart.Kind == WorksheetChartKind.Pie)
            {
                var values = series[0].Values.Select(v => Math.Abs(v ?? 0)).ToArray(); var largest = values.DefaultIfEmpty().Max();
                var total = largest == 0 ? 0 : values.Sum(v => v / largest); var angle = -Math.PI / 2;
                var radius = Math.Min(plot.Width, plot.Height) / 2; var center = new Point(plot.Left + plot.Width / 2, plot.Top + plot.Height / 2);
                for (var i = 0; i < values.Length && total > 0; i++)
                {
                    var sweep = values[i] / largest / total * Math.PI * 2; if (sweep == 0) continue;
                    var brush = Color(series[0].PointColors[i]);
                    if (sweep >= Math.PI * 2 - 1e-10) dc.DrawEllipse(brush, new Pen(Brushes.White, 1), center, radius, radius);
                    else
                    {
                        var path = new StreamGeometry();
                        using (var g = path.Open())
                        {
                            g.BeginFigure(center, true, true); g.LineTo(new(center.X + radius * Math.Cos(angle), center.Y + radius * Math.Sin(angle)), true, false);
                            g.ArcTo(new(center.X + radius * Math.Cos(angle + sweep), center.Y + radius * Math.Sin(angle + sweep)), new(radius, radius), 0, sweep > Math.PI, SweepDirection.Clockwise, true, false);
                        }
                        dc.DrawGeometry(brush, new Pen(Brushes.White, 1), path);
                    }
                    angle += sweep;
                }
                return;
            }
            var all = series.SelectMany(s => s.Values).OfType<double>().ToArray(); var minimum = chart.Minimum ?? Math.Min(0, all.DefaultIfEmpty().Min()); var maximum = chart.Maximum ?? Math.Max(0, all.DefaultIfEmpty().Max());
            if (minimum >= maximum)
            {
                if (chart.Maximum is null) maximum = minimum + Math.Max(1, Math.Abs(minimum) * .1);
                else minimum = maximum - Math.Max(1, Math.Abs(maximum) * .1);
            }
            if (!double.IsFinite(minimum) || !double.IsFinite(maximum) || minimum >= maximum) throw new InvalidOperationException("自动坐标轴超出数值范围，请明确设置可显示的上下限。");
            var magnitude = Math.Max(Math.Abs(minimum), Math.Abs(maximum)); var range = maximum / magnitude - minimum / magnitude;
            double Y(double v) => plot.Bottom - (v / magnitude - minimum / magnitude) / range * plot.Height;
            var baseline = Y(Math.Clamp(0, minimum, maximum));
            for (var i = 0; i <= 5; i++)
            {
                var value = minimum * (1 - i / 5d) + maximum * (i / 5d); var y = Y(value);
                dc.DrawLine(new Pen(Brushes.LightGray, .5), new(plot.Left, y), new(plot.Right, y));
                var label = CellDisplayFormatter.Format(CellValue.FromNumber(value), new StyleDefinition { NumberFormatCode = chart.NumberFormat }, book.DateSystem);
                Label(label, bounds.Left + (chart.Legend == "l" ? legendWidth : 0), y - 7, 48, 10, TextAlignment.Right);
            }
            dc.DrawLine(new Pen(Brushes.Gray, 1), new(plot.Left, plot.Top), new(plot.Left, plot.Bottom)); dc.DrawLine(new Pen(Brushes.Gray, 1), new(plot.Left, baseline), new(plot.Right, baseline));
            var step = plot.Width / count; var labelStep = Math.Max(1, (int)Math.Ceiling(60 / step));
            for (var i = 0; i < series[0].Categories.Length; i += labelStep) Label(series[0].Categories[i], plot.Left + i * step, plot.Bottom + 6, step * labelStep, 10, TextAlignment.Center);
            dc.PushClip(new RectangleGeometry(plot));
            for (var s = 0; s < series.Count; s++)
            {
                var values = series[s].Values; var brush = Color(series[s].Color); Point? previous = null;
                for (var i = 0; i < values.Length; i++)
                {
                    if (values[i] is not { } value) { previous = null; continue; }
                    var x = plot.Left + (i + .5) * step; var y = Y(value);
                    if (chart.Kind == WorksheetChartKind.Column)
                    {
                        var barWidth = step / (series.Count + 1.5); var left = x - series.Count * barWidth / 2 + s * barWidth;
                        dc.DrawRectangle(brush, null, new Rect(left, Math.Min(y, baseline), barWidth, Math.Abs(baseline - y)));
                    }
                    else { if (previous is { } p) dc.DrawLine(new Pen(brush, 2), p, new(x, y)); previous = new(x, y); }
                }
            }
            dc.Pop();
        }
        catch (InvalidOperationException ex) when (!failIfStale) { Label(ex.Message, bounds.Left + 12, bounds.Top + 45, bounds.Width - 24); }
        finally { dc.Pop(); }
    }
}
