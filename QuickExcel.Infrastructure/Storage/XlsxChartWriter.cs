using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxDrawingAnchors;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    // Replace only modeled fields. Existing DrawingML, fonts, style references and extensions survive data edits.
    private static void Put(XElement parent, XElement child, string order)
    {
        if (parent.Element(child.Name) is { } old) { old.ReplaceWith(child); return; }
        var names = order.Split(' '); var index = Array.IndexOf(names, child.Name.LocalName);
        var next = parent.Elements().FirstOrDefault(e => Array.IndexOf(names, e.Name.LocalName) > index);
        if (next is null) parent.Add(child); else next.AddBeforeSelf(child);
    }
    private static XElement Solid(string argb) => new(Art + "solidFill", new XElement(Art + "srgbClr", new XAttribute("val", argb[2..]),
        argb.StartsWith("FF", StringComparison.OrdinalIgnoreCase) ? null : new XElement(Art + "alpha", new XAttribute("val", int.Parse(argb[..2], NumberStyles.HexNumber) * 100000 / 255))));
    private static XElement Shape(XNamespace ns, string fill, string? stroke = null, double width = 1) => new(ns + "spPr", Solid(fill),
        stroke is null ? null : new XElement(Art + "ln", new XAttribute("w", (int)Math.Round(width * 9525)), Solid(stroke)));
    private static XElement TextProperties(XNamespace ns, ChartTextStyle text) => new(ns + "txPr", new XElement(Art + "bodyPr"), new XElement(Art + "lstStyle"),
        new XElement(Art + "p", new XElement(Art + "pPr", new XElement(Art + "defRPr", new XAttribute("sz", (int)Math.Round(text.Size * 100)), new XAttribute("b", text.Bold ? 1 : 0),
            new XAttribute("i", text.Italic ? 1 : 0), Solid(text.Color), text.FontFamily is { } font ? new XElement(Art + "latin", new XAttribute("typeface", font)) : null))));
    private static XElement RichTitle(XNamespace ns, string text) => new(ns + "tx", new XElement(ns + "rich", new XElement(Art + "bodyPr"), new XElement(Art + "lstStyle"), new XElement(Art + "p", new XElement(Art + "r", new XElement(Art + "t", text)))));
    private static string TitleText(XElement? title) => string.Concat(title?.Descendants(Art + "t").Select(t => t.Value) ?? []);
    private static XElement Cache(string formula, IEnumerable<string?> values, bool numeric, string format)
    {
        var array = values.ToArray();
        return new XElement(Chart + (numeric ? "numRef" : "strRef"), new XElement(Chart + "f", formula.TrimStart('=')),
            new XElement(Chart + (numeric ? "numCache" : "strCache"), numeric ? new XElement(Chart + "formatCode", format) : null, Val("ptCount", array.Length),
                array.Select((v, i) => v is null ? null : new XElement(Chart + "pt", new XAttribute("idx", i), new XElement(Chart + "v", v)))));
    }
    private static IEnumerable<string?> Numbers(IEnumerable<double?> values) => values.Select(v => v?.ToString("R", CultureInfo.InvariantCulture));

    private static XElement StandardXml(WorkbookSnapshot book, WorksheetChart model)
    {
        var data = ChartData.Read(book, model with { VisibleOnly = false }, includeFiltered: true);
        var graphName = model.Kind switch { WorksheetChartKind.Column => "barChart", WorksheetChartKind.Line or WorksheetChartKind.SmoothLine => "lineChart", WorksheetChartKind.Scatter => "scatterChart", _ => "pieChart" };
        var prior = model.OriginalXml.IsEmpty ? null : Xml(model.OriginalXml.ToArray(), new(CancellationToken.None)).Root;
        if (prior?.Name == Chart + "chartSpace") NormalizeStandardFilters(prior);
        var root = prior?.Name == Chart + "chartSpace" ? prior : new XElement(Chart + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", Chart), new XAttribute(XNamespace.Xmlns + "a", Art),
            Val("date1904", book.DateSystem == ExcelDateSystem.Mac1904 ? 1 : 0), new XElement(Chart + "chart", new XElement(Chart + "plotArea", new XElement(Chart + "layout"))));
        if (root.Element(Chart + "roundedCorners") is null) Put(root, Val("roundedCorners", 0), "date1904 lang roundedCorners style chart spPr txPr printSettings extLst");
        var chart = root.Element(Chart + "chart")!; var plot = chart.Element(Chart + "plotArea")!;
        var title = chart.Element(Chart + "title");
        if (TitleText(title) != model.Title)
        {
            if (model.Title.Length == 0) title?.Remove();
            else { title ??= new XElement(Chart + "title", Val("overlay", 0)); Put(title, RichTitle(Chart, model.Title), "tx layout overlay spPr txPr extLst"); if (title.Parent is null) chart.AddFirst(title); }
        }
        if (title is not null && title.Element(Chart + "txPr") is null) title.Add(TextProperties(Chart, model.Appearance.Title));
        var graph = plot.Element(Chart + graphName);
        if (graph is null)
        {
            plot.Elements().Where(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal) || e.Name.LocalName is "catAx" or "valAx").Remove();
            graph = new XElement(Chart + graphName); Put(plot, graph, "layout barChart lineChart pieChart scatterChart catAx valAx spPr extLst");
        }
        var graphOrder = "barDir grouping scatterStyle varyColors ser dLbls gapWidth overlap marker smooth firstSliceAng axId extLst";
        if (model.Kind == WorksheetChartKind.Column) { Put(graph, Val("barDir", "col"), graphOrder); Put(graph, Val("grouping", "clustered"), graphOrder); }
        if (model.Kind is WorksheetChartKind.Line or WorksheetChartKind.SmoothLine) Put(graph, Val("grouping", "standard"), graphOrder);
        if (model.Kind == WorksheetChartKind.Scatter) Put(graph, Val("scatterStyle", model.Series.Any(s => s.ConnectPoints) ? model.Series.Any(s => s.Smooth) ? "smoothMarker" : "lineMarker" : "marker"), graphOrder);
        Put(graph, Val("varyColors", model.Kind == WorksheetChartKind.Pie ? 1 : 0), graphOrder);
        var oldSeries = graph.Elements(Chart + "ser").ToArray(); graph.Elements(Chart + "ser").Remove();
        var seriesOrder = "idx order tx spPr invertIfNegative marker dPt dLbls cat val xVal yVal smooth extLst";
        for (var i = 0; i < model.Series.Length; i++)
        {
            var s = model.Series[i]; var values = data[i]; var ser = i < oldSeries.Length ? oldSeries[i] : new XElement(Chart + "ser");
            ser.Elements(Chart + "dPt").Where(p => !int.TryParse(Value(p, "idx"), out var index) || index >= values.Values.Length).Remove();
            Put(ser, Val("idx", i), seriesOrder); Put(ser, Val("order", i), seriesOrder);
            Put(ser, new XElement(Chart + "tx", s.NameReference is { } name ? Cache(name, [values.Name], false, "General") : new XElement(Chart + "v", s.Name)), seriesOrder);
            if (model.Kind == WorksheetChartKind.Column) Put(ser, Val("invertIfNegative", 0), seriesOrder);
            if (ser.Element(Chart + "spPr") is null) Put(ser, model.Kind is WorksheetChartKind.Line or WorksheetChartKind.SmoothLine or WorksheetChartKind.Scatter ?
                new XElement(Chart + "spPr", new XElement(Art + "ln", new XAttribute("w", (int)Math.Round(s.LineWidth * 9525)), s.ConnectPoints ? Solid(s.Color) : new XElement(Art + "noFill"))) : Shape(Chart, s.Color), seriesOrder);
            if (model.Kind is WorksheetChartKind.Line or WorksheetChartKind.SmoothLine or WorksheetChartKind.Scatter)
            {
                var marker = ser.Element(Chart + "marker") ?? new XElement(Chart + "marker"); Put(marker, Val("symbol", s.Marker), "symbol size spPr extLst"); Put(marker, Val("size", s.MarkerSize), "symbol size spPr extLst"); Put(ser, marker, seriesOrder);
                if (!s.ConnectPoints && marker.Element(Chart + "spPr") is null) marker.Add(Shape(Chart, s.Color));
                Put(ser, Val("smooth", s.Smooth || model.Kind == WorksheetChartKind.SmoothLine ? 1 : 0), seriesOrder);
            }
            for (var point = 0; point < s.PointColors.Length; point++)
            {
                if (ser.Elements(Chart + "dPt").Any(p => Value(p, "idx") == point.ToString(CultureInfo.InvariantCulture))) continue;
                var node = new XElement(Chart + "dPt", Val("idx", point), Shape(Chart, s.PointColors[point]));
                var after = ser.Elements().LastOrDefault(e => e.Name.LocalName is "idx" or "order" or "tx" or "spPr" or "invertIfNegative" or "marker" or "dPt");
                if (after is null) ser.AddFirst(node); else after.AddAfterSelf(node);
            }
            if (model.Kind == WorksheetChartKind.Column)
                foreach (var point in ser.Elements(Chart + "dPt")) Put(point, Val("invertIfNegative", 0), "idx invertIfNegative marker bubble3D explosion spPr pictureOptions extLst");
            var labels = ser.Element(Chart + "dLbls");
            if (labels is not null || s.ShowValues || s.ShowCategories || s.ShowSeriesName)
            {
                labels ??= new XElement(Chart + "dLbls");
                const string labelOrder = "numFmt spPr txPr dLblPos showLegendKey showVal showCatName showSerName showPercent showBubbleSize separator showLeaderLines extLst";
                Put(labels, Val("showVal", s.ShowValues ? 1 : 0), labelOrder); Put(labels, Val("showCatName", s.ShowCategories ? 1 : 0), labelOrder); Put(labels, Val("showSerName", s.ShowSeriesName ? 1 : 0), labelOrder); Put(ser, labels, seriesOrder);
            }
            var numeric = s.NumericCategories || model.Kind == WorksheetChartKind.Scatter;
            Put(ser, new XElement(Chart + (model.Kind == WorksheetChartKind.Scatter ? "xVal" : "cat"), Cache(s.Categories, numeric ? Numbers(values.XValues) : values.Categories, numeric, model.XNumberFormat)), seriesOrder);
            Put(ser, new XElement(Chart + (model.Kind == WorksheetChartKind.Scatter ? "yVal" : "val"), Cache(s.Values, Numbers(values.Values), true, model.NumberFormat)), seriesOrder);
            var before = graph.Elements().FirstOrDefault(e => Array.IndexOf(graphOrder.Split(' '), e.Name.LocalName) > 4);
            if (before is null) graph.Add(ser); else before.AddBeforeSelf(ser);
        }
        if (model.Kind == WorksheetChartKind.Column) { Put(graph, Val("gapWidth", model.GapWidth), graphOrder); Put(graph, Val("overlap", model.Overlap), graphOrder); }
        if (model.Kind != WorksheetChartKind.Pie)
        {
            var axes = plot.Elements().Where(e => e.Name.LocalName is "catAx" or "valAx").ToArray();
            if (axes.Length == 0)
            {
                XElement Axis(bool y) => new(Chart + (y || model.Kind == WorksheetChartKind.Scatter ? "valAx" : "catAx"), Val("axId", y ? 200 : 100), new XElement(Chart + "scaling", Val("orientation", "minMax")), Val("delete", 0), Val("axPos", y ? "l" : "b"),
                    y ? new XElement(Chart + "majorGridlines", Shape(Chart, "00FFFFFF", model.Appearance.GridColor)) : null, new XElement(Chart + "numFmt", new XAttribute("formatCode", y ? model.NumberFormat : model.XNumberFormat), new XAttribute("sourceLinked", 0)),
                    Val("majorTickMark", "none"), Val("minorTickMark", "none"), Val("tickLblPos", "nextTo"), TextProperties(Chart, model.Appearance.Axis), Val("crossAx", y ? 100 : 200), Val("crosses", "autoZero"), y || model.Kind == WorksheetChartKind.Scatter ? Val("crossBetween", "between") : Val("auto", 1));
                var x = Axis(false); var y = Axis(true); var before = plot.Element(Chart + "spPr"); if (before is null) plot.Add(x, y); else before.AddBeforeSelf(x, y);
                graph.Elements(Chart + "axId").Remove(); graph.Add(Val("axId", 100), Val("axId", 200)); axes = [x, y];
            }
            foreach (var axis in axes)
            {
                var y = Value(axis, "axPos") == "l"; var scaling = axis.Element(Chart + "scaling")!;
                foreach (var name in new[] { "min", "max" }) scaling.Element(Chart + name)?.Remove();
                var min = y ? model.Minimum : model.XMinimum; var max = y ? model.Maximum : model.XMaximum;
                if (max.HasValue) scaling.Add(Val("max", max.Value)); if (min.HasValue) scaling.Add(Val("min", min.Value));
                if (y || model.Kind == WorksheetChartKind.Scatter) axis.Element(Chart + "numFmt")?.SetAttributeValue("formatCode", y ? model.NumberFormat : model.XNumberFormat);
            }
        }
        var legend = chart.Element(Chart + "legend");
        if (model.Legend is null) legend?.Remove();
        else { legend ??= new XElement(Chart + "legend", Val("overlay", 0), TextProperties(Chart, model.Appearance.Legend)); Put(legend, Val("legendPos", model.Legend), "legendPos legendEntry layout overlay spPr txPr extLst"); Put(chart, legend, "title autoTitleDeleted plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst"); }
        Put(chart, Val("plotVisOnly", model.VisibleOnly ? 1 : 0), "title autoTitleDeleted plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst");
        Put(chart, Val("dispBlanksAs", "gap"), "title autoTitleDeleted plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst");
        if (root.Element(Chart + "spPr") is null) Put(root, Shape(Chart, model.Appearance.Background, model.Appearance.BorderColor, model.Appearance.BorderWidth), "date1904 lang roundedCorners style chart spPr txPr printSettings extLst");
        if (plot.Element(Chart + "spPr") is null) plot.Add(Shape(Chart, model.Appearance.PlotBackground));
        WriteStandardFilters(graph, model);
        return root;
    }
}
