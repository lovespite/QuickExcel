using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxDrawingAnchors;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    internal static readonly XNamespace ChartEx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    private const string ChartExType = "application/vnd.ms-office.chartex+xml";
    private const string ChartExRelationship = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
    private static XElement ExVal(string name, object value) => new(ChartEx + name, new XAttribute("val", value));

    private static void ValidateExtended(XElement root)
    {
        // ChartEx has its own data/dimension and axis schema. Do not accept other statistical layouts as ordinary bars.
        var containers = new Dictionary<string, string>
        {
            ["chartSpace"] = "chartData chart spPr txPr printSettings extLst", ["extLst"] = "ext", ["chartData"] = "data", ["data"] = "numDim strDim",
            ["numDim"] = "f lvl", ["strDim"] = "f lvl", ["lvl"] = "pt", ["f"] = "", ["pt"] = "", ["v"] = "",
            ["chart"] = "title plotArea legend", ["title"] = "tx spPr txPr", ["tx"] = "txData rich", ["txData"] = "f v",
            ["plotArea"] = "plotAreaRegion axis spPr", ["plotAreaRegion"] = "plotSurface series", ["plotSurface"] = "spPr",
            ["series"] = "tx spPr dataPt dataLabels dataId layoutPr axisId", ["dataPt"] = "spPr", ["dataId"] = "", ["axisId"] = "",
            ["layoutPr"] = "visibility subtotals", ["subtotals"] = "idx", ["idx"] = "", ["visibility"] = "",
            ["dataLabels"] = "numFmt spPr txPr visibility separator", ["separator"] = "", ["numFmt"] = "",
            ["axis"] = "catScaling valScaling majorGridlines minorGridlines majorTickMarks minorTickMarks tickLabels numFmt spPr txPr",
            ["catScaling"] = "", ["valScaling"] = "", ["majorGridlines"] = "spPr", ["minorGridlines"] = "spPr",
            ["majorTickMarks"] = "", ["minorTickMarks"] = "", ["tickLabels"] = "", ["legend"] = "spPr txPr",
            ["printSettings"] = "headerFooter pageMargins pageSetup", ["headerFooter"] = "", ["pageMargins"] = "", ["pageSetup"] = ""
        };
        foreach (var e in root.DescendantsAndSelf())
        {
            if (e.Ancestors().Any(IsExtendedFilter)) continue;
            if (e.Name == ChartEx + "ext") { if (!IsExtendedFilter(e) || e.Parent?.Parent != root) throw new NotSupportedException("未知 ChartEx 扩展。"); continue; }
            if (e.Name.Namespace == Art) { ValidateDrawingElement(e); continue; }
            if (e.Name.Namespace != ChartEx) throw new NotSupportedException("未知 ChartEx 命名空间");
            if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.Namespace != XNamespace.None)) throw new NotSupportedException("未知 ChartEx 引用");
            var name = e.Name.LocalName;
            if (name is "rich" or "spPr" or "txPr")
            { if (e.Elements().Any(c => c.Name.Namespace != Art)) throw new NotSupportedException("未知 ChartEx 样式"); continue; }
            if (!containers.TryGetValue(name, out var children) || e.Elements().Any(c => c.Name.Namespace != ChartEx || !children.Split(' ').Contains(c.Name.LocalName))) throw new NotSupportedException("未建模 ChartEx 结构 " + name);
            var attributes = name switch
            {
                "chartSpace" => "version", "data" => "id", "numDim" or "strDim" => "type", "lvl" => "ptCount formatCode name", "pt" => "idx", "f" => "dir",
                "series" => "layoutId hidden uniqueId formatIdx", "dataPt" => "idx", "dataId" or "idx" => "val", "axis" => "id hidden",
                "title" or "legend" => "pos align overlay", "catScaling" => "gapWidth", "valScaling" => "min max majorUnit minorUnit",
                "numFmt" => "formatCode sourceLinked", "visibility" => "connectorLines seriesName categoryName value", "dataLabels" => "pos",
                "majorTickMarks" or "minorTickMarks" => "type", "pageMargins" => "b l r t header footer", "pageSetup" => "paperSize orientation", _ => ""
            };
            if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && !attributes.Split(' ').Contains(a.Name.LocalName))) throw new NotSupportedException("未建模 ChartEx 属性 " + name);
            if (name != "series" && (string?)e.Attribute("hidden") is "1" or "true" || (string?)e.Attribute("overlay") is "1" or "true") throw new NotSupportedException("隐藏或叠加 ChartEx 元素");
            if (name == "valScaling" && e.Attributes().Any(a => a.Name.LocalName is "majorUnit" or "minorUnit" && a.Value != "auto")) throw new NotSupportedException("ChartEx 自定义刻度");
        }
    }

    private static WorksheetChart ReadExtended(XElement root, ChartStyleReader styles)
    {
        if (root.Name != ChartEx + "chartSpace") throw new InvalidDataException("ChartEx 根元素无效。");
        ValidateExtended(root);
        var filter = NormalizeExtendedFilters(root);
        var chart = root.Element(ChartEx + "chart")!; var plot = chart.Element(ChartEx + "plotArea")!;
        var seriesNodes = plot.Element(ChartEx + "plotAreaRegion")?.Elements(ChartEx + "series").ToArray() ?? [];
        if (seriesNodes.Length != 1) throw new NotSupportedException("瀑布图和漏斗图仅支持一个系列");
        var ser = seriesNodes[0]; var kind = (string?)ser.Attribute("layoutId") switch { "waterfall" => WorksheetChartKind.Waterfall, "funnel" => WorksheetChartKind.Funnel, _ => throw new NotSupportedException("此 ChartEx 图表类型") };
        var id = (string?)ser.Element(ChartEx + "dataId")?.Attribute("val");
        var dataNodes = root.Element(ChartEx + "chartData")?.Elements(ChartEx + "data").ToArray() ?? [];
        if (dataNodes.Length != 1 || (string?)dataNodes[0].Attribute("id") != id) throw new NotSupportedException("ChartEx 数据绑定不唯一");
        var data = dataNodes[0];
        var category = data.Elements(ChartEx + "strDim").SingleOrDefault(e => (string?)e.Attribute("type") == "cat");
        var value = data.Elements(ChartEx + "numDim").SingleOrDefault(e => (string?)e.Attribute("type") == "val");
        if (data.Elements().Count() != 2 || category is null || value is null || data.Elements().Any(d => d.Elements(ChartEx + "lvl").Count() > 1)) throw new NotSupportedException("ChartEx 多层分类或非标准数据维度");
        string Reference(XElement node) => "=" + (node.Element(ChartEx + "f")?.Value.TrimStart('=') ?? throw new NotSupportedException("ChartEx 数据必须直接引用单元格"));
        var values = Reference(value); var count = WorksheetChart.Count(WorksheetChart.ParseReference(values));
        var tx = ser.Element(ChartEx + "tx")?.Element(ChartEx + "txData"); var nameRef = tx?.Element(ChartEx + "f")?.Value;
        var color = styles.SeriesColor(ser.Element(ChartEx + "spPr"), 0, false);
        var points = kind == WorksheetChartKind.Funnel || ser.Elements(ChartEx + "dataPt").Any() ? Enumerable.Range(0, count).Select(i => kind == WorksheetChartKind.Funnel ? styles.Palette(i) : "").ToArray() : [];
        foreach (var p in ser.Elements(ChartEx + "dataPt"))
        {
            var index = TypedStyles.Int((string?)p.Attribute("idx") ?? "", count - 1);
            points[index] = styles.Fill(p.Element(ChartEx + "spPr"), points[index]);
        }
        var visibility = ser.Element(ChartEx + "dataLabels")?.Element(ChartEx + "visibility");
        var layout = ser.Element(ChartEx + "layoutPr"); var axis = plot.Elements(ChartEx + "axis").FirstOrDefault(a => a.Element(ChartEx + "valScaling") is not null);
        var categoryAxis = plot.Elements(ChartEx + "axis").FirstOrDefault(a => a.Element(ChartEx + "catScaling") is not null);
        var scaling = axis?.Element(ChartEx + "valScaling");
        double? Limit(string name) => (string?)scaling?.Attribute(name) is { } text && text != "auto" ? double.Parse(text, CultureInfo.InvariantCulture) : null;
        if (chart.Element(ChartEx + "title")?.Descendants(ChartEx + "f").Any() == true) throw new NotSupportedException("引用公式的图表标题");
        var title = chart.Element(ChartEx + "title");
        return new()
        {
            Kind = kind, Title = title?.Descendants(ChartEx + "v").FirstOrDefault()?.Value ?? TitleText(title), Legend = chart.Element(ChartEx + "legend") is { } legend ? (string?)legend.Attribute("pos") ?? "r" : null,
            Series = [new(nameRef is null ? tx?.Element(ChartEx + "v")?.Value ?? "" : "", nameRef is null ? null : "=" + nameRef.TrimStart('='), Reference(category), values, color)
            {
                IsFiltered = filter.Hidden, HiddenPoints = filter.Points, NumericCategories = category.Name.LocalName == "numDim", PointColors = [.. points],
                TotalPoints = [.. layout?.Element(ChartEx + "subtotals")?.Elements(ChartEx + "idx").Select(e => TypedStyles.Int((string?)e.Attribute("val") ?? "", count - 1)) ?? []],
                ShowValues = (string?)visibility?.Attribute("value") is "1" or "true", ShowCategories = (string?)visibility?.Attribute("categoryName") is "1" or "true", ShowSeriesName = (string?)visibility?.Attribute("seriesName") is "1" or "true"
            }],
            Minimum = Limit("min"), Maximum = Limit("max"), NumberFormat = (string?)axis?.Element(ChartEx + "numFmt")?.Attribute("formatCode") ?? (string?)value.Element(ChartEx + "lvl")?.Attribute("formatCode") ?? "General",
            GapWidth = (string?)categoryAxis?.Element(ChartEx + "catScaling")?.Attribute("gapWidth") is { } gap && gap != "auto" ? checked((int)Math.Round(double.Parse(gap, CultureInfo.InvariantCulture) * 100)) : 150,
            ShowConnectors = (string?)layout?.Element(ChartEx + "visibility")?.Attribute("connectorLines") is not ("0" or "false"),
            IncreaseColor = kind == WorksheetChartKind.Waterfall ? styles.Palette(0) : "FF4472C4", DecreaseColor = kind == WorksheetChartKind.Waterfall ? styles.Palette(1) : "FFED7D31", TotalColor = kind == WorksheetChartKind.Waterfall ? styles.Palette(2) : "FFA5A5A5", Appearance = styles.Appearance(root, chart, plot)
        };
    }

    private static XElement ExtendedXml(WorkbookSnapshot book, WorksheetChart model)
    {
        var data = ChartData.Read(book, model with { VisibleOnly = false }, includeFiltered: true)[0]; var series = model.Series[0];
        var old = model.OriginalXml.IsEmpty ? null : Xml(model.OriginalXml.ToArray(), new(CancellationToken.None)).Root;
        if (old?.Name == ChartEx + "chartSpace") NormalizeExtendedFilters(old);
        var root = old?.Name == ChartEx + "chartSpace" ? old : new XElement(ChartEx + "chartSpace", new XAttribute(XNamespace.Xmlns + "cx", ChartEx), new XAttribute(XNamespace.Xmlns + "a", Art),
            new XElement(ChartEx + "chartData"), new XElement(ChartEx + "chart", new XElement(ChartEx + "plotArea", new XElement(ChartEx + "plotAreaRegion"))));
        var chart = root.Element(ChartEx + "chart")!; var plot = chart.Element(ChartEx + "plotArea")!; var region = plot.Element(ChartEx + "plotAreaRegion")!;
        XElement Dimension(string type, bool numeric, string formula, IEnumerable<string?> values)
        {
            var array = values.ToArray();
            return new XElement(ChartEx + (numeric ? "numDim" : "strDim"), new XAttribute("type", type), new XElement(ChartEx + "f", formula.TrimStart('=')),
                new XElement(ChartEx + "lvl", new XAttribute("ptCount", array.Length), numeric ? new XAttribute("formatCode", model.NumberFormat) : null,
                    array.Select((v, i) => v is null ? null : new XElement(ChartEx + "pt", new XAttribute("idx", i), v))));
        }
        root.Element(ChartEx + "chartData")!.ReplaceNodes(new XElement(ChartEx + "data", new XAttribute("id", 0),
            Dimension("cat", series.NumericCategories, series.Categories, series.NumericCategories ? Numbers(data.XValues) : data.Categories), Dimension("val", true, series.Values, Numbers(data.Values))));
        var ser = region.Element(ChartEx + "series") ?? new XElement(ChartEx + "series");
        ser.Elements(ChartEx + "dataPt").Where(p => !int.TryParse((string?)p.Attribute("idx"), out var index) || index >= data.Values.Length).Remove();
        ser.SetAttributeValue("layoutId", model.Kind == WorksheetChartKind.Waterfall ? "waterfall" : "funnel");
        const string seriesOrder = "tx spPr valueColors valueColorPositions dataPt dataLabels dataId layoutPr axisId extLst";
        Put(ser, new XElement(ChartEx + "tx", new XElement(ChartEx + "txData", series.NameReference is { } name ? new XElement(ChartEx + "f", name.TrimStart('=')) : null, new XElement(ChartEx + "v", data.Name))), seriesOrder);
        if (ser.Element(ChartEx + "spPr") is null) Put(ser, Shape(ChartEx, series.Color), seriesOrder);
        var labels = ser.Element(ChartEx + "dataLabels") ?? new XElement(ChartEx + "dataLabels");
        Put(labels, new XElement(ChartEx + "visibility", new XAttribute("seriesName", series.ShowSeriesName ? 1 : 0), new XAttribute("categoryName", series.ShowCategories ? 1 : 0), new XAttribute("value", series.ShowValues ? 1 : 0)), "numFmt spPr txPr visibility separator dataLabel dataLabelHidden extLst");
        Put(ser, labels, seriesOrder);
        Put(ser, ExVal("dataId", 0), seriesOrder);
        if (model.Kind == WorksheetChartKind.Waterfall) Put(ser, new XElement(ChartEx + "layoutPr", new XElement(ChartEx + "visibility", new XAttribute("connectorLines", model.ShowConnectors ? 1 : 0)),
            series.TotalPoints.IsEmpty ? null : new XElement(ChartEx + "subtotals", series.TotalPoints.Select(i => ExVal("idx", i)))), seriesOrder);
        else ser.Element(ChartEx + "layoutPr")?.Remove();
        for (var i = 0; i < series.PointColors.Length; i++)
        {
            if (series.PointColors[i].Length == 0) continue;
            if (ser.Elements(ChartEx + "dataPt").Any(p => (string?)p.Attribute("idx") == i.ToString(CultureInfo.InvariantCulture))) continue;
            var point = new XElement(ChartEx + "dataPt", new XAttribute("idx", i), Shape(ChartEx, series.PointColors[i]));
            ser.Element(ChartEx + "dataLabels")!.AddBeforeSelf(point);
        }
        if (ser.Parent is null) region.Add(ser);
        // Waterfall/funnel use their layout's implicit axes. Explicit series axisId entries
        // pass the OOXML schema but make Excel reject the entire workbook.
        ser.Elements(ChartEx + "axisId").Remove();
        if (model.Kind == WorksheetChartKind.Waterfall)
        {
            var axes = plot.Elements(ChartEx + "axis").ToArray();
            if (axes.Length == 0) { region.AddAfterSelf(new XElement(ChartEx + "axis", new XAttribute("id", 0), new XElement(ChartEx + "catScaling"), new XElement(ChartEx + "tickLabels")),
                new XElement(ChartEx + "axis", new XAttribute("id", 1), new XElement(ChartEx + "valScaling"), new XElement(ChartEx + "majorGridlines", Shape(ChartEx, "00FFFFFF", model.Appearance.GridColor)), new XElement(ChartEx + "tickLabels"))); axes = plot.Elements(ChartEx + "axis").ToArray(); }
            foreach (var axis in axes)
            {
                if (axis.Element(ChartEx + "catScaling") is { } cat) cat.SetAttributeValue("gapWidth", model.GapWidth / 100d);
                if (axis.Element(ChartEx + "valScaling") is { } val) { val.SetAttributeValue("min", model.Minimum?.ToString("R", CultureInfo.InvariantCulture) ?? "auto"); val.SetAttributeValue("max", model.Maximum?.ToString("R", CultureInfo.InvariantCulture) ?? "auto"); }
                Put(axis, new XElement(ChartEx + "numFmt", new XAttribute("formatCode", model.NumberFormat), new XAttribute("sourceLinked", 0)), "catScaling valScaling title units majorGridlines minorGridlines majorTickMarks minorTickMarks tickLabels numFmt spPr txPr extLst");
                if (axis.Element(ChartEx + "txPr") is null) Put(axis, TextProperties(ChartEx, model.Appearance.Axis), "catScaling valScaling title units majorGridlines minorGridlines majorTickMarks minorTickMarks tickLabels numFmt spPr txPr extLst");
            }
        }
        else plot.Elements(ChartEx + "axis").Remove();
        var title = chart.Element(ChartEx + "title"); var oldTitle = title?.Descendants(ChartEx + "v").FirstOrDefault()?.Value ?? TitleText(title);
        if (oldTitle != model.Title)
        {
            if (model.Title.Length == 0) title?.Remove();
            else { title ??= new XElement(ChartEx + "title", new XAttribute("pos", "t"), new XAttribute("align", "ctr"), new XAttribute("overlay", 0)); Put(title, RichTitle(ChartEx, model.Title), "tx spPr txPr offset extLst"); if (title.Parent is null) chart.AddFirst(title); }
        }
        if (title is not null && title.Element(ChartEx + "txPr") is null) title.Add(TextProperties(ChartEx, model.Appearance.Title));
        if (model.Legend is null) chart.Element(ChartEx + "legend")?.Remove();
        else { var legend = chart.Element(ChartEx + "legend") ?? new XElement(ChartEx + "legend", new XAttribute("align", "ctr"), new XAttribute("overlay", 0)); legend.SetAttributeValue("pos", model.Legend); if (legend.Element(ChartEx + "txPr") is null) legend.Add(TextProperties(ChartEx, model.Appearance.Legend)); if (legend.Parent is null) chart.Add(legend); }
        if (root.Element(ChartEx + "spPr") is null) root.Add(Shape(ChartEx, model.Appearance.Background, model.Appearance.BorderColor, model.Appearance.BorderWidth));
        if (plot.Element(ChartEx + "spPr") is null) plot.Add(Shape(ChartEx, model.Appearance.PlotBackground));
        WriteExtendedFilters(root, model);
        return root;
    }
}
