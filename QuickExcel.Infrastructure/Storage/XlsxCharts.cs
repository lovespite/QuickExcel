using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.XlsxDrawingAnchors;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    internal static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string ChartType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    private static XElement Val(string name, object value) => new(Chart + name, new XAttribute("val", value));
    private static string? Value(XElement? e, string child) => (string?)e?.Element(Chart + child)?.Attribute("val");

    internal static ImmutableArray<WorksheetChart> Read(XlsxPackage package, string workbook, string sheetPart, XElement sheet, WorksheetSnapshot geometry)
    {
        var charts = ImmutableArray.CreateBuilder<WorksheetChart>();
        if (!sheet.Elements(Main + "drawing").Any()) return [];
        var colors = new XlsxColors(package, workbook, new XElement(Main + "styleSheet"), sheetPart, tableDefaults: true);
        foreach (var drawing in sheet.Elements(Main + "drawing"))
        {
            var path = package.Related(sheetPart, Attr(drawing, Rel + "id"), "drawing"); var root = package.XmlPart(path).Root!;
            foreach (var anchor in root.Elements().Where(a => a.Element(Drawing + "graphicFrame") is not null))
            {
                var frame = anchor.Element(Drawing + "graphicFrame")!;
                var graphicData = frame.Element(Art + "graphic")?.Element(Art + "graphicData");
                var reference = graphicData?.Descendants(ChartEx + "chart").FirstOrDefault() ?? graphicData?.Element(Chart + "chart");
                if (reference is null) { package.Loss($"{path}: 非图表绘图尚未支持", "drawing/graphicFrame", path); continue; }
                var extended = reference.Name.Namespace == ChartEx;
                var relation = package.Relationships(path).GetValueOrDefault(Attr(reference, Rel + "id"));
                var chartPath = extended ? package.RelatedUri(path, Attr(reference, Rel + "id"), ChartExRelationship) : package.Related(path, Attr(reference, Rel + "id"), "chart");
                if (relation is null || relation.External || relation.Target != chartPath) throw new InvalidDataException("图表关系不匹配。");
                package.RequireType(chartPath, extended ? ChartExType : ChartType); var xml = package.XmlPart(chartPath).Root!; var raw = package.BinaryPart(chartPath);
                if (raw.Length > 512 * 1024 || charts.Count >= 128) throw new InvalidDataException("图表定义超过容量限制。");
                try
                {
                    var styleParts = ReadStyleParts(package, chartPath);
                    var styleReader = new ChartStyleReader(colors, package, chartPath, styleParts);
                    if (extended)
                    {
                        var extendedModel = ReadExtended(xml, styleReader) with { Anchor = XlsxDrawingAnchors.Read(anchor, geometry), OriginalXml = [.. raw], StyleParts = styleParts };
                        extendedModel = extendedModel with { StyleParts = ExtendedStyleParts(extendedModel, Enumerable.Range(0, 6).Select(styleReader.Palette).ToArray()) };
                        extendedModel.Validate(); package.Budget.Parse(extendedModel.EstimatedBytes); charts.Add(extendedModel); continue;
                    }
                    if (xml.Name != Chart + "chartSpace") throw new InvalidDataException("图表根元素无效。");
                    var filters = NormalizeStandardFilters(xml);
                    ValidateShape(xml);
                    var chart = xml.Element(Chart + "chart") ?? throw new InvalidDataException("图表定义缺失。"); var plot = chart.Element(Chart + "plotArea") ?? throw new InvalidDataException("图表绘图区缺失。");
                    var kinds = plot.Elements().Where(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
                    if (kinds.Length != 1) throw new NotSupportedException("组合或非标准图表");
                    var graph = kinds[0]; var kind = graph.Name.LocalName switch { "barChart" => WorksheetChartKind.Column, "lineChart" => WorksheetChartKind.Line, "pieChart" => WorksheetChartKind.Pie, "scatterChart" => WorksheetChartKind.Scatter, _ => throw new NotSupportedException("此图表类型") };
                    if (kind == WorksheetChartKind.Column && (Value(graph, "barDir") != "col" || Value(graph, "grouping") is not (null or "clustered")) ||
                        kind == WorksheetChartKind.Line && Value(graph, "grouping") is not (null or "standard") || Value(graph, "firstSliceAng") is not (null or "0")) throw new NotSupportedException("堆积、水平或旋转图表");
                    var unsupported = new HashSet<string>(["trendline", "errBars", "explosion", "dateAx", "serAx", "manualLayout", "logBase", "dispUnits", "majorUnit", "minorUnit", "externalData", "pivotSource", "view3D", "userShapes", "pictureOptions", "multiLvlStrRef", "numLit", "strLit", "titleRef"]);
                    if (xml.Descendants().Any(e => unsupported.Contains(e.Name.LocalName)) || xml.Descendants(Chart + "scaling").Any(e => Value(e, "orientation") is not (null or "minMax")) ||
                        xml.Descendants(Chart + "dLbls").Elements().Any(e => e.Name.LocalName is "showLegendKey" or "showPercent" or "showBubbleSize" && (string?)e.Attribute("val") is "1" or "true") ||
                        xml.Descendants(Chart + "delete").Any(e => (string?)e.Attribute("val") is "1" or "true") || Value(chart, "dispBlanksAs") is not (null or "gap")) throw new NotSupportedException("图表含尚未显示的轴、标记、标签或其他语义");
                    // Office's known identity and NA-as-blank extensions carry no source references.
                    foreach (var ext in xml.Descendants(Chart + "ext"))
                    {
                        XNamespace c16 = "http://schemas.microsoft.com/office/drawing/2014/chart", c16r3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart";
                        var child = ext.Elements().FirstOrDefault();
                        var valid = (string?)ext.Attribute("uri") switch
                        {
                            "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" => child?.Name == c16 + "uniqueId" && !child.HasElements && Guid.TryParse((string?)child.Attribute("val"), out _),
                            "{56B9EC1D-385E-4148-901F-78D8002777C0}" => child?.Name == c16r3 + "dataDisplayOptions16" && child.Elements().Count() == 1 && child.Element(c16r3 + "dispNaAsBlank") is { HasElements: false } na && (string?)na.Attribute("val") is "0" or "1" or "false" or "true",
                            _ => false
                        };
                        if (!valid || ext.Elements().Count() != 1 || ext.DescendantsAndSelf().Any(e => e.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != (e == ext ? "uri" : "val")))) throw new NotSupportedException("未知图表扩展结构");
                    }
                    var series = new List<ChartSeries>();
                    foreach (var ser in graph.Elements(Chart + "ser"))
                    {
                        string Reference(XElement? container, string referenceType) => container?.Element(Chart + referenceType)?.Element(Chart + "f")?.Value is { } f ? "=" + f.TrimStart('=') : throw new NotSupportedException("非单元格引用的图表数据");
                        var nameRef = ser.Element(Chart + "tx")?.Element(Chart + "strRef") is null ? null : Reference(ser.Element(Chart + "tx"), "strRef");
                        var name = nameRef is null ? ser.Element(Chart + "tx")?.Element(Chart + "v")?.Value ?? "" : "";
                        var cat = ser.Element(Chart + (kind == WorksheetChartKind.Scatter ? "xVal" : "cat")); var numericCategories = cat?.Element(Chart + "strRef") is null;
                        var catRef = Reference(cat, numericCategories ? "numRef" : "strRef");
                        var valueRef = Reference(ser.Element(Chart + (kind == WorksheetChartKind.Scatter ? "yVal" : "val")), "numRef");
                        var shape = ser.Element(Chart + "spPr"); var color = shape?.Element(Art + "ln")?.Element(Art + "noFill") is not null
                            ? styleReader.SeriesColor(ser.Element(Chart + "marker")?.Element(Chart + "spPr"), series.Count, false)
                            : styleReader.SeriesColor(shape, series.Count, kind is WorksheetChartKind.Line or WorksheetChartKind.Scatter);
                        var pointColors = kind == WorksheetChartKind.Pie || ser.Elements(Chart + "dPt").Any() ? Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(valueRef)))
                            .Select(i => kind == WorksheetChartKind.Pie ? styleReader.Palette(i) : color).ToArray() : [];
                        foreach (var point in ser.Elements(Chart + "dPt"))
                        {
                            if (point.Elements().Any(e => e.Name != Chart + "idx" && e.Name != Chart + "spPr" && e.Name != Chart + "extLst")) throw new NotSupportedException("自定义数据点");
                            var pointIndex = TypedStyles.Int(Value(point, "idx") ?? "", pointColors.Length - 1);
                            pointColors[pointIndex] = styleReader.Fill(point.Element(Chart + "spPr"), pointColors[pointIndex]);
                        }
                        var scatterStyle = Value(graph, "scatterStyle") ?? "marker";
                        if (kind == WorksheetChartKind.Scatter && scatterStyle is not ("marker" or "line" or "lineMarker" or "smooth" or "smoothMarker")) throw new NotSupportedException("未知散点图布局");
                        var marker = ser.Element(Chart + "marker"); var labels = ser.Element(Chart + "dLbls") ?? graph.Element(Chart + "dLbls");
                        series.Add(new(name, nameRef, catRef, valueRef, color) { IsFiltered = filters[ser].Hidden, HiddenPoints = filters[ser].Points, PointColors = [.. pointColors], NumericCategories = numericCategories,
                            ConnectPoints = (kind != WorksheetChartKind.Scatter || scatterStyle != "marker") && shape?.Element(Art + "ln")?.Element(Art + "noFill") is null,
                            Smooth = Value(ser, "smooth") is { } smooth ? smooth is "1" or "true" : Value(graph, "smooth") is "1" or "true" || kind == WorksheetChartKind.Scatter && scatterStyle.StartsWith("smooth", StringComparison.Ordinal),
                            Marker = Value(marker, "symbol") ?? (kind == WorksheetChartKind.Scatter && scatterStyle is "marker" or "lineMarker" or "smoothMarker" || Value(graph, "marker") is "1" or "true" ? "circle" : "none"),
                            MarkerSize = Value(marker, "size") is { } size ? double.Parse(size, CultureInfo.InvariantCulture) : 5, LineWidth = styleReader.LineWidth(shape),
                            ShowValues = Value(labels, "showVal") is "1" or "true", ShowCategories = Value(labels, "showCatName") is "1" or "true", ShowSeriesName = Value(labels, "showSerName") is "1" or "true" });
                    }
                    var titleNode = chart.Element(Chart + "title");
                    if (titleNode?.Descendants(Chart + "strRef").Any() == true) throw new NotSupportedException("引用公式的图表标题");
                    if (kind == WorksheetChartKind.Line && series.Count > 0 && series.All(s => s.Smooth)) kind = WorksheetChartKind.SmoothLine;
                    var axis = plot.Elements(Chart + "valAx").FirstOrDefault(a => Value(a, "axPos") == "l");
                    var catAxis = kind == WorksheetChartKind.Scatter ? plot.Elements(Chart + "valAx").FirstOrDefault(a => Value(a, "axPos") == "b") : plot.Element(Chart + "catAx");
                    if (kind != WorksheetChartKind.Pie)
                    {
                        if (plot.Elements(Chart + "catAx").Count() != (kind == WorksheetChartKind.Scatter ? 0 : 1) || plot.Elements(Chart + "valAx").Count() != (kind == WorksheetChartKind.Scatter ? 2 : 1) || Value(catAxis, "axPos") != "b" || Value(axis, "axPos") != "l" ||
                            Value(catAxis, "crossAx") != Value(axis, "axId") || Value(axis, "crossAx") != Value(catAxis, "axId")) throw new NotSupportedException("非标准或多组坐标轴");
                    }
                    double? Limit(string name) => Value(axis?.Element(Chart + "scaling"), name) is { } value ? double.Parse(value, CultureInfo.InvariantCulture) : null;
                    var model = new WorksheetChart { Kind = kind, Title = string.Concat(titleNode?.Descendants(Art + "t").Select(t => t.Value) ?? []), Anchor = XlsxDrawingAnchors.Read(anchor, geometry), Series = [.. series],
                        Legend = chart.Element(Chart + "legend") is null ? null : Value(chart.Element(Chart + "legend"), "legendPos") ?? "r", VisibleOnly = Value(chart, "plotVisOnly") is not ("0" or "false"), Minimum = Limit("min"), Maximum = Limit("max"),
                        NumberFormat = (string?)axis?.Element(Chart + "numFmt")?.Attribute("formatCode") ?? "General", OriginalXml = [.. raw], StyleParts = styleParts,
                        GapWidth = Value(graph, "gapWidth") is { } gap ? int.Parse(gap, CultureInfo.InvariantCulture) : 150, Overlap = Value(graph, "overlap") is { } overlap ? int.Parse(overlap, CultureInfo.InvariantCulture) : 0,
                        XMinimum = kind == WorksheetChartKind.Scatter && Value(catAxis?.Element(Chart + "scaling"), "min") is { } xmin ? double.Parse(xmin, CultureInfo.InvariantCulture) : null,
                        XMaximum = kind == WorksheetChartKind.Scatter && Value(catAxis?.Element(Chart + "scaling"), "max") is { } xmax ? double.Parse(xmax, CultureInfo.InvariantCulture) : null,
                        XNumberFormat = kind == WorksheetChartKind.Scatter ? (string?)catAxis?.Element(Chart + "numFmt")?.Attribute("formatCode") ?? "General" : "General",
                        Appearance = styleReader.Appearance(xml, chart, plot) };
                    model.Validate(); package.Budget.Parse(model.EstimatedBytes); charts.Add(model);
                }
                catch (Exception ex) when (ex is NotSupportedException or ArgumentException or FormatException or OverflowException)
                { package.Loss($"{chartPath}: {ex.Message}，图表保留于原包；修改源数据后不显示旧缓存。", "chart/unsupported", chartPath); }
            }
        }
        return charts.ToImmutable();
    }

    private static void ValidateShape(XElement root)
    {
        // A closed schema prevents unmodeled semantics from being silently lost when caches are rebuilt.
        var containers = new Dictionary<string, string>
        {
            ["chartSpace"] = "date1904 lang roundedCorners style chart printSettings spPr txPr",
            ["chart"] = "title autoTitleDeleted plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst",
            ["title"] = "tx overlay layout spPr txPr", ["tx"] = "rich strRef v", ["plotArea"] = "layout barChart lineChart pieChart scatterChart catAx valAx spPr",
            ["barChart"] = "barDir grouping varyColors ser dLbls gapWidth overlap axId", ["lineChart"] = "grouping varyColors ser dLbls marker smooth axId",
            ["scatterChart"] = "scatterStyle varyColors ser dLbls axId",
            ["pieChart"] = "varyColors ser dLbls firstSliceAng", ["ser"] = "idx order tx spPr invertIfNegative marker dPt dLbls cat val xVal yVal smooth extLst",
            ["cat"] = "strRef numRef", ["val"] = "numRef", ["xVal"] = "numRef", ["yVal"] = "numRef", ["strRef"] = "f strCache", ["numRef"] = "f numCache",
            ["strCache"] = "ptCount pt", ["numCache"] = "formatCode ptCount pt", ["pt"] = "v", ["marker"] = "symbol size spPr", ["dPt"] = "idx spPr extLst",
            ["dLbls"] = "showLegendKey showVal showCatName showSerName showPercent showBubbleSize showLeaderLines dLblPos numFmt spPr txPr separator",
            ["catAx"] = "axId scaling delete axPos numFmt majorTickMark minorTickMark tickLblPos crossAx crosses auto lblAlgn lblOffset noMultiLvlLbl spPr txPr",
            ["valAx"] = "axId scaling delete axPos majorGridlines numFmt majorTickMark minorTickMark tickLblPos crossAx crosses crossBetween spPr txPr",
            ["scaling"] = "orientation max min", ["legend"] = "legendPos overlay layout spPr txPr", ["printSettings"] = "headerFooter pageMargins pageSetup",
            ["layout"] = "", ["majorGridlines"] = "spPr", ["headerFooter"] = "", ["pageMargins"] = "", ["pageSetup"] = ""
        };
        var leaves = "date1904 lang roundedCorners style autoTitleDeleted overlay plotVisOnly dispBlanksAs showDLblsOverMax barDir grouping varyColors gapWidth overlap smooth firstSliceAng axId idx order invertIfNegative symbol size scatterStyle dLblPos separator ptCount f v formatCode showLegendKey showVal showCatName showSerName showPercent showBubbleSize showLeaderLines delete axPos majorTickMark minorTickMark tickLblPos crossAx crosses auto lblAlgn lblOffset noMultiLvlLbl crossBetween orientation max min legendPos numFmt".Split(' ').ToHashSet();
        foreach (var e in root.DescendantsAndSelf())
        {
            if (e.Ancestors(Chart + "extLst").Any() || e.Name == Chart + "extLst") continue;
            if (e.Name.Namespace == Chart)
            {
                var name = e.Name.LocalName;
                if (name is "rich" or "spPr" or "txPr")
                {
                    if (e.Elements().Any(child => child.Name.Namespace != Art)) throw new NotSupportedException("未知图表样式结构");
                    continue;
                }
                if (!containers.ContainsKey(name) && !leaves.Contains(name)) throw new NotSupportedException("未建模图表元素 " + name);
                var allowed = containers.GetValueOrDefault(name, "").Split(' ').ToHashSet();
                foreach (var child in e.Elements())
                    if (!(child.Name.Namespace == Chart && allowed.Contains(child.Name.LocalName)) && !(name == "chartSpace" && child.Name.LocalName == "AlternateContent")) throw new NotSupportedException("未建模图表结构 " + name + "/" + child.Name.LocalName);
                var attributes = name switch { "marker" => "val", "pt" => "idx", "numFmt" => "formatCode sourceLinked", "pageMargins" => "b l r t header footer", "pageSetup" => "", _ => leaves.Contains(name) ? "val" : "" };
                if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && !attributes.Split(' ').Contains(a.Name.LocalName))) throw new NotSupportedException("未建模图表属性 " + name);
            }
            else if (e.Name.Namespace == Art)
            {
                ValidateDrawingElement(e);
            }
            else if (e.Name.LocalName is not ("AlternateContent" or "Choice" or "Fallback" or "style")) throw new NotSupportedException("未知图表命名空间");
        }
        foreach (var e in root.Descendants())
        {
            var v = (string?)e.Attribute("val");
            if (e.Name.Namespace != Chart) continue;
            if (e.Name.LocalName is "overlay" or "roundedCorners" or "invertIfNegative" && v is not (null or "0" or "false") ||
                e.Name.LocalName == "crosses" && v != "autoZero" || e.Name.LocalName == "crossBetween" && v != "between") throw new NotSupportedException("非基础图表布局 " + e.Name.LocalName);
        }
    }

    internal static void Write(WorkbookSnapshot book, WorksheetSnapshot sheet, int sheetNumber, Dictionary<string, byte[]> parts, XElement root, XElement types)
    {
        if (sheet.Data.Charts.IsEmpty) return;
        var path = $"xl/drawings/drawing{sheetNumber}.xml";
        var drawing = parts.TryGetValue(path, out var existing) ? Xml(existing, new(CancellationToken.None)).Root! : new XElement(Drawing + "wsDr", new XAttribute(XNamespace.Xmlns + "a", Art), new XAttribute(XNamespace.Xmlns + "r", Rel));
        var relPath = RelationshipPart(path); var rels = parts.TryGetValue(relPath, out var relationships) ? Xml(relationships, new(CancellationToken.None)).Root! : new XElement(PackageRel + "Relationships");
        for (var i = 0; i < sheet.Data.Charts.Length; i++)
        {
            var chart = sheet.Data.Charts[i]; var id = $"chart{i + 1}"; var chartPath = $"xl/charts/chart{sheetNumber}-{i + 1}.xml";
            var chartNamespace = chart.IsExtended ? ChartEx : Chart;
            var graphic = new XElement(Drawing + "graphicFrame", new XAttribute("macro", ""), new XElement(Drawing + "nvGraphicFramePr",
                new XElement(Drawing + "cNvPr", new XAttribute("id", sheet.Images.Length + i + 1), new XAttribute("name", $"Chart {i + 1}")), new XElement(Drawing + "cNvGraphicFramePr")),
                new XElement(Drawing + "xfrm", new XElement(Art + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(Art + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))),
                new XElement(Art + "graphic", new XElement(Art + "graphicData", new XAttribute("uri", chartNamespace.NamespaceName), new XElement(chartNamespace + "chart", new XAttribute(Rel + "id", id)))));
            drawing.Add(XlsxDrawingAnchors.Write(chart.Anchor, graphic));
            rels.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", chart.IsExtended ? ChartExRelationship : Rel.NamespaceName + "/chart"), new XAttribute("Target", $"../charts/chart{sheetNumber}-{i + 1}.xml")));
            parts[chartPath] = XmlBytes(new XDocument(chart.IsExtended ? ExtendedXml(book, chart) : StandardXml(book, chart)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + chartPath), new XAttribute("ContentType", chart.IsExtended ? ChartExType : ChartType)));
            WriteStyleParts(chart, chartPath, parts, types);
        }
        if (!parts.ContainsKey(path))
        {
            root.Add(new XElement(Main + "drawing", new XAttribute(Rel + "id", "drawing")));
            var sheetRelPath = RelationshipPart($"xl/worksheets/sheet{sheetNumber}.xml");
            var sheetRels = parts.TryGetValue(sheetRelPath, out var bytes) ? Xml(bytes, new(CancellationToken.None)).Root! : new XElement(PackageRel + "Relationships");
            sheetRels.Add(Relationship("drawing", "drawing", $"../drawings/drawing{sheetNumber}.xml")); parts[sheetRelPath] = XmlBytes(new XDocument(sheetRels));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
        }
        parts[path] = XmlBytes(new XDocument(drawing)); parts[relPath] = XmlBytes(new XDocument(rels));
    }

}
