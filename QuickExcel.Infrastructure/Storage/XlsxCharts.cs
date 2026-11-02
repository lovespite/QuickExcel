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

internal static class XlsxCharts
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
                var reference = frame.Element(Art + "graphic")?.Element(Art + "graphicData")?.Element(Chart + "chart");
                if (reference is null) { package.Loss($"{path}: 非图表绘图尚未支持", "drawing/graphicFrame", path); continue; }
                var chartPath = package.Related(path, Attr(reference, Rel + "id"), "chart");
                package.RequireType(chartPath, ChartType); var xml = package.XmlPart(chartPath).Root!; var raw = package.BinaryPart(chartPath);
                if (raw.Length > 512 * 1024 || charts.Count >= 128) throw new InvalidDataException("图表定义超过容量限制。");
                try
                {
                    if (package.Relationships(chartPath).Count != 0) throw new NotSupportedException("含外部数据或附加样式关系的图表");
                    if (xml.Name != Chart + "chartSpace") throw new InvalidDataException("图表根元素无效。");
                    ValidateShape(xml);
                    var chart = xml.Element(Chart + "chart") ?? throw new InvalidDataException("图表定义缺失。"); var plot = chart.Element(Chart + "plotArea") ?? throw new InvalidDataException("图表绘图区缺失。");
                    var kinds = plot.Elements().Where(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
                    if (kinds.Length != 1) throw new NotSupportedException("组合或非标准图表");
                    var graph = kinds[0]; var kind = graph.Name.LocalName switch { "barChart" => WorksheetChartKind.Column, "lineChart" => WorksheetChartKind.Line, "pieChart" => WorksheetChartKind.Pie, _ => throw new NotSupportedException("此图表类型") };
                    if (kind == WorksheetChartKind.Column && (Value(graph, "barDir") != "col" || Value(graph, "grouping") is not (null or "clustered")) ||
                        kind == WorksheetChartKind.Line && Value(graph, "grouping") is not (null or "standard") || Value(graph, "firstSliceAng") is not (null or "0")) throw new NotSupportedException("堆积、水平或旋转图表");
                    var unsupported = new HashSet<string>(["trendline", "errBars", "explosion", "dateAx", "serAx", "manualLayout", "logBase", "dispUnits", "majorUnit", "minorUnit", "externalData", "pivotSource", "view3D", "userShapes", "pictureOptions", "multiLvlStrRef", "numLit", "strLit", "titleRef"]);
                    if (xml.Descendants().Any(e => unsupported.Contains(e.Name.LocalName)) || xml.Descendants(Chart + "smooth").Any(e => (string?)e.Attribute("val") == "1") ||
                        xml.Descendants(Chart + "marker").Any(e => Value(e, "symbol") is not (null or "none")) || xml.Descendants(Chart + "scaling").Any(e => Value(e, "orientation") is not (null or "minMax")) ||
                        xml.Descendants(Chart + "dLbls").Elements().Any(e => e.Name.LocalName != "showLeaderLines" && (string?)e.Attribute("val") != "0") ||
                        xml.Descendants(Chart + "delete").Any(e => (string?)e.Attribute("val") == "1") || Value(chart, "dispBlanksAs") is not (null or "gap")) throw new NotSupportedException("图表含尚未显示的轴、标记、标签或其他语义");
                    // Office's known identity and NA-as-blank extensions carry no source references.
                    foreach (var ext in xml.Descendants(Chart + "ext"))
                    {
                        XNamespace c16 = "http://schemas.microsoft.com/office/drawing/2014/chart", c16r3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart";
                        var child = ext.Elements().FirstOrDefault();
                        var valid = (string?)ext.Attribute("uri") switch
                        {
                            "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" => child?.Name == c16 + "uniqueId" && !child.HasElements && Guid.TryParse((string?)child.Attribute("val"), out _),
                            "{56B9EC1D-385E-4148-901F-78D8002777C0}" => child?.Name == c16r3 + "dataDisplayOptions16" && child.Elements().Count() == 1 && child.Element(c16r3 + "dispNaAsBlank") is { HasElements: false } na && (string?)na.Attribute("val") is "0" or "1",
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
                        var cat = ser.Element(Chart + "cat"); var catRef = Reference(cat, cat?.Element(Chart + "strRef") is not null ? "strRef" : "numRef");
                        var valueRef = Reference(ser.Element(Chart + "val"), "numRef");
                        var explicitColor = ser.Element(Chart + "spPr")?.Element(Art + "solidFill")?.Element(Art + "srgbClr") ?? ser.Element(Chart + "spPr")?.Element(Art + "ln")?.Element(Art + "solidFill")?.Element(Art + "srgbClr");
                        var color = explicitColor is null ? colors.Read(new XElement(Main + "color", new XAttribute("theme", 4 + series.Count % 6)), package, chartPath) ?? "FF4472C4" : "FF" + Attr(explicitColor, "val");
                        var pointColors = kind == WorksheetChartKind.Pie ? Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(valueRef))).Select(i =>
                            colors.Read(new XElement(Main + "color", new XAttribute("theme", 4 + i % 6)), package, chartPath) ?? "FF4472C4").ToArray() : [];
                        foreach (var point in ser.Elements(Chart + "dPt"))
                        {
                            if (kind != WorksheetChartKind.Pie || point.Elements().Any(e => e.Name != Chart + "idx" && e.Name != Chart + "spPr")) throw new NotSupportedException("自定义数据点");
                            var pointIndex = TypedStyles.Int(Value(point, "idx") ?? "", pointColors.Length - 1);
                            var fill = point.Element(Chart + "spPr")?.Element(Art + "solidFill")?.Element(Art + "srgbClr") ?? throw new NotSupportedException("数据点非实色填充");
                            pointColors[pointIndex] = "FF" + Attr(fill, "val");
                        }
                        series.Add(new(name, nameRef, catRef, valueRef, color) { PointColors = [.. pointColors] });
                    }
                    var titleNode = chart.Element(Chart + "title");
                    if (titleNode?.Descendants(Chart + "strRef").Any() == true) throw new NotSupportedException("引用公式的图表标题");
                    var axis = plot.Element(Chart + "valAx");
                    if (kind != WorksheetChartKind.Pie)
                    {
                        var catAxis = plot.Element(Chart + "catAx");
                        if (plot.Elements(Chart + "catAx").Count() != 1 || plot.Elements(Chart + "valAx").Count() != 1 || Value(catAxis, "axPos") != "b" || Value(axis, "axPos") != "l" ||
                            Value(catAxis, "crossAx") != Value(axis, "axId") || Value(axis, "crossAx") != Value(catAxis, "axId")) throw new NotSupportedException("非标准或多组坐标轴");
                    }
                    double? Limit(string name) => Value(axis?.Element(Chart + "scaling"), name) is { } value ? double.Parse(value, CultureInfo.InvariantCulture) : null;
                    var model = new WorksheetChart { Kind = kind, Title = string.Concat(titleNode?.Descendants(Art + "t").Select(t => t.Value) ?? []), Anchor = XlsxDrawingAnchors.Read(anchor, geometry), Series = [.. series],
                        Legend = chart.Element(Chart + "legend") is null ? null : Value(chart.Element(Chart + "legend"), "legendPos") ?? "r", VisibleOnly = Value(chart, "plotVisOnly") is not "0", Minimum = Limit("min"), Maximum = Limit("max"),
                        NumberFormat = (string?)axis?.Element(Chart + "numFmt")?.Attribute("formatCode") ?? "General", OriginalXml = [.. raw] };
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
            ["chartSpace"] = "date1904 lang roundedCorners style chart printSettings",
            ["chart"] = "title autoTitleDeleted plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst",
            ["title"] = "tx overlay", ["tx"] = "rich strRef v", ["plotArea"] = "layout barChart lineChart pieChart catAx valAx",
            ["barChart"] = "barDir grouping varyColors ser dLbls gapWidth overlap axId", ["lineChart"] = "grouping varyColors ser dLbls smooth axId",
            ["pieChart"] = "varyColors ser dLbls firstSliceAng", ["ser"] = "idx order tx spPr invertIfNegative marker dPt cat val smooth extLst",
            ["cat"] = "strRef numRef", ["val"] = "numRef", ["strRef"] = "f strCache", ["numRef"] = "f numCache",
            ["strCache"] = "ptCount pt", ["numCache"] = "formatCode ptCount pt", ["pt"] = "v", ["marker"] = "symbol", ["dPt"] = "idx spPr",
            ["dLbls"] = "showLegendKey showVal showCatName showSerName showPercent showBubbleSize showLeaderLines",
            ["catAx"] = "axId scaling delete axPos numFmt majorTickMark minorTickMark tickLblPos crossAx crosses auto lblAlgn lblOffset noMultiLvlLbl",
            ["valAx"] = "axId scaling delete axPos majorGridlines numFmt majorTickMark minorTickMark tickLblPos crossAx crosses crossBetween",
            ["scaling"] = "orientation max min", ["legend"] = "legendPos overlay", ["printSettings"] = "headerFooter pageMargins pageSetup",
            ["layout"] = "", ["majorGridlines"] = "", ["headerFooter"] = "", ["pageMargins"] = "", ["pageSetup"] = ""
        };
        var leaves = "date1904 lang roundedCorners style autoTitleDeleted overlay plotVisOnly dispBlanksAs showDLblsOverMax barDir grouping varyColors gapWidth overlap smooth firstSliceAng axId idx order invertIfNegative symbol ptCount f v formatCode showLegendKey showVal showCatName showSerName showPercent showBubbleSize showLeaderLines delete axPos majorTickMark minorTickMark tickLblPos crossAx crosses auto lblAlgn lblOffset noMultiLvlLbl crossBetween orientation max min legendPos numFmt".Split(' ').ToHashSet();
        foreach (var e in root.DescendantsAndSelf())
        {
            if (e.Ancestors(Chart + "extLst").Any() || e.Name == Chart + "extLst") continue;
            if (e.Name.Namespace == Chart)
            {
                var name = e.Name.LocalName;
                if (name is "rich" or "spPr") continue;
                if (!containers.ContainsKey(name) && !leaves.Contains(name)) throw new NotSupportedException("未建模图表元素 " + name);
                var allowed = containers.GetValueOrDefault(name, "").Split(' ').ToHashSet();
                foreach (var child in e.Elements())
                    if (!(child.Name.Namespace == Chart && allowed.Contains(child.Name.LocalName)) && !(name == "chartSpace" && child.Name.LocalName == "AlternateContent")) throw new NotSupportedException("未建模图表结构 " + name + "/" + child.Name.LocalName);
                var attributes = name switch { "pt" => "idx", "numFmt" => "formatCode sourceLinked", "pageMargins" => "b l r t header footer", "pageSetup" => "", _ => leaves.Contains(name) ? "val" : "" };
                if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && !attributes.Split(' ').Contains(a.Name.LocalName))) throw new NotSupportedException("未建模图表属性 " + name);
            }
            else if (e.Name.Namespace == Art)
            {
                if (e.Name.LocalName is not ("bodyPr" or "lstStyle" or "p" or "pPr" or "defRPr" or "r" or "rPr" or "t" or "solidFill" or "srgbClr" or "ln" or "noFill")) throw new NotSupportedException("未建模图表绘图属性 " + e.Name.LocalName);
                var attributes = e.Name.LocalName switch { "rPr" => "lang altLang", "srgbClr" => "val", "ln" => "w", _ => "" };
                if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && !attributes.Split(' ').Contains(a.Name.LocalName)) || e.Name.LocalName == "srgbClr" && e.HasElements) throw new NotSupportedException("自定义图表文字、颜色或线条");
            }
            else if (e.Name.LocalName is not ("AlternateContent" or "Choice" or "Fallback" or "style")) throw new NotSupportedException("未知图表命名空间");
        }
        foreach (var e in root.Descendants())
        {
            var v = (string?)e.Attribute("val");
            if (e.Name.Namespace != Chart) continue;
            if (e.Name.LocalName is "overlay" or "roundedCorners" or "invertIfNegative" or "overlap" && v is not (null or "0") ||
                e.Name.LocalName == "gapWidth" && v != "150" || e.Name.LocalName == "crosses" && v != "autoZero" ||
                e.Name.LocalName == "crossBetween" && v != "between" || e.Name.LocalName == "style" && v is not ("2" or "102")) throw new NotSupportedException("非基础图表布局 " + e.Name.LocalName);
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
            var graphic = new XElement(Drawing + "graphicFrame", new XAttribute("macro", ""), new XElement(Drawing + "nvGraphicFramePr",
                new XElement(Drawing + "cNvPr", new XAttribute("id", sheet.Images.Length + i + 1), new XAttribute("name", $"Chart {i + 1}")), new XElement(Drawing + "cNvGraphicFramePr")),
                new XElement(Drawing + "xfrm", new XElement(Art + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(Art + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))),
                new XElement(Art + "graphic", new XElement(Art + "graphicData", new XAttribute("uri", Chart.NamespaceName), new XElement(Chart + "chart", new XAttribute(Rel + "id", id)))));
            drawing.Add(XlsxDrawingAnchors.Write(chart.Anchor, graphic)); rels.Add(Relationship(id, "chart", $"../charts/chart{sheetNumber}-{i + 1}.xml"));
            parts[chartPath] = XmlBytes(new XDocument(ChartXml(book, chart)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + chartPath), new XAttribute("ContentType", ChartType)));
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

    private static XElement ChartXml(WorkbookSnapshot book, WorksheetChart model)
    {
        // Cache arrays always include hidden source cells. plotVisOnly controls drawing, not reference indexing.
        var data = ChartData.Read(book, model with { VisibleOnly = false });
        var root = model.OriginalXml.IsEmpty ? new XElement(Chart + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", Chart), new XAttribute(XNamespace.Xmlns + "a", Art),
            Val("date1904", book.DateSystem == ExcelDateSystem.Mac1904 ? 1 : 0), new XElement(Chart + "chart", new XElement(Chart + "plotArea", new XElement(Chart + "layout")))) : Xml(model.OriginalXml.ToArray(), new(CancellationToken.None)).Root!;
        var chart = root.Element(Chart + "chart")!; chart.Element(Chart + "title")?.Remove();
        if (model.Title.Length > 0) chart.AddFirst(new XElement(Chart + "title", new XElement(Chart + "tx", new XElement(Chart + "rich", new XElement(Art + "bodyPr"), new XElement(Art + "lstStyle"),
            new XElement(Art + "p", new XElement(Art + "r", new XElement(Art + "t", model.Title))))), Val("overlay", 0)));
        var plot = chart.Element(Chart + "plotArea")!; plot.Elements().Where(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal) || e.Name.LocalName is "catAx" or "valAx").Remove();
        var graph = new XElement(Chart + (model.Kind switch { WorksheetChartKind.Column => "barChart", WorksheetChartKind.Line => "lineChart", _ => "pieChart" }));
        if (model.Kind == WorksheetChartKind.Column) graph.Add(Val("barDir", "col"));
        if (model.Kind != WorksheetChartKind.Pie) graph.Add(Val("grouping", model.Kind == WorksheetChartKind.Line ? "standard" : "clustered"));
        graph.Add(Val("varyColors", model.Kind == WorksheetChartKind.Pie ? 1 : 0));
        XElement Cache(string formula, IEnumerable<string?> values, bool numeric)
        {
            var array = values.ToArray();
            return new XElement(Chart + (numeric ? "numRef" : "strRef"), new XElement(Chart + "f", formula.TrimStart('=')),
                new XElement(Chart + (numeric ? "numCache" : "strCache"), numeric ? new XElement(Chart + "formatCode", model.NumberFormat) : null,
                    Val("ptCount", array.Length), array.Select((v, i) => v is null ? null : new XElement(Chart + "pt", new XAttribute("idx", i), new XElement(Chart + "v", v)))));
        }
        for (var i = 0; i < model.Series.Length; i++)
        {
            var series = model.Series[i]; var values = data[i];
            var color = new XElement(Art + "solidFill", new XElement(Art + "srgbClr", new XAttribute("val", series.Color[2..])));
            var ser = new XElement(Chart + "ser", Val("idx", i), Val("order", i), new XElement(Chart + "tx", series.NameReference is { } name ? Cache(name, [values.Name], false) : new XElement(Chart + "v", series.Name)),
                new XElement(Chart + "spPr", model.Kind == WorksheetChartKind.Line ? new XElement(Art + "ln", new XAttribute("w", 19050), color) : color));
            if (model.Kind == WorksheetChartKind.Line) ser.Add(new XElement(Chart + "marker", Val("symbol", "none")));
            if (model.Kind == WorksheetChartKind.Pie)
                foreach (var (pointColor, pointIndex) in series.PointColors.Select((c, n) => (c, n)))
                    ser.Add(new XElement(Chart + "dPt", Val("idx", pointIndex), new XElement(Chart + "spPr", new XElement(Art + "solidFill", new XElement(Art + "srgbClr", new XAttribute("val", pointColor[2..]))))));
            ser.Add(new XElement(Chart + "cat", Cache(series.Categories, values.Categories, false)), new XElement(Chart + "val", Cache(series.Values, values.Values.Select(v => v?.ToString("R", CultureInfo.InvariantCulture)), true)));
            if (model.Kind == WorksheetChartKind.Line) ser.Add(Val("smooth", 0)); graph.Add(ser);
        }
        if (model.Kind == WorksheetChartKind.Column) graph.Add(Val("gapWidth", 150));
        if (model.Kind != WorksheetChartKind.Pie) graph.Add(Val("axId", 100), Val("axId", 200));
        plot.Add(graph);
        if (model.Kind != WorksheetChartKind.Pie)
        {
            XElement Axis(bool value) => new(Chart + (value ? "valAx" : "catAx"), Val("axId", value ? 200 : 100), new XElement(Chart + "scaling", Val("orientation", "minMax"),
                value && model.Maximum is { } max ? Val("max", max) : null, value && model.Minimum is { } min ? Val("min", min) : null), Val("delete", 0), Val("axPos", value ? "l" : "b"),
                value ? new XElement(Chart + "majorGridlines") : null, new XElement(Chart + "numFmt", new XAttribute("formatCode", model.NumberFormat), new XAttribute("sourceLinked", 0)),
                Val("majorTickMark", "out"), Val("minorTickMark", "none"), Val("tickLblPos", "nextTo"), Val("crossAx", value ? 100 : 200), Val("crosses", "autoZero"),
                value ? Val("crossBetween", "between") : Val("auto", 1));
            plot.Add(Axis(false), Axis(true));
        }
        chart.Element(Chart + "legend")?.Remove();
        if (model.Legend is { } legend) plot.AddAfterSelf(new XElement(Chart + "legend", Val("legendPos", legend), Val("overlay", 0)));
        foreach (var name in new[] { "plotVisOnly", "dispBlanksAs" }) chart.Element(Chart + name)?.Remove();
        var visible = Val("plotVisOnly", model.VisibleOnly ? 1 : 0); var blanks = Val("dispBlanksAs", "gap");
        (chart.Element(Chart + "legend") ?? plot).AddAfterSelf(visible, blanks);
        return root;
    }
}
