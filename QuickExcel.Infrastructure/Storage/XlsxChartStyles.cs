using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.XlsxDrawingAnchors;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    private static readonly XNamespace ChartStyle = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
    private const string StyleRelationship = "http://schemas.microsoft.com/office/2011/relationships/";
    private static string StyleContentType(string kind) => kind == "chartStyle" ? "application/vnd.ms-office.chartstyle+xml" : "application/vnd.ms-office.chartcolorstyle+xml";

    private static ImmutableArray<ChartStylePart> ReadStyleParts(XlsxPackage package, string path)
    {
        var result = ImmutableArray.CreateBuilder<ChartStylePart>();
        foreach (var relation in package.Relationships(path).Values)
        {
            var kind = relation.Type[(relation.Type.LastIndexOf('/') + 1)..];
            if (relation.External || kind is not ("chartStyle" or "chartColorStyle") || relation.Type != StyleRelationship + kind)
                throw new NotSupportedException("图表含外部数据或未知附加关系");
            if (result.Any(p => p.Kind == kind) || package.Relationships(relation.Target).Count != 0) throw new NotSupportedException("图表样式含重复或附加依赖");
            var target = package.RelatedByUri(path, relation.Type)!;
            package.RequireType(target, StyleContentType(kind));
            var bytes = package.BinaryPart(target);
            if (bytes.Length > 512 * 1024) throw new InvalidDataException("图表样式超过容量限制。");
            var root = package.XmlPart(target).Root!;
            if (root.Name != ChartStyle + (kind == "chartStyle" ? "chartStyle" : "colorStyle")) throw new NotSupportedException("图表样式根元素无效");
            ValidateStylePart(root);
            result.Add(new(relation.Id, kind, [.. bytes]));
        }
        return result.ToImmutable();
    }

    private static void ValidateStylePart(XElement root)
    {
        var names = ("chartStyle colorStyle variation axisTitle categoryAxis chartArea dataLabel dataLabelCallout dataPoint dataPoint3D dataPointLine dataPointMarker dataPointMarkerLayout dataPointWireframe dataTable downBar dropLine errorBar floor gridlineMajor gridlineMinor hiLoLine leaderLine legend plotArea plotArea3D seriesAxis seriesLine title trendline trendlineLabel upBar valueAxis wall lnRef fillRef effectRef fontRef defRPr spPr bodyPr styleClr").Split(' ').ToHashSet();
        foreach (var e in root.DescendantsAndSelf())
        {
            if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.Namespace != XNamespace.None)) throw new NotSupportedException("图表样式含未知属性或引用");
            if (e.Name.Namespace == Art) ValidateDrawingElement(e);
            else if (e.Name.Namespace != ChartStyle || !names.Contains(e.Name.LocalName)) throw new NotSupportedException("未知附加图表样式 " + e.Name.LocalName);
        }
        if (root.Name.LocalName == "colorStyle" && (string?)root.Attribute("meth") is not (null or "cycle")) throw new NotSupportedException("渐变图表调色板");
    }

    private static void ValidateDrawingElement(XElement e)
    {
        var names = "bodyPr lstStyle p pPr defRPr r rPr endParaRPr t br solidFill srgbClr schemeClr sysClr lumMod lumOff tint shade alpha ln noFill effectLst latin ea cs sym round bevel miter prstDash headEnd tailEnd spAutoFit noAutofit normAutofit".Split(' ');
        if (!names.Contains(e.Name.LocalName)) throw new NotSupportedException("未建模图表绘图属性 " + e.Name.LocalName);
        if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.Namespace != XNamespace.None && a.Name != XNamespace.Xml + "space")) throw new NotSupportedException("图表样式含未知属性或引用");
        if (e.Name.LocalName == "effectLst" && e.HasElements) throw new NotSupportedException("图表阴影或特效");
        if (e.Name.LocalName is "headEnd" or "tailEnd" && (string?)e.Attribute("type") is not (null or "none")) throw new NotSupportedException("图表线条箭头");
    }

    private sealed class ChartStyleReader(XlsxColors colors, XlsxPackage package, string path, ImmutableArray<ChartStylePart> parts)
    {
        private readonly XElement? _style = parts.FirstOrDefault(p => p.Kind == "chartStyle") is { } style ? Xml(style.Bytes.ToArray(), new(CancellationToken.None)).Root : null;
        private readonly XElement? _palette = parts.FirstOrDefault(p => p.Kind == "chartColorStyle") is { } palette ? Xml(palette.Bytes.ToArray(), new(CancellationToken.None)).Root : null;

        public string Palette(int index)
        {
            var entries = _palette?.Elements().Where(e => e.Name.Namespace == Art).ToArray() ?? [];
            var color = entries.Length == 0 ? Theme("accent" + (index % 6 + 1)) : Color(entries[index % entries.Length], "FF4472C4");
            var variations = _palette?.Elements(ChartStyle + "variation").ToArray() ?? [];
            return entries.Length > 0 && variations.Length > 0 ? Transform(color, variations[index / entries.Length % variations.Length].Elements()) : color;
        }

        private string Theme(string name)
        {
            var slot = name switch { "lt1" or "bg1" => 0, "dk1" or "tx1" => 1, "lt2" or "bg2" => 2, "dk2" or "tx2" => 3, "accent1" => 4, "accent2" => 5, "accent3" => 6, "accent4" => 7, "accent5" => 8, "accent6" => 9, "hlink" => 10, "folHlink" => 11, _ => throw new NotSupportedException("未知图表主题颜色") };
            return colors.Read(new XElement(Main + "color", new XAttribute("theme", slot)), package, path) ?? "FF4472C4";
        }

        private string Color(XElement? node, string fallback)
        {
            if (node is null) return fallback;
            var value = (string?)node.Attribute("val");
            var color = node.Name.LocalName switch
            {
                "srgbClr" => "FF" + value,
                "sysClr" => "FF" + (string?)node.Attribute("lastClr"),
                "schemeClr" => value == "phClr" ? fallback : Theme(value ?? ""),
                "styleClr" => fallback,
                _ => throw new NotSupportedException("非实色图表填充")
            };
            if (color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c))) throw new NotSupportedException("图表颜色无效");
            return Transform(color, node.Elements());
        }

        private static string Transform(string color, IEnumerable<XElement> transforms)
        {
            var r = int.Parse(color.AsSpan(2, 2), NumberStyles.HexNumber) / 255d; var g = int.Parse(color.AsSpan(4, 2), NumberStyles.HexNumber) / 255d; var b = int.Parse(color.AsSpan(6, 2), NumberStyles.HexNumber) / 255d;
            var alpha = int.Parse(color.AsSpan(0, 2), NumberStyles.HexNumber) / 255d;
            foreach (var t in transforms)
            {
                var value = double.Parse((string?)t.Attribute("val") ?? "0", CultureInfo.InvariantCulture) / 100000;
                if (!double.IsFinite(value) || value is < 0 or > 1) throw new NotSupportedException("图表颜色变换超出范围");
                if (t.Name.LocalName == "alpha") { alpha = value; continue; }
                if (t.Name.LocalName is "tint" or "shade")
                { r = t.Name.LocalName == "tint" ? r * value + 1 - value : r * value; g = t.Name.LocalName == "tint" ? g * value + 1 - value : g * value; b = t.Name.LocalName == "tint" ? b * value + 1 - value : b * value; continue; }
                var max = Math.Max(r, Math.Max(g, b)); var min = Math.Min(r, Math.Min(g, b)); var delta = max - min; var l = (max + min) / 2;
                var s = delta == 0 ? 0 : delta / (1 - Math.Abs(2 * l - 1));
                var h = delta == 0 ? 0 : max == r ? ((g - b) / delta + 6) % 6 : max == g ? (b - r) / delta + 2 : (r - g) / delta + 4;
                l = Math.Clamp(t.Name.LocalName switch { "lumMod" => l * value, "lumOff" => l + value, _ => throw new NotSupportedException("未知图表颜色变换") }, 0, 1);
                var c = (1 - Math.Abs(2 * l - 1)) * s; var x = c * (1 - Math.Abs(h % 2 - 1)); var m = l - c / 2;
                (r, g, b) = h switch { < 1 => (c, x, 0d), < 2 => (x, c, 0d), < 3 => (0d, c, x), < 4 => (0d, x, c), < 5 => (x, 0d, c), _ => (c, 0d, x) };
                r += m; g += m; b += m;
            }
            string Byte(double v) => Math.Clamp((int)Math.Round(v * 255), 0, 255).ToString("X2");
            return Byte(alpha) + Byte(r) + Byte(g) + Byte(b);
        }

        public string Fill(XElement? shape, string fallback)
        {
            if (shape?.Element(Art + "noFill") is not null) return "00000000";
            return Color(shape?.Element(Art + "solidFill")?.Elements().FirstOrDefault(), fallback);
        }
        public string SeriesColor(XElement? shape, int index, bool line) => Fill(line ? shape?.Element(Art + "ln") : shape, Palette(index));
        public double LineWidth(XElement? shape, double fallback = 2) => shape?.Element(Art + "ln")?.Attribute("w") is { } w ? double.Parse(w.Value, CultureInfo.InvariantCulture) / 9525 : fallback;
        private XElement? Slot(string name) => _style?.Element(ChartStyle + name);
        private ChartTextStyle Text(XElement? node, string slot, ChartTextStyle fallback)
        {
            var defaults = Slot(slot); var run = node?.Descendants(Art + "defRPr").FirstOrDefault() ?? node?.Descendants(Art + "rPr").FirstOrDefault();
            var styleRun = defaults?.Element(ChartStyle + "defRPr");
            string? Attribute(string name) => (string?)run?.Attribute(name) ?? (string?)styleRun?.Attribute(name);
            var font = (string?)run?.Element(Art + "latin")?.Attribute("typeface");
            if (font?.StartsWith('+') == true) font = colors.Font(font.Contains("mj", StringComparison.Ordinal) ? "major" : "minor", null).Latin;
            return new(Fill(run, Color(defaults?.Element(ChartStyle + "fontRef")?.Elements().FirstOrDefault(), fallback.Color)), font ?? fallback.FontFamily,
                Attribute("sz") is { } size ? double.Parse(size, CultureInfo.InvariantCulture) / 100 : fallback.Size,
                Attribute("b") is { } bold ? bold is "1" or "true" : fallback.Bold, Attribute("i") is { } italic ? italic is "1" or "true" : fallback.Italic);
        }
        public ChartAppearance Appearance(XElement root, XElement chart, XElement plot)
        {
            var ns = root.Name.Namespace; var shape = root.Element(ns + "spPr"); var defaults = Slot("chartArea")?.Element(ChartStyle + "spPr");
            var axis = plot.Elements().FirstOrDefault(e => e.Name.LocalName is "catAx" or "axis" || e.Name.LocalName == "valAx" && Value(e, "axPos") == "b");
            var grid = plot.Descendants(ns + "majorGridlines").FirstOrDefault()?.Element(ns + "spPr");
            return new()
            {
                Background = Fill(shape, Fill(defaults, "FFFFFFFF")), PlotBackground = Fill(plot.Element(ns + "spPr"), Fill(Slot("plotArea")?.Element(ChartStyle + "spPr"), "00FFFFFF")),
                BorderColor = Fill(shape?.Element(Art + "ln"), Fill(defaults?.Element(Art + "ln"), "FFD3D3D3")), BorderWidth = LineWidth(shape, LineWidth(defaults, 1)),
                GridColor = Fill(grid?.Element(Art + "ln"), Fill(Slot("gridlineMajor")?.Element(ChartStyle + "spPr")?.Element(Art + "ln"), "FFD3D3D3")),
                Title = Text(chart.Element(ns + "title"), "title", new(Size: 12)), Axis = Text(axis, "categoryAxis", new()), Legend = Text(chart.Element(ns + "legend"), "legend", new())
            };
        }
    }

    private static void WriteStyleParts(WorksheetChart chart, string chartPath, Dictionary<string, byte[]> parts, XElement types)
    {
        var styleParts = chart.StyleParts;
        if (chart.Kind == WorksheetChartKind.Waterfall && !styleParts.Any(p => p.Kind == "chartColorStyle"))
        {
            var palette = new XElement(ChartStyle + "colorStyle", new XAttribute(XNamespace.Xmlns + "cs", ChartStyle), new XAttribute(XNamespace.Xmlns + "a", Art), new XAttribute("meth", "cycle"), new XAttribute("id", 10),
                new[] { chart.IncreaseColor, chart.DecreaseColor, chart.TotalColor }.Select(c => new XElement(Art + "srgbClr", new XAttribute("val", c[2..]))), new XElement(ChartStyle + "variation"));
            var id = "chartColors"; while (styleParts.Any(p => p.RelationshipId == id)) id += "1";
            styleParts = styleParts.Add(new(id, "chartColorStyle", [.. XmlBytes(new XDocument(palette))]));
        }
        if (styleParts.IsEmpty) return;
        var rels = new XElement(PackageRel + "Relationships");
        foreach (var part in styleParts)
        {
            var name = Path.GetFileNameWithoutExtension(chartPath) + "-" + part.Kind + ".xml";
            var path = "xl/charts/" + name;
            parts[path] = part.Bytes.ToArray();
            rels.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", part.RelationshipId), new XAttribute("Type", StyleRelationship + part.Kind), new XAttribute("Target", name)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", StyleContentType(part.Kind))));
        }
        parts[RelationshipPart(chartPath)] = XmlBytes(new XDocument(rels));
    }
}
