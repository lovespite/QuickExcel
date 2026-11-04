using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static partial class XlsxConditionalFormatting
{
    private static readonly XNamespace X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static readonly XNamespace Xm = "http://schemas.microsoft.com/office/excel/2006/main";
    // Excel recognizes this URI case-sensitively: uppercasing the 'd' silently loses the bar link.
    private const string VisualExtension = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}", BarLinkExtension = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
    private static readonly string[] ThresholdTypes = ["min", "max", "num", "percent", "percentile", "formula", "autoMin", "autoMax"];
    private static bool Flag(XElement element, string name, bool fallback = true) => element.Attribute(name) is { } a ? Bool(a.Value) : fallback;
    private static void ShapeVisual(XElement element, XNamespace ns, string children, string attributes, params XName[] foreign)
    {
        var allowed = children.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(n => ns + n).Concat(foreign).ToArray();
        if (element.Elements().Any(e => !allowed.Contains(e.Name)) || element.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.NamespaceName.Length != 0 || !attributes.Split(' ').Contains(a.Name.LocalName))))
            throw new NotSupportedException("未支持的视觉条件格式属性或子元素：" + element.Name.LocalName);
    }
    private static ConditionalThreshold ReadThreshold(XElement element, bool extended = false)
    {
        ShapeVisual(element, extended ? X14 : Main, "", extended ? "type gte" : "type val gte", extended ? [Xm + "f"] : []);
        var kind = Array.IndexOf(ThresholdTypes, Attr(element, "type")); if (kind < 0) throw new NotSupportedException("未支持的条件格式阈值类型。");
        if (element.Elements().Count() > 1) throw new NotSupportedException("阈值公式重复。");
        var raw = extended ? element.Element(Xm + "f")?.Value : (string?)element.Attribute("val");
        if (element.Element(Xm + "f") is { } formula) ShapeVisual(formula, Xm, "", "");
        double number = 0; string? text = null;
        if (kind == 5) text = XlsxStructuredFormula.Read("=" + Decode(raw ?? throw new NotSupportedException("阈值公式缺失。")).TrimStart('='));
        else if (kind is >= 2 and <= 4)
        {
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number) || !double.IsFinite(number)) throw new NotSupportedException("阈值数字无效。");
        }
        else if (extended && raw is not null) throw new NotSupportedException("最小/最大阈值不允许公式。");
        return new((ConditionalThresholdKind)kind, number, text, Flag(element, "gte"));
    }
    private static ConditionalColorScale ReadScale(XElement element, XlsxColors colors, XlsxPackage package, string path)
    {
        XlsxTables.Shape(element, "cfvo color", "");
        return new(element.Elements(Main + "cfvo").Select(e => ReadThreshold(e)),
            element.Elements(Main + "color").Select(e => colors.Read(e, package, path) ?? throw new NotSupportedException("色阶颜色缺失。")));
    }
    private static ConditionalIconSet ReadIcons(XElement element)
    {
        XlsxTables.Shape(element, "cfvo", "iconSet showValue percent reverse");
        return new((string?)element.Attribute("iconSet") ?? "3TrafficLights1", element.Elements(Main + "cfvo").Select(e => ReadThreshold(e)), Flag(element, "showValue"), Flag(element, "reverse", false));
    }
    private static ConditionalDataBar ReadBar(XElement element, XlsxColors colors, XlsxPackage package, string path)
    {
        XlsxTables.Shape(element, "cfvo color", "minLength maxLength showValue");
        var thresholds = element.Elements(Main + "cfvo").Select(e => ReadThreshold(e)).ToArray();
        if (thresholds.Length != 2 || element.Elements(Main + "color").Count() != 1) throw new NotSupportedException("数据条需要两个阈值和一种颜色。");
        var color = colors.Read(element.Element(Main + "color"), package, path) ?? throw new NotSupportedException("数据条颜色缺失。");
        return new(thresholds[0], thresholds[1], color, ShowValue: Flag(element, "showValue"), MinLength: Int((string?)element.Attribute("minLength") ?? "10", 100),
            MaxLength: Int((string?)element.Attribute("maxLength") ?? "90", 100), Axis: DataBarAxis.None, NegativeColor: color);
    }
    private static string ReadBarLink(XElement list)
    {
        XlsxTables.Shape(list, "ext", "");
        if (list.Elements().Count() != 1) throw new NotSupportedException("数据条关联缺失或重复。");
        var ext = list.Elements().Single();
        ShapeVisual(ext, Main, "", "uri", X14 + "id");
        if (!string.Equals((string?)ext.Attribute("uri"), BarLinkExtension, StringComparison.OrdinalIgnoreCase) || ext.Elements().Count() != 1 || !Guid.TryParse(ext.Element(X14 + "id")?.Value, out var id))
            throw new NotSupportedException("未支持的规则扩展。");
        ShapeVisual(ext.Element(X14 + "id")!, X14, "", ""); return id.ToString("B");
    }
    private static Dictionary<string, (CellRectangle[] Ranges, XElement Bar)> ReadExtendedBars(XElement sheet)
    {
        var result = new Dictionary<string, (CellRectangle[], XElement)>(StringComparer.OrdinalIgnoreCase);
        var containers = sheet.Elements(Main + "extLst").Elements(Main + "ext").Where(e => string.Equals((string?)e.Attribute("uri"), VisualExtension, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (containers.Length > 1) throw new NotSupportedException("条件格式扩展容器重复。");
        if (sheet.Descendants().Any(e => e.Name.Namespace != Main && e.Name.LocalName is "conditionalFormatting" or "conditionalFormattings" && !e.Ancestors().Any(containers.Contains)))
            throw new NotSupportedException("未知条件格式扩展。");
        foreach (var ext in containers)
        {
            ShapeVisual(ext, Main, "", "uri", X14 + "conditionalFormattings");
            if (ext.Elements().Count() != 1) throw new NotSupportedException("条件格式扩展容器缺失或重复。");
            var root = ext.Elements().Single();
            ShapeVisual(root, X14, "conditionalFormatting", "");
            foreach (var group in root.Elements())
            {
                ShapeVisual(group, X14, "cfRule", "", Xm + "sqref");
                if (group.Elements(Xm + "sqref").Count() != 1) throw new NotSupportedException("扩展规则区域缺失或重复。");
                ShapeVisual(group.Element(Xm + "sqref")!, Xm, "", "");
                var ranges = group.Element(Xm + "sqref")!.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse).ToArray();
                foreach (var rule in group.Elements(X14 + "cfRule"))
                {
                    ShapeVisual(rule, X14, "dataBar", "type id");
                    if ((string?)rule.Attribute("type") != "dataBar" || rule.Elements().Count() != 1 || !Guid.TryParse((string?)rule.Attribute("id"), out var id) ||
                        !result.TryAdd(id.ToString("B"), (ranges, rule.Element(X14 + "dataBar")!))) throw new NotSupportedException("独立、重复或未支持的视觉规则扩展。");
                }
            }
        }
        return result;
    }
    private static ConditionalDataBar ReadExtendedBar(XElement element, ConditionalDataBar basic, XlsxColors colors, XlsxPackage package, string path)
    {
        ShapeVisual(element, X14, "cfvo borderColor negativeFillColor negativeBorderColor axisColor", "minLength maxLength showValue border gradient direction negativeBarColorSameAsPositive negativeBarBorderColorSameAsPositive axisPosition");
        foreach (var name in new[] { "borderColor", "negativeFillColor", "negativeBorderColor", "axisColor" })
            if (element.Elements(X14 + name).Count() > 1) throw new NotSupportedException("数据条颜色重复。");
        var values = element.Elements(X14 + "cfvo").Select(e => ReadThreshold(e, true)).ToArray();
        if (values.Length != 2) throw new NotSupportedException("扩展数据条需要两个阈值。");
        var axis = ((string?)element.Attribute("axisPosition") ?? "automatic") switch { "automatic" => DataBarAxis.Automatic, "middle" => DataBarAxis.Middle, "none" => DataBarAxis.None, _ => throw new NotSupportedException("未知数据条轴位置。") };
        var direction = ((string?)element.Attribute("direction") ?? "context") switch { "context" => DataBarDirection.Context, "leftToRight" => DataBarDirection.LeftToRight, "rightToLeft" => DataBarDirection.RightToLeft, _ => throw new NotSupportedException("未知数据条方向。") };
        string Color(string name) => colors.Read(element.Element(X14 + name), package, path) ?? throw new NotSupportedException("扩展数据条颜色缺失：" + name);
        var border = Flag(element, "border", false) ? Color("borderColor") : null;
        return basic with { Minimum = values[0], Maximum = values[1], Gradient = Flag(element, "gradient"), ShowValue = Flag(element, "showValue", basic.ShowValue),
            MinLength = Int((string?)element.Attribute("minLength") ?? "10", 100), MaxLength = Int((string?)element.Attribute("maxLength") ?? "90", 100), Axis = axis, Direction = direction,
            NegativeColor = Flag(element, "negativeBarColorSameAsPositive", false) ? basic.Color : Color("negativeFillColor"),
            BorderColor = border, NegativeBorderColor = border is null ? null : Flag(element, "negativeBarBorderColorSameAsPositive") ? border : Color("negativeBorderColor"),
            AxisColor = axis == DataBarAxis.None ? "FF000000" : Color("axisColor") };
    }
    private static XElement WriteThreshold(ConditionalThreshold value, bool extended = false)
    {
        var kind = value.Kind;
        if (!extended) kind = kind switch { ConditionalThresholdKind.AutoMinimum => ConditionalThresholdKind.Minimum, ConditionalThresholdKind.AutoMaximum => ConditionalThresholdKind.Maximum, _ => kind };
        var element = new XElement((extended ? X14 : Main) + "cfvo", new XAttribute("type", ThresholdTypes[(int)kind]), !value.Inclusive ? new XAttribute("gte", 0) : null);
        if (kind is ConditionalThresholdKind.Number or ConditionalThresholdKind.Percent or ConditionalThresholdKind.Percentile or ConditionalThresholdKind.Formula)
        {
            var text = value.Formula is { } f ? Encode(XlsxStructuredFormula.Write(f).TrimStart('=')) : value.Value.ToString("R", CultureInfo.InvariantCulture);
            if (extended) element.Add(new XElement(Xm + "f", text)); else element.Add(new XAttribute("val", text));
        }
        return element;
    }
    private static XElement WriteVisualRule(ConditionalFormatRule rule)
    {
        var element = new XElement(Main + "cfRule", new XAttribute("type", Types[(int)rule.Kind]), new XAttribute("priority", rule.Priority), new XAttribute("stopIfTrue", rule.StopIfTrue ? 1 : 0));
        if (rule.ColorScale is { } scale) element.Add(new XElement(Main + "colorScale", scale.Thresholds.Select(v => WriteThreshold(v)), scale.Colors.Select(c => new XElement(Main + "color", new XAttribute("rgb", c)))));
        if (rule.IconSet is { } icons) element.Add(new XElement(Main + "iconSet", new XAttribute("iconSet", icons.Name), new XAttribute("showValue", icons.ShowValue ? 1 : 0), new XAttribute("reverse", icons.Reverse ? 1 : 0), icons.Thresholds.Select(v => WriteThreshold(v))));
        if (rule.DataBar is { } bar)
        {
            element.Add(new XElement(Main + "dataBar", new XAttribute("minLength", bar.MinLength == 0 && bar.MaxLength == 100 ? 10 : bar.MinLength), new XAttribute("maxLength", bar.MinLength == 0 && bar.MaxLength == 100 ? 90 : bar.MaxLength), new XAttribute("showValue", bar.ShowValue ? 1 : 0),
                WriteThreshold(bar.Minimum), WriteThreshold(bar.Maximum), new XElement(Main + "color", new XAttribute("rgb", bar.Color))));
            // Keep the same qualified id representation as Excel's native data bar links.
            element.Add(new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", BarLinkExtension), new XAttribute(XNamespace.Xmlns + "x14", X14),
                new XElement(X14 + "id", rule.ScriptId.ToString("B").ToUpperInvariant()))));
        }
        return element;
    }
    public static void WriteExtensions(WorksheetSnapshot sheet, XElement root)
    {
        var bars = sheet.Data.ConditionalFormats.Where(r => r.DataBar is not null).ToArray(); if (bars.Length == 0) return;
        var list = root.Element(Main + "extLst"); if (list is null) root.Add(list = new(Main + "extLst"));
        list.Add(new XElement(Main + "ext", new XAttribute("uri", VisualExtension), new XAttribute(XNamespace.Xmlns + "x14", X14), new XAttribute(XNamespace.Xmlns + "xm", Xm),
            new XElement(X14 + "conditionalFormattings", bars.Select(rule =>
            {
                var bar = rule.DataBar!;
                XElement Color(string name, string color) => new(X14 + name, new XAttribute("rgb", color));
                return new XElement(X14 + "conditionalFormatting",
                    new XElement(X14 + "cfRule", new XAttribute("type", "dataBar"), new XAttribute("id", rule.ScriptId.ToString("B").ToUpperInvariant()),
                        new XElement(X14 + "dataBar", new XAttribute("minLength", bar.MinLength), new XAttribute("maxLength", bar.MaxLength), new XAttribute("showValue", bar.ShowValue ? 1 : 0), new XAttribute("gradient", bar.Gradient ? 1 : 0),
                            new XAttribute("border", bar.BorderColor is null ? 0 : 1), new XAttribute("direction", bar.Direction switch { DataBarDirection.LeftToRight => "leftToRight", DataBarDirection.RightToLeft => "rightToLeft", _ => "context" }),
                            new XAttribute("axisPosition", bar.Axis switch { DataBarAxis.Middle => "middle", DataBarAxis.None => "none", _ => "automatic" }), new XAttribute("negativeBarColorSameAsPositive", bar.Color == bar.NegativeColor ? 1 : 0),
                            new XAttribute("negativeBarBorderColorSameAsPositive", bar.NegativeBorderColor is null || bar.NegativeBorderColor == bar.BorderColor ? 1 : 0),
                            WriteThreshold(bar.Minimum, true), WriteThreshold(bar.Maximum, true),
                            bar.BorderColor is { } border ? Color("borderColor", border) : null, bar.Color != bar.NegativeColor ? Color("negativeFillColor", bar.NegativeColor) : null,
                            bar.NegativeBorderColor is { } negative && negative != bar.BorderColor ? Color("negativeBorderColor", negative) : null, bar.Axis != DataBarAxis.None ? Color("axisColor", bar.AxisColor) : null)),
                    new XElement(Xm + "sqref", string.Join(' ', rule.Ranges)));
            }))));
    }
    public static XElement Shape(XElement root)
    {
        var copy = new XElement(root);
        copy.Elements(Main + "extLst").Elements(Main + "ext").Where(e => string.Equals((string?)e.Attribute("uri"), VisualExtension, StringComparison.OrdinalIgnoreCase)).Remove();
        copy.Elements(Main + "extLst").Where(e => !e.HasElements && !e.Attributes().Any(a => !a.IsNamespaceDeclaration)).Remove(); return copy;
    }
}
