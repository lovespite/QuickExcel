using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Core;
using QuickExcel.Formulas;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static partial class XlsxConditionalFormatting
{
    private static readonly string[] Types = ["cellIs", "expression", "duplicateValues", "uniqueValues", "containsText", "notContainsText", "beginsWith", "endsWith", "colorScale", "dataBar", "iconSet", "timePeriod"];
    private static readonly string[] Operators = ["between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"];
    private static readonly string[] TextOperators = ["containsText", "notContains", "beginsWith", "endsWith"];
    public static ImmutableArray<ConditionalFormatRule> Read(XlsxPackage package, string workbook, string path, XElement sheet)
    {
        var groups = sheet.Elements(Main + "conditionalFormatting").ToArray();
        var extensions = sheet.Descendants().Where(e => e.Name.Namespace != Main && e.Name.LocalName is "conditionalFormattings" or "conditionalFormatting").ToArray();
        if (groups.Length == 0 && extensions.Length == 0) return [];
        try
        {
            var extendedBars = ReadExtendedBars(sheet); var usedBars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<ConditionalFormatRule>();
            var stylePath = package.RelatedByType(workbook, "styles", true)!; var styles = package.XmlPart(stylePath).Root!;
            var dxfs = styles.Element(Main + "dxfs")?.Elements(Main + "dxf").ToArray() ?? [];
            var colors = new XlsxColors(package, workbook, styles, stylePath);
            foreach (var group in groups)
            {
                XlsxTables.Shape(group, "cfRule", "sqref");
                var ranges = Attr(group, "sqref").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse).ToArray();
                foreach (var rule in group.Elements())
                {
                    XlsxTables.Shape(rule, "formula colorScale dataBar iconSet extLst", "type dxfId priority stopIfTrue operator text timePeriod");
                    var type = Array.IndexOf(Types, Attr(rule, "type")); if (type < 0) throw new NotSupportedException("此条件格式类型尚未支持。");
                    var op = (string?)rule.Attribute("operator"); var index = Array.IndexOf(Operators, op ?? "between");
                    if (type is >= 4 and <= 7) { if (op is not null && op != TextOperators[type - 4]) throw new NotSupportedException("文本条件类型和比较符不一致。"); index = 0; }
                    if (index < 0) throw new NotSupportedException("条件格式比较符未支持。");
                    var formulas = rule.Elements(Main + "formula").Select(f => { XlsxTables.Shape(f, "", ""); return XlsxStructuredFormula.Read("=" + Decode(f.Value).TrimStart('=')); }).ToArray();
                    if (formulas.Length > 2) throw new NotSupportedException("条件格式公式数量超出基础规则。");
                    ConditionalDatePeriod? datePeriod = null;
                    if (type == 11)
                    {
                        var name = (string?)rule.Attribute("timePeriod");
                        datePeriod = Enum.GetValues<ConditionalDatePeriod>().Cast<ConditionalDatePeriod?>().SingleOrDefault(p => ConditionalDates.XmlName(p!.Value) == name);
                        if (datePeriod is null || op is not null || rule.Attribute("text") is not null || formulas.Length > 1 || formulas.Length == 1 &&
                            FormulaSyntax.Print(FormulaSyntax.Parse(formulas[0])) != FormulaSyntax.Print(FormulaSyntax.Parse(ConditionalDates.Formula(datePeriod.Value, new(ranges[0].Top, ranges[0].Left)))))
                            throw new NotSupportedException(MessageText.Get("发生日期规则包含未知日期范围或非标准公式，需保留原规则。"));
                        formulas = [];
                    }
                    else if (rule.Attribute("timePeriod") is not null) throw new NotSupportedException(MessageText.Get("发生日期规则包含未知日期范围或非标准公式，需保留原规则。"));
                    var dxf = rule.Attribute("dxfId") is { } id ? ReadStyle(dxfs[Int(id.Value, dxfs.Length - 1)], colors, package, stylePath) : new DifferentialStyle();
                    var visual = rule.Elements().Where(e => e.Name.LocalName is "colorScale" or "dataBar" or "iconSet").ToArray();
                    if (type is >= 8 and <= 10 ? visual.Length != 1 || visual[0].Name.LocalName != Types[type] : visual.Length != 0)
                        throw new NotSupportedException("条件格式视觉类型和定义不一致。");
                    ConditionalColorScale? scale = null; ConditionalDataBar? bar = null; ConditionalIconSet? icons = null;
                    if (type == 8) scale = ReadScale(visual[0], colors, package, stylePath);
                    if (type == 9) bar = ReadBar(visual[0], colors, package, stylePath);
                    if (type == 10) icons = ReadIcons(visual[0]);
                    if (rule.Element(Main + "extLst") is { } ext)
                    {
                        var link = ReadBarLink(ext);
                        if (bar is null || !extendedBars.TryGetValue(link, out var extended) || !usedBars.Add(link) || !ranges.SequenceEqual(extended.Ranges))
                            throw new NotSupportedException("数据条扩展缺失、重复或应用区域不匹配。");
                        bar = ReadExtendedBar(extended.Bar, bar, colors, package, stylePath);
                    }
                    result.Add(new(ranges, (ConditionalFormatKind)type, Int(Attr(rule, "priority"), int.MaxValue), dxf,
                        formulas.ElementAtOrDefault(0), formulas.ElementAtOrDefault(1), (ConditionalComparison)index,
                        Bool((string?)rule.Attribute("stopIfTrue")), rule.Attribute("text") is { } text ? Decode(text.Value) : null, scale, bar, icons, datePeriod));
                    if (result.Count > 1024) throw new InvalidDataException("条件格式规则超过 1024。");
                }
            }
            if (usedBars.Count != extendedBars.Count) throw new NotSupportedException("独立或未关联的扩展条件格式尚未支持。");
            var ordered = result.OrderBy(r => r.Priority).ToImmutableArray();
            _ = new WorksheetSnapshot(SheetId.New(), "Conditional", data: new() { ConditionalFormats = ordered });
            return ordered;
        }
        catch (Exception ex) when (ex is NotSupportedException or FormatException or ArgumentException)
        {
            // A missing high-priority rule may suppress lower rules. Preserve the whole collection.
            package.Loss("条件格式超出可解释范围，整组规则随原包保留：" + ex.Message, "worksheet/conditionalFormatting", path);
            return [];
        }
    }
    private static DifferentialStyle ReadStyle(XElement dxf, XlsxColors colors, XlsxPackage package, string path)
    {
        XlsxTables.Shape(dxf, "font fill border numFmt", "");
        var font = dxf.Element(Main + "font");
        if (font is not null) XlsxTables.Shape(font, "b i strike u color condense extend", "");
        bool? Flag(string name)
        {
            var e = font?.Element(Main + name); if (e is null) return null;
            XlsxTables.Shape(e, "", "val"); return e.Attribute("val") is null || Bool((string?)e.Attribute("val"));
        }
        if (Flag("condense") == true || Flag("extend") == true) throw new NotSupportedException("压缩/扩展字体尚未支持。");
        var underline = font?.Element(Main + "u"); if (underline is not null) XlsxTables.Shape(underline, "", "val");
        var fill = dxf.Element(Main + "fill"); var pattern = fill?.Element(Main + "patternFill");
        if (fill is not null) XlsxTables.Shape(fill, "patternFill", "");
        string? fillColor = null;
        if (pattern is not null)
        {
            XlsxTables.Shape(pattern, "fgColor bgColor", "patternType"); var type = (string?)pattern.Attribute("patternType");
            if (type is not (null or "solid")) throw new NotSupportedException("条件格式只支持纯色填充。");
            fillColor = colors.Read(type == "solid" ? pattern.Element(Main + "fgColor") ?? pattern.Element(Main + "bgColor") :
                pattern.Element(Main + "bgColor") ?? pattern.Element(Main + "fgColor"), package, path);
            if (fillColor is null) throw new NotSupportedException("条件格式填充未提供明确颜色。");
        }
        var border = dxf.Element(Main + "border"); if (border is not null) XlsxTables.Shape(border, "left right top bottom", "");
        TableBorder? Edge(string name)
        {
            var e = border?.Element(Main + name); if (e is null) return null;
            XlsxTables.Shape(e, "color", "style"); var index = Array.IndexOf(XlsxTableStyles.Borders, (string?)e.Attribute("style") ?? "none");
            if (index < 0) throw new NotSupportedException("条件格式边框线型未支持。");
            return new(index, colors.Read(e.Element(Main + "color"), package, path));
        }
        var number = dxf.Element(Main + "numFmt"); if (number is not null) XlsxTables.Shape(number, "", "numFmtId formatCode");
        return new(Flag("b"), Flag("i"), Flag("strike"), underline is null ? null : (string?)underline.Attribute("val") ?? "single",
            colors.Read(font?.Element(Main + "color"), package, path), fillColor, Edge("left"), Edge("right"), Edge("top"), Edge("bottom"),
            number is null ? null : Int(Attr(number, "numFmtId"), 65535) is var id && id < 164 ? id : 0, number is null ? null : Decode(Attr(number, "formatCode")));
    }
    private static int WriteStyle(DifferentialStyle style, XElement styles)
    {
        // Reuse table dxf numeric-format allocation, then add the conditional font properties.
        var basic = new TableStyleElement("wholeTable", Bold: style.Bold, FontColor: style.FontColor, FillColor: style.FillColor,
            Left: style.Left, Right: style.Right, Top: style.Top, Bottom: style.Bottom, NumberFormatId: style.NumberFormatId, NumberFormatCode: style.NumberFormatCode);
        var dxf = XlsxTableStyles.CreateDxf(basic, styles);
        // Excel's conditional dxf uses bgColor without patternType for a solid interior.
        // The ordinary/table solid fgColor encoding passes schema validation but paints white in Excel CF.
        if (style.FillColor is { } fillColor)
            dxf.Element(Main + "fill")!.ReplaceNodes(new XElement(Main + "patternFill", new XElement(Main + "bgColor", new XAttribute("rgb", fillColor))));
        if (style.Italic.HasValue || style.Strike.HasValue || style.Underline is not null)
        {
            var font = dxf.Element(Main + "font"); if (font is null) { font = new(Main + "font"); dxf.AddFirst(font); }
            if (style.Italic.HasValue) font.Add(new XElement(Main + "i", new XAttribute("val", style.Italic.Value ? 1 : 0)));
            if (style.Strike.HasValue) font.Add(new XElement(Main + "strike", new XAttribute("val", style.Strike.Value ? 1 : 0)));
            if (style.Underline is not null) font.Add(new XElement(Main + "u", new XAttribute("val", style.Underline)));
            // CT_Font in a dxf uses a sequence (unlike ordinary workbook fonts).
            var order = new[] { "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme" };
            var children = font.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName)).ToArray();
            font.ReplaceNodes(children);
        }
        return XlsxTableStyles.InternDxf(dxf, styles);
    }
    public static IEnumerable<XElement> Write(WorksheetSnapshot sheet, XElement styles)
    {
        foreach (var rule in sheet.Data.ConditionalFormats.OrderBy(r => r.Priority))
        {
            if (rule.IsVisual)
            {
                yield return new(Main + "conditionalFormatting", new XAttribute("sqref", string.Join(" ", rule.Ranges)), WriteVisualRule(rule));
                continue;
            }
            var element = new XElement(Main + "cfRule", new XAttribute("type", Types[(int)rule.Kind]), new XAttribute("priority", rule.Priority),
                new XAttribute("stopIfTrue", rule.StopIfTrue ? 1 : 0), new XAttribute("dxfId", WriteStyle(rule.Style, styles)),
                rule.Kind == ConditionalFormatKind.CellIs ? new XAttribute("operator", Operators[(int)rule.Comparison]) :
                    (int)rule.Kind is >= 4 and <= 7 ? new XAttribute("operator", TextOperators[(int)rule.Kind - 4]) : null,
                rule.DatePeriod is { } period ? new XAttribute("timePeriod", ConditionalDates.XmlName(period)) : null,
                rule.Text is { } text ? new XAttribute("text", Encode(text)) : null,
                new[] { rule.Formula1, rule.Formula2 }.OfType<string>().Select(f => new XElement(Main + "formula", Encode(XlsxStructuredFormula.Write(f).TrimStart('=')))));
            if (rule.DatePeriod is { } datePeriod) element.Add(new XElement(Main + "formula", ConditionalDates.Formula(datePeriod, rule.Anchor).TrimStart('=')));
            yield return new(Main + "conditionalFormatting", new XAttribute("sqref", string.Join(" ", rule.Ranges)), element);
        }
    }
}
