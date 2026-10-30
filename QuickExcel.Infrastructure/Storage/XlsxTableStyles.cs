using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static class XlsxTableStyles
{
    private static readonly Lazy<XElement> Builtins = new(() => XElement.Load(typeof(XlsxTableStyles).Assembly.GetManifestResourceStream("QuickExcel.Infrastructure.Storage.BuiltinTableStyles.xml")!));
    private static readonly string[] Borders = ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair"];
    private static readonly HashSet<string> Roles = ["wholeTable", "headerRow", "totalRow", "firstColumn", "lastColumn", "firstRowStripe", "secondRowStripe",
        "firstColumnStripe", "secondColumnStripe", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell"];
    public static ImmutableArray<TableStyleElement> Resolve(XlsxPackage package, string workbook, string? name)
    {
        if (name is null) return [];
        var path = package.RelatedByType(workbook, "styles", true)!; var styles = package.XmlPart(path).Root!;
        var colors = new XlsxColors(package, workbook, styles, path, tableDefaults: true);
        if (styles.Element(Main + "tableStyles")?.Elements(Main + "tableStyle").SingleOrDefault(e => (string?)e.Attribute("name") == name) is { } custom)
        {
            XlsxTables.Shape(custom, "tableStyleElement", "name pivot table count");
            return custom.Elements().Select(e =>
            {
                XlsxTables.Shape(e, "", "type size dxfId");
                var result = ReadDxf(Dxf(styles, Attr(e, "dxfId")), Attr(e, "type"), e.Attribute("size") is { } size ? Int(size.Value, 9) : 1, colors, package, path);
                if (result.NumberFormatId.HasValue) throw new NotSupportedException("表区域差异数字格式暂不支持，列数字格式可保留。");
                return result;
            }).ToImmutableArray();
        }
        var builtin = Builtins.Value.Elements("style").SingleOrDefault(e => (string?)e.Attribute("name") == name)
            ?? throw new NotSupportedException("未知表样式：" + name);
        return builtin.Elements().Select(e =>
        {
            var dxf = new XElement(e.Element(Main + "dxf")!);
            // Duplicating a built-in style serializes its percentage as a signed 16-bit fraction.
            // Excel renders the built-in's original percentage; a custom style retains the raw tint.
            foreach (var tint in dxf.Descendants().Attributes("tint"))
                tint.Value = Math.Round(double.Parse(tint.Value, System.Globalization.CultureInfo.InvariantCulture), 2).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            return ReadDxf(dxf, Attr(e, "type"), Int(Attr(e, "size"), 9), colors, package, path);
        }).ToImmutableArray();
    }
    private static XElement Dxf(XElement styles, string index)
    {
        var dxfs = styles.Element(Main + "dxfs")?.Elements(Main + "dxf").ToArray() ?? [];
        return dxfs[Int(index, dxfs.Length - 1)];
    }
    public static TableStyleElement? ColumnStyle(XlsxPackage package, string workbook, XAttribute? id)
    {
        if (id is null) return null;
        var path = package.RelatedByType(workbook, "styles", true)!; var styles = package.XmlPart(path).Root!;
        return ReadDxf(Dxf(styles, id.Value), "wholeTable", 1, new(package, workbook, styles, path, tableDefaults: true), package, path);
    }
    private static TableStyleElement ReadDxf(XElement dxf, string role, int size, XlsxColors colors, XlsxPackage package, string path)
    {
        if (!Roles.Contains(role) || size < 1) throw new NotSupportedException("表样式元素类型或条纹大小未支持。");
        XlsxTables.Shape(dxf, "font fill border numFmt", "");
        var font = dxf.Element(Main + "font"); if (font is not null) XlsxTables.Shape(font, "b color", "");
        var bold = font?.Element(Main + "b"); if (bold is not null) XlsxTables.Shape(bold, "", "val");
        var fill = dxf.Element(Main + "fill"); var pattern = fill?.Element(Main + "patternFill");
        if (fill is not null) { XlsxTables.Shape(fill, "patternFill", ""); if (pattern is null) throw new NotSupportedException("表填充样式缺少图案。"); }
        if (pattern is not null)
        {
            XlsxTables.Shape(pattern, "fgColor bgColor", "patternType");
            if ((string?)pattern.Attribute("patternType") != "solid") throw new NotSupportedException("表样式仅支持纯色填充。");
        }
        var border = dxf.Element(Main + "border"); if (border is not null) XlsxTables.Shape(border, "left right top bottom horizontal vertical", "");
        TableBorder? Edge(string edge)
        {
            if (border?.Element(Main + edge) is not { } element) return null;
            XlsxTables.Shape(element, "color", "style"); var index = Array.IndexOf(Borders, (string?)element.Attribute("style") ?? "none");
            if (index < 0) throw new NotSupportedException("表样式边框类型未支持。");
            return new(index, colors.Read(element.Element(Main + "color"), package, path, tableTint: true));
        }
        var number = dxf.Element(Main + "numFmt"); if (number is not null) XlsxTables.Shape(number, "", "numFmtId formatCode");
        return new(role, size, bold is null ? null : bold.Attribute("val") is null || Bool((string?)bold.Attribute("val")),
            colors.Read(font?.Element(Main + "color"), package, path, tableTint: true), colors.Read(pattern?.Element(Main + "fgColor"), package, path, tableTint: true),
            Edge("left"), Edge("right"), Edge("top"), Edge("bottom"), Edge("horizontal"), Edge("vertical"),
            number is null ? null : Int(Attr(number, "numFmtId"), 65535) is var id && id < 164 ? id : 0,
            number is null ? null : Decode(Attr(number, "formatCode")));
    }
    internal static int WriteDxf(TableStyleElement style, XElement styles)
    {
        var dxf = new XElement(Main + "dxf");
        if (style.Bold.HasValue || style.FontColor is not null) dxf.Add(new XElement(Main + "font", style.Bold.HasValue ? new XElement(Main + "b", new XAttribute("val", style.Bold.Value ? 1 : 0)) : null,
            style.FontColor is { } color ? new XElement(Main + "color", new XAttribute("rgb", color)) : null));
        if (style.NumberFormatId.HasValue)
        {
            var formatId = style.NumberFormatId.Value; var code = style.NumberFormatCode ?? "General";
            if (formatId == 0 && code != "General")
            {
                var formats = styles.Descendants(Main + "numFmt").Where(f => (int?)f.Attribute("numFmtId") >= 164).ToArray();
                formatId = formats.FirstOrDefault(f => (string?)f.Attribute("formatCode") == Encode(code)) is { } same ? (int)same.Attribute("numFmtId")! :
                    formats.Select(f => (int)f.Attribute("numFmtId")!).DefaultIfEmpty(163).Max() + 1;
                if (formatId > 65535) throw new InvalidOperationException("差异数字格式编号已满。");
            }
            dxf.Add(new XElement(Main + "numFmt", new XAttribute("numFmtId", formatId), new XAttribute("formatCode", Encode(code))));
        }
        if (style.FillColor is { } fill) dxf.Add(new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "solid"), new XElement(Main + "fgColor", new XAttribute("rgb", fill)))));
        var border = new XElement(Main + "border");
        foreach (var (name, edge) in new[] { ("left", style.Left), ("right", style.Right), ("top", style.Top), ("bottom", style.Bottom), ("vertical", style.Vertical), ("horizontal", style.Horizontal) })
            if (edge is not null) border.Add(new XElement(Main + name, new XAttribute("style", Borders[edge.Style]), edge.Color is null ? null : new XElement(Main + "color", new XAttribute("rgb", edge.Color))));
        if (border.HasElements) dxf.Add(border);
        var dxfs = styles.Element(Main + "dxfs");
        if (dxfs is null)
        {
            dxfs = new(Main + "dxfs"); var after = styles.Elements().FirstOrDefault(e => e.Name.LocalName is "tableStyles" or "colors" or "extLst");
            if (after is null) styles.Add(dxfs); else after.AddBeforeSelf(dxfs);
        }
        var all = dxfs.Elements().ToArray(); var id = Array.FindIndex(all, e => XNode.DeepEquals(e, dxf));
        if (id >= 0) return id;
        if (all.Length >= 4096) throw new InvalidOperationException("差异样式超过 4096。");
        dxfs.Add(dxf); dxfs.SetAttributeValue("count", all.Length + 1); return all.Length;
    }
    public static void WriteCustom(TableStyleInfo style, XElement styles)
    {
        if (style.Name is null || Builtins.Value.Elements().Any(e => (string?)e.Attribute("name") == style.Name)) return;
        var tableStyles = styles.Element(Main + "tableStyles");
        if (tableStyles is null)
        {
            tableStyles = new(Main + "tableStyles"); var after = styles.Elements().FirstOrDefault(e => e.Name.LocalName is "colors" or "extLst");
            if (after is null) styles.Add(tableStyles); else after.AddBeforeSelf(tableStyles);
        }
        if (tableStyles.Elements().Any(e => (string?)e.Attribute("name") == style.Name)) return; // Original definition and its dxf indexes are retained.
        tableStyles.Add(new XElement(Main + "tableStyle", new XAttribute("name", style.Name), new XAttribute("pivot", 0), new XAttribute("count", style.Elements.Length),
            style.Elements.Select(e => new XElement(Main + "tableStyleElement", new XAttribute("type", e.Type), new XAttribute("size", e.Size), new XAttribute("dxfId", WriteDxf(e, styles))))));
        tableStyles.SetAttributeValue("count", tableStyles.Elements().Count());
    }
}
