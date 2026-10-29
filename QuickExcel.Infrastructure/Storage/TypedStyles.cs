using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static class TypedStyles
{
    private static readonly string[] Horizontal = ["general", "left", "center", "right", "justify"];
    private static readonly string[] Vertical = ["top", "center", "bottom", "justify"];
    private static readonly string[] Borders = ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair"];

    public static List<StyleDefinition> Read(XlsxPackage package, string workbookPart)
    {
        var path = package.RelatedByType(workbookPart, "styles");
        if (path is null) return [new()];
        package.RequireType(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        var root = package.XmlPart(path).Root!;
        if (root.Name != Main + "styleSheet") throw new InvalidDataException("无效样式表。");
        package.Shape(root, "numFmts fonts fills borders cellXfs cellStyleXfs cellStyles colors", "", path);
        var colors = new XlsxColors(package, workbookPart, root, path);
        // Named/inherited styles need a separate editing model. Keep them in the compatibility report.
        if (root.Element(Main + "cellStyleXfs") is { } inherited && (inherited.Elements().Count() > 1 ||
            inherited.Elements().Any(x => x.Attributes().Any(a => a.Name.LocalName is not ("numFmtId" or "fontId" or "fillId" or "borderId") || a.Value != "0"))))
            package.Issues.Add($"{path}: 命名/继承样式未建模");
        if (root.Element(Main + "cellStyles") is { } named && named.Elements().Any(x => (string?)x.Attribute("builtinId") != "0"))
            package.Issues.Add($"{path}: 自定义命名样式未建模");
        var formats = new Dictionary<int, string>();
        foreach (var item in root.Element(Main + "numFmts")?.Elements(Main + "numFmt") ?? [])
        {
            var id = Int(Attr(item, "numFmtId"), 65535);
            if (!formats.TryAdd(id, Decode(Attr(item, "formatCode")))) throw new InvalidDataException("重复数字格式。");
        }
        var fonts = root.Element(Main + "fonts")?.Elements(Main + "font").ToArray() ?? [new XElement(Main + "font")];
        var fills = root.Element(Main + "fills")?.Elements(Main + "fill").ToArray() ?? [new XElement(Main + "fill")];
        var borders = root.Element(Main + "borders")?.Elements(Main + "border").ToArray() ?? [new XElement(Main + "border")];
        var result = new List<StyleDefinition>();
        foreach (var xf in root.Element(Main + "cellXfs")?.Elements(Main + "xf") ?? [])
        {
            if (result.Count >= 4096) throw new InvalidDataException("样式数量超过 4096。");
            package.Shape(xf, "alignment", "numFmtId fontId fillId borderId xfId applyFont applyFill applyBorder applyAlignment applyNumberFormat", path);
            if ((string?)xf.Attribute("xfId") is not (null or "0") || xf.Attributes().Any(a => a.Name.LocalName.StartsWith("apply", StringComparison.Ordinal) && a.Value is "0" or "false"))
                package.Issues.Add($"{path}: 继承样式或显式禁用的样式分量未建模");
            var font = fonts[Int((string?)xf.Attribute("fontId") ?? "0", fonts.Length - 1)];
            var fill = fills[Int((string?)xf.Attribute("fillId") ?? "0", fills.Length - 1)];
            var border = borders[Int((string?)xf.Attribute("borderId") ?? "0", borders.Length - 1)];
            package.Shape(font, "name sz color b i", "", path);
            foreach (var item in font.Elements().Where(e => e.Name != Main + "color")) package.Shape(item, "", "val", path);
            package.Shape(fill, "patternFill", "", path);
            package.Shape(border, "left right top bottom diagonal", "", path);
            foreach (var side in border.Elements()) package.Shape(side, "color", "style", path);
            if (border.Element(Main + "diagonal") is { } diagonal && (diagonal.HasAttributes || diagonal.HasElements))
                package.Issues.Add($"{path}: 对角边框未建模");
            var pattern = fill.Element(Main + "patternFill");
            if (pattern is not null)
            {
                package.Shape(pattern, "fgColor bgColor", "patternType", path);
                if ((string?)pattern.Attribute("patternType") is not (null or "none" or "gray125" or "solid"))
                    package.Issues.Add($"{path}: 非实色填充未建模");
                if (pattern.Element(Main + "bgColor") is { } bg && (string?)bg.Attribute("indexed") != "64")
                    package.Issues.Add($"{path}: 填充背景色未建模");
            }
            var alignment = xf.Element(Main + "alignment");
            if (alignment is not null) package.Shape(alignment, "", "horizontal vertical wrapText", path);
            var numberId = Int((string?)xf.Attribute("numFmtId") ?? "0", 65535);
            if (numberId >= 164 && !formats.ContainsKey(numberId)) throw new InvalidDataException("数字格式定义缺失。");
            int Side(string name) => Named((string?)border.Element(Main + name)?.Attribute("style"), Borders, 0, package, path);
            string? SideColor(string name) => colors.Read(border.Element(Main + name)?.Element(Main + "color"), package, path);
            result.Add(new StyleDefinition
            {
                NumberFormatId = formats.ContainsKey(numberId) ? 0 : numberId,
                NumberFormatCode = formats.GetValueOrDefault(numberId),
                FontName = (string?)font.Element(Main + "name")?.Attribute("val"),
                FontSize = OptionalDouble(font.Element(Main + "sz")?.Attribute("val")),
                FontColor = colors.Read(font.Element(Main + "color"), package, path),
                Bold = Flag(font.Element(Main + "b")), Italic = Flag(font.Element(Main + "i")),
                FillColor = (string?)pattern?.Attribute("patternType") == "solid" ? colors.Read(pattern.Element(Main + "fgColor"), package, path) : null,
                Horizontal = Named((string?)alignment?.Attribute("horizontal"), Horizontal, 0, package, path),
                Vertical = Named((string?)alignment?.Attribute("vertical"), Vertical, 2, package, path),
                WrapText = Bool((string?)alignment?.Attribute("wrapText")),
                LeftBorder = Side("left"), RightBorder = Side("right"), TopBorder = Side("top"), BottomBorder = Side("bottom"),
                LeftBorderColor = SideColor("left"), RightBorderColor = SideColor("right"), TopBorderColor = SideColor("top"), BottomBorderColor = SideColor("bottom")
            });
        }
        return result.Count == 0 ? [new()] : result;
    }

    public static XDocument Write(IReadOnlyList<StyleDefinition> styles)
    {
        if (styles.Count > 4096) throw new InvalidDataException("样式数量超过 4096。");
        var fonts = new XElement(Main + "fonts", new XAttribute("count", styles.Count));
        var fills = new XElement(Main + "fills", new XAttribute("count", styles.Count + 2),
            new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "none"))),
            new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "gray125"))));
        var borders = new XElement(Main + "borders", new XAttribute("count", styles.Count));
        var xfs = new XElement(Main + "cellXfs", new XAttribute("count", styles.Count));
        var formats = new XElement(Main + "numFmts");
        for (var index = 0; index < styles.Count; index++)
        {
            var style = styles[index];
            if (style.NumberFormatId is < 0 or >= 164) throw new InvalidDataException("自定义数字格式必须提供格式代码。");
            var numFmtId = style.NumberFormatCode is null ? style.NumberFormatId : 164 + index;
            if (style.NumberFormatCode is not null) formats.Add(new XElement(Main + "numFmt", new XAttribute("numFmtId", numFmtId), new XAttribute("formatCode", Encode(style.NumberFormatCode))));
            XElement? Val(string tag, object? value) => value is null ? null : new XElement(Main + tag, new XAttribute("val", value));
            XElement? Rgb(string tag, string? value) => value is null ? null : new XElement(Main + tag, new XAttribute("rgb", value));
            fonts.Add(new XElement(Main + "font", Val("name", style.FontName), Val("sz", style.FontSize), Rgb("color", style.FontColor),
                style.Bold ? new XElement(Main + "b") : null, style.Italic ? new XElement(Main + "i") : null));
            fills.Add(new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", style.FillColor is null ? "none" : "solid"), Rgb("fgColor", style.FillColor))));
            XElement Side(string name, int kind, string? color) => new(Main + name, new XAttribute("style", Borders[kind]), Rgb("color", color));
            borders.Add(new XElement(Main + "border", Side("left", style.LeftBorder, style.LeftBorderColor), Side("right", style.RightBorder, style.RightBorderColor),
                Side("top", style.TopBorder, style.TopBorderColor), Side("bottom", style.BottomBorder, style.BottomBorderColor)));
            xfs.Add(new XElement(Main + "xf", new XAttribute("numFmtId", numFmtId), new XAttribute("fontId", index), new XAttribute("fillId", index + 2),
                new XAttribute("borderId", index), new XAttribute("xfId", 0), new XAttribute("applyFont", 1), new XAttribute("applyFill", 1), new XAttribute("applyBorder", 1),
                new XAttribute("applyNumberFormat", 1), new XAttribute("applyAlignment", 1), new XElement(Main + "alignment", new XAttribute("horizontal", Horizontal[style.Horizontal]),
                    new XAttribute("vertical", Vertical[style.Vertical]), new XAttribute("wrapText", style.WrapText ? 1 : 0))));
        }
        formats.SetAttributeValue("count", formats.Elements().Count());
        return new XDocument(new XElement(Main + "styleSheet", formats, fonts, fills, borders,
            new XElement(Main + "cellStyleXfs", new XAttribute("count", 1), new XElement(Main + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))), xfs));
    }

    public static int Int(string text, int maximum)
    {
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value < 0 || value > maximum)
            throw new InvalidDataException($"无效或超限的整数：{text}");
        return value;
    }
    public static double? OptionalDouble(XAttribute? value)
    {
        if (value is null) return null;
        if (!double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) || !double.IsFinite(result) || result < 0)
            throw new InvalidDataException("尺寸无效。");
        return result;
    }
    public static bool Bool(string? value) => value switch
    { null or "0" or "false" => false, "1" or "true" => true, _ => throw new InvalidDataException("无效布尔值。") };
    private static bool Flag(XElement? element) => element is not null && Bool((string?)element.Attribute("val") ?? "1");
    private static int Named(string? value, string[] supported, int fallback, XlsxPackage package, string path)
    {
        if (value is null) return fallback;
        var index = Array.IndexOf(supported, value);
        if (index >= 0) return index;
        package.Issues.Add($"{path}: 未建模样式 {value}");
        return fallback;
    }
    public static string Decode(string value) => System.Text.RegularExpressions.Regex.Replace(value, "_x([0-9A-Fa-f]{4})_",
        match => ((char)int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());
    public static string Encode(string value)
    {
        var escaped = System.Text.RegularExpressions.Regex.Replace(value, "_x[0-9A-Fa-f]{4}_", match => "_x005F_" + match.Value[1..]);
        var result = new System.Text.StringBuilder();
        foreach (var ch in escaped)
            if (ch < 32 && ch is not ('\n' or '\t') || ch is '\uFFFE' or '\uFFFF') result.Append($"_x{(int)ch:X4}_");
            else result.Append(ch);
        return result.ToString();
    }
}
