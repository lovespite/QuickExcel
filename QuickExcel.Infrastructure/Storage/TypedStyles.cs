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
        if (!XlsxStylePreservation.CanPreserveStyles(package, path, root))
            package.Loss($"{path}: 样式含未识别扩展或依赖，无法证明关联可安全保留", "styles/preservation", path);
        var shape = new XElement(root);
        shape.Attributes().Where(a => a.Name == XName.Get("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006")).Remove();
        if (XlsxStylePreservation.CanPreserveStyles(package, path, root)) shape.Element(Main + "extLst")?.Remove();
        package.Shape(shape, "numFmts fonts fills borders cellXfs cellStyleXfs cellStyles colors dxfs tableStyles", "", path);
        var colors = new XlsxColors(package, workbookPart, root, path);
        var baseStyles = root.Element(Main + "cellStyleXfs")?.Elements(Main + "xf").ToArray() ?? [];
        if (baseStyles.Length > 4096) throw new InvalidDataException("命名样式基表超过 4096。");
        var namedStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var named in root.Element(Main + "cellStyles")?.Elements(Main + "cellStyle") ?? [])
        {
            if (!namedStyles.Add(Attr(named, "name")) || namedStyles.Count > 4096) throw new InvalidDataException("命名样式名称重复或数量超限。");
            _ = Int(Attr(named, "xfId"), baseStyles.Length - 1);
            var namedShape = new XElement(named);
            var uid = namedShape.Attribute(XName.Get("uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"));
            if (uid is not null && Guid.TryParse(uid.Value, out _)) uid.Remove();
            package.Shape(namedShape, "", "name xfId builtinId iLevel hidden customBuiltin", path);
        }
        var formats = new Dictionary<int, string>();
        foreach (var item in root.Element(Main + "numFmts")?.Elements(Main + "numFmt") ?? [])
        {
            var id = Int(Attr(item, "numFmtId"), 65535);
            if (!formats.TryAdd(id, Decode(Attr(item, "formatCode")))) throw new InvalidDataException("重复数字格式。");
        }
        var fonts = root.Element(Main + "fonts")?.Elements(Main + "font").ToArray() ?? [new XElement(Main + "font")];
        var fills = root.Element(Main + "fills")?.Elements(Main + "fill").ToArray() ?? [new XElement(Main + "fill")];
        var borders = root.Element(Main + "borders")?.Elements(Main + "border").ToArray() ?? [new XElement(Main + "border")];
        foreach (var parent in baseStyles)
        {
            package.Shape(parent, "alignment protection", "numFmtId fontId fillId borderId applyFont applyFill applyBorder applyAlignment applyNumberFormat applyProtection", path);
            _ = Int((string?)parent.Attribute("fontId") ?? "0", fonts.Length - 1);
            _ = Int((string?)parent.Attribute("fillId") ?? "0", fills.Length - 1);
            _ = Int((string?)parent.Attribute("borderId") ?? "0", borders.Length - 1);
            var numberId = Int((string?)parent.Attribute("numFmtId") ?? "0", 65535);
            if (numberId >= 164 && !formats.ContainsKey(numberId)) throw new InvalidDataException("命名样式数字格式定义缺失。");
            foreach (var flag in parent.Attributes().Where(a => a.Name.LocalName.StartsWith("apply", StringComparison.Ordinal))) _ = Bool(flag.Value);
        }
        var result = new List<StyleDefinition>();
        foreach (var xf in root.Element(Main + "cellXfs")?.Elements(Main + "xf") ?? [])
        {
            if (result.Count >= 4096) throw new InvalidDataException("样式数量超过 4096。");
            package.Shape(xf, "alignment", "numFmtId fontId fillId borderId xfId applyFont applyFill applyBorder applyAlignment applyNumberFormat", path);
            if (xf.Attribute("xfId") is not null && baseStyles.Length == 0) throw new InvalidDataException("单元格引用了缺失的命名样式基表。");
            var parentId = Int((string?)xf.Attribute("xfId") ?? "0", Math.Max(0, baseStyles.Length - 1));
            foreach (var flag in xf.Attributes().Where(a => a.Name.LocalName.StartsWith("apply", StringComparison.Ordinal))) _ = Bool(flag.Value);
            var font = fonts[Int((string?)xf.Attribute("fontId") ?? "0", fonts.Length - 1)];
            var fill = fills[Int((string?)xf.Attribute("fillId") ?? "0", fills.Length - 1)];
            var border = borders[Int((string?)xf.Attribute("borderId") ?? "0", borders.Length - 1)];
            package.Shape(font, "name sz color b i u strike family charset scheme", "", path);
            foreach (var item in font.Elements().Where(e => e.Name != Main + "color")) package.Shape(item, "", "val", path);
            package.Shape(fill, "patternFill", "", path);
            package.Shape(border, "left right top bottom diagonal", "", path);
            foreach (var side in border.Elements()) package.Shape(side, "color", "style", path);
            if (border.Element(Main + "diagonal") is { } diagonal && (diagonal.HasAttributes || diagonal.HasElements))
                package.Loss($"{path}: 对角边框未建模", "styles/diagonalBorder", path, range: "cellXfs[" + result.Count + "]");
            var pattern = fill.Element(Main + "patternFill");
            if (pattern is not null)
            {
                package.Shape(pattern, "fgColor bgColor", "patternType", path);
                if ((string?)pattern.Attribute("patternType") is not (null or "none" or "gray125" or "solid"))
                    package.Loss($"{path}: 非实色填充未建模", "styles/patternFill", path, range: "cellXfs[" + result.Count + "]");
                if (pattern.Element(Main + "bgColor") is { } bg && (string?)bg.Attribute("indexed") != "64")
                    package.Loss($"{path}: 填充背景色未建模", "styles/backgroundColor", path, range: "cellXfs[" + result.Count + "]");
            }
            var alignment = xf.Element(Main + "alignment");
            if (alignment is not null) package.Shape(alignment, "", "horizontal vertical wrapText indent shrinkToFit textRotation", path);
            var rotation = Int((string?)alignment?.Attribute("textRotation") ?? "0", 255);
            if (rotation is > 180 and < 255) throw new InvalidDataException("文字旋转值无效。");
            if (rotation == 255) package.Loss($"{path}: 竖排文字暂仅保留", "styles/stackedText", path, range: "cellXfs[" + result.Count + "]");
            var scheme = (string?)font.Element(Main + "scheme")?.Attribute("val");
            if (scheme == "none") scheme = null;
            if (scheme is not (null or "major" or "minor")) throw new InvalidDataException("主题字体方案无效。");
            var resolvedFont = colors.Font(scheme, (string?)font.Element(Main + "name")?.Attribute("val"));
            var numberId = Int((string?)xf.Attribute("numFmtId") ?? "0", 65535);
            if (numberId >= 164 && !formats.ContainsKey(numberId)) throw new InvalidDataException("数字格式定义缺失。");
            int Side(string name) => Named((string?)border.Element(Main + name)?.Attribute("style"), Borders, 0, package, path);
            string? SideColor(string name) => colors.Read(border.Element(Main + name)?.Element(Main + "color"), package, path);
            result.Add(new StyleDefinition
            {
                NumberFormatId = formats.ContainsKey(numberId) ? 0 : numberId,
                NumberFormatCode = formats.GetValueOrDefault(numberId),
                FontName = resolvedFont.Latin, EastAsianFontName = resolvedFont.EastAsian, FontScheme = scheme, NamedStyleId = parentId,
                FontSize = OptionalDouble(font.Element(Main + "sz")?.Attribute("val")),
                FontColor = colors.Read(font.Element(Main + "color"), package, path),
                Bold = Flag(font.Element(Main + "b")), Italic = Flag(font.Element(Main + "i")),
                Underline = font.Element(Main + "u") is null ? null : (string?)font.Element(Main + "u")?.Attribute("val") ?? "single", Strike = Flag(font.Element(Main + "strike")),
                FillColor = (string?)pattern?.Attribute("patternType") == "solid" ? colors.Read(pattern.Element(Main + "fgColor"), package, path) : null,
                Horizontal = Named((string?)alignment?.Attribute("horizontal"), Horizontal, 0, package, path),
                Vertical = Named((string?)alignment?.Attribute("vertical"), Vertical, 2, package, path),
                WrapText = Bool((string?)alignment?.Attribute("wrapText")),
                TextRotation = rotation, Indent = Int((string?)alignment?.Attribute("indent") ?? "0", 250),
                ShrinkToFit = Bool((string?)alignment?.Attribute("shrinkToFit")),
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
            // CT_Font uses an ordered sequence; our reader accepts any order, but schema validation requires this order.
            fonts.Add(new XElement(Main + "font", style.Bold ? new XElement(Main + "b") : null,
                style.Italic ? new XElement(Main + "i") : null, style.Strike ? new XElement(Main + "strike") : null,
                Val("u", style.Underline), Val("sz", style.FontSize), Rgb("color", style.FontColor), Val("name", style.FontName), Val("scheme", style.FontScheme)));
            fills.Add(new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", style.FillColor is null ? "none" : "solid"), Rgb("fgColor", style.FillColor))));
            XElement Side(string name, int kind, string? color) => new(Main + name, new XAttribute("style", Borders[kind]), Rgb("color", color));
            borders.Add(new XElement(Main + "border", Side("left", style.LeftBorder, style.LeftBorderColor), Side("right", style.RightBorder, style.RightBorderColor),
                Side("top", style.TopBorder, style.TopBorderColor), Side("bottom", style.BottomBorder, style.BottomBorderColor)));
            xfs.Add(new XElement(Main + "xf", new XAttribute("numFmtId", numFmtId), new XAttribute("fontId", index), new XAttribute("fillId", index + 2),
                new XAttribute("borderId", index), new XAttribute("xfId", 0), new XAttribute("applyFont", 1), new XAttribute("applyFill", 1), new XAttribute("applyBorder", 1),
                new XAttribute("applyNumberFormat", 1), new XAttribute("applyAlignment", 1), new XElement(Main + "alignment", new XAttribute("horizontal", Horizontal[style.Horizontal]),
                    new XAttribute("vertical", Vertical[style.Vertical]), new XAttribute("wrapText", style.WrapText ? 1 : 0),
                    style.Indent == 0 ? null : new XAttribute("indent", style.Indent), style.TextRotation == 0 ? null : new XAttribute("textRotation", style.TextRotation),
                    style.ShrinkToFit ? new XAttribute("shrinkToFit", 1) : null)));
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
        package.Loss($"{path}: 未建模样式 {value}", "styles/value", path);
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
