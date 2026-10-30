using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal sealed record XlsxTextValue(string Text, RichCellText? Rich = null);

internal sealed class XlsxCellText
{
    private readonly XlsxPackage _package;
    private readonly XlsxColors _colors;
    public XlsxCellText(XlsxPackage package, string workbook)
    {
        _package = package;
        var styles = package.RelatedByType(workbook, "styles");
        _colors = new(package, workbook, styles is null ? new XElement(Main + "styleSheet") : package.XmlPart(styles).Root!, styles ?? workbook);
    }
    public XlsxTextValue Read(XElement root, string path)
    {
        _package.Shape(root, "t r rPh phoneticPr", "", path);
        if (root.Elements(Main + "rPh").Any() || root.Element(Main + "phoneticPr") is not null)
            _package.Loss($"{path}: 拼音文本未建模", "text/phonetic", path, root);
        string Text(XElement t)
        {
            if (t.HasElements || t.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != XNamespace.Xml + "space"))
                _package.Loss($"{path}: 文本附加属性未建模", "text/attributes", path, t);
            return Decode(t.Value);
        }
        if (!root.Elements(Main + "r").Any())
        {
            if (root.Elements(Main + "t").Count() > 1) throw new InvalidDataException("文本包含重复内容。");
            var value = root.Element(Main + "t") is { } t ? Text(t) : "";
            WorkbookLimits.ValidateText(value); return new(value);
        }
        if (root.Element(Main + "t") is not null) throw new InvalidDataException("文本同时包含普通内容和富文本分段。");
        var runs = new List<TextRun>();
        foreach (var run in root.Elements(Main + "r"))
        {
            _package.Budget.Parse(192);
            if (runs.Count >= 32767) throw new InvalidDataException("富文本分段数量过多。");
            _package.Shape(run, "rPr t", "", path);
            if (run.Elements(Main + "t").Count() != 1 || run.Elements(Main + "rPr").Count() > 1) throw new InvalidDataException("富文本分段结构无效。");
            TextRunFont? font = null;
            if (run.Element(Main + "rPr") is { } props)
            {
                _package.Shape(props, "rFont charset family b i strike color sz u vertAlign scheme", "", path);
                if (props.Elements().GroupBy(e => e.Name).Any(g => g.Count() > 1)) throw new InvalidDataException("重复富文本字体属性。");
                foreach (var p in props.Elements().Where(p => p.Name != Main + "color")) _package.Shape(p, "", "val", path);
                foreach (var propertyName in new[] { "rFont", "charset", "family", "sz", "vertAlign", "scheme" })
                    if (props.Element(Main + propertyName) is { } required) _ = Attr(required, "val");
                string? Val(string name) => (string?)props.Element(Main + name)?.Attribute("val");
                bool? Flag(string name) => props.Element(Main + name) is { } e ? Bool((string?)e.Attribute("val") ?? "1") : null;
                var name = Val("rFont"); var scheme = Val("scheme"); var resolved = _colors.Font(scheme == "none" ? null : scheme, name);
                var color = props.Element(Main + "color");
                if (color is not null)
                {
                    _package.Shape(color, "", "rgb theme indexed auto tint", path);
                    if (new[] { "rgb", "theme", "indexed", "auto" }.Count(a => color.Attribute(a) is not null) > 1) throw new InvalidDataException("富文本颜色存在冲突的来源。");
                }
                font = new(name, OptionalDouble(props.Element(Main + "sz")?.Attribute("val")), Flag("b"), Flag("i"),
                    (string?)color?.Attribute("indexed") == "81" ? null : _colors.Read(color, _package, path), color?.Attribute("theme") is { } theme ? Int(theme.Value, 11) : null,
                    color?.Attribute("indexed") is { } indexed ? Int(indexed.Value, 81) : null,
                    color?.Attribute("auto") is { } auto ? Bool(auto.Value) : null, OptionalDouble(color?.Attribute("tint")),
                    scheme, resolved.EastAsian, resolved.Latin,
                    props.Element(Main + "u") is null ? null : Val("u") ?? "single", Flag("strike"), Val("vertAlign"),
                    Val("family") is { } family ? Int(family, 5) : null, Val("charset") is { } charset ? Int(charset, 255) : null, (string?)color?.Attribute("rgb"));
            }
            runs.Add(new(Text(run.Element(Main + "t")!), font));
        }
        var rich = new RichCellText(runs); return new(rich.Text, rich);
    }

    public static XElement Write(XName name, RichCellText rich) => new(name, rich.Runs.Select(run =>
        new XElement(Main + "r", run.Font is { } font ? WriteFont(font) : null, Text(run.Text))));
    public static XElement Text(string value) => new(Main + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Encode(value));
    private static XElement WriteFont(TextRunFont font)
    {
        var result = new XElement(Main + "rPr");
        void Val(string name, object? value) { if (value is not null) result.Add(new XElement(Main + name, new XAttribute("val", value is bool b ? b ? "1" : "0" : value))); }
        Val("rFont", font.Name); Val("charset", font.Charset); Val("family", font.Family); Val("b", font.Bold); Val("i", font.Italic); Val("strike", font.Strike);
        var color = new XElement(Main + "color");
        if (font.ThemeColor is { } theme) color.SetAttributeValue("theme", theme);
        else if (font.IndexedColor is { } indexed) color.SetAttributeValue("indexed", indexed);
        else if (font.AutoColor is { } auto) color.SetAttributeValue("auto", auto ? 1 : 0);
        else if ((font.RawRgb ?? font.Color) is { } rgb) color.SetAttributeValue("rgb", rgb);
        if (font.Tint is { } tint) color.SetAttributeValue("tint", tint);
        if (color.HasAttributes) result.Add(color);
        Val("sz", font.Size); Val("u", font.Underline); Val("vertAlign", font.Vertical); Val("scheme", font.Scheme);
        return result;
    }
}
