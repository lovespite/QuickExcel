using System.Globalization;
using System.Xml.Linq;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

/// <summary>Resolves SpreadsheetML colors to the model's concrete ARGB values.</summary>
internal sealed class XlsxColors
{
    // SpreadsheetML theme indices use light/dark order, unlike DrawingML's XML order.
    private static readonly string[] ThemeSlots = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    private static readonly string[] DefaultIndexed = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
        "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
        "800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
        "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCCC " +
        "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
        "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 " +
        "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
        "003366 339966 003300 333300 993300 993366 333399 333333").Split(' ');
    private readonly string?[] _theme = new string?[12];
    private readonly string[] _indexed;

    public XlsxColors(XlsxPackage package, string workbookPart, XElement styles, string stylesPath)
    {
        _indexed = DefaultIndexed.Select(rgb => "FF" + rgb).ToArray();
        if (styles.Element(Main + "colors") is { } colors)
        {
            package.Shape(colors, "indexedColors", "", stylesPath);
            if (colors.Element(Main + "indexedColors") is { } indexed)
            {
                package.Shape(indexed, "rgbColor", "", stylesPath);
                var entries = indexed.Elements(Main + "rgbColor").ToArray();
                if (entries.Length > 64) throw new InvalidDataException("索引调色板超过 64 色。");
                for (var i = 0; i < entries.Length; i++)
                {
                    package.Shape(entries[i], "", "rgb", stylesPath);
                    // Indexed palette alpha is reserved, and is commonly written as 00 by Excel.
                    _indexed[i] = "FF" + Rgb(Attr(entries[i], "rgb"), 8)[2..];
                }
            }
        }
        if (package.RelatedByType(workbookPart, "theme") is not { } path) return;
        package.RequireType(path, "application/vnd.openxmlformats-officedocument.theme+xml");
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var root = package.XmlPart(path).Root;
        if (root?.Name != drawing + "theme") throw new InvalidDataException("无效主题部件。");
        var scheme = root.Element(drawing + "themeElements")?.Element(drawing + "clrScheme");
        for (var i = 0; i < ThemeSlots.Length; i++)
        {
            var slot = scheme?.Element(drawing + ThemeSlots[i]);
            var rgb = (string?)slot?.Element(drawing + "srgbClr")?.Attribute("val") ??
                (string?)slot?.Element(drawing + "sysClr")?.Attribute("lastClr");
            if (rgb is not null) _theme[i] = "FF" + Rgb(rgb, 6);
        }
        // Only cell colors are flattened. Theme fonts/effects and theme-editing semantics still need preservation.
        package.Loss($"{path}: 主题颜色已解析；主题字体、效果及主题关联未建模", "theme", path);
    }

    public string? Read(XElement? element, XlsxPackage package, string path)
    {
        if (element is null) return null;
        package.Shape(element, "", "rgb theme indexed auto tint", path);
        var tint = 0d;
        if (element.Attribute("tint") is { } value &&
            (!double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out tint) || !double.IsFinite(tint) || tint is < -1 or > 1))
            throw new InvalidDataException("颜色 tint 必须在 -1 到 1 之间。");
        string? color = null;
        if (element.Attribute("rgb") is { } rgb) color = Rgb(rgb.Value, 8);
        else if (element.Attribute("theme") is { } theme)
        {
            color = _theme[Int(theme.Value, 11)];
            if (color is null) package.Loss($"{path}: 主题颜色 {theme.Value} 缺少可解析的定义", "theme/color", path, element);
        }
        else if (element.Attribute("indexed") is { } indexed)
        {
            var index = Int(indexed.Value, 65);
            color = index switch { 64 => "FF000000", 65 => "FFFFFFFF", _ => _indexed[index] };
        }
        else if (Bool((string?)element.Attribute("auto"))) color = "FF000000";
        return color is null || tint == 0 ? color : Tint(color, tint);
    }

    private static string Rgb(string value, int length)
    {
        if (value.Length != length || value.Any(c => !char.IsAsciiHexDigit(c))) throw new InvalidDataException("颜色值无效。");
        return value;
    }

    // ISO 29500: tint adjusts HLS luminance, not the individual RGB channels.
    private static string Tint(string argb, double tint)
    {
        var r = int.Parse(argb.AsSpan(2, 2), NumberStyles.HexNumber) / 255d;
        var g = int.Parse(argb.AsSpan(4, 2), NumberStyles.HexNumber) / 255d;
        var b = int.Parse(argb.AsSpan(6, 2), NumberStyles.HexNumber) / 255d;
        var max = Math.Max(r, Math.Max(g, b)); var min = Math.Min(r, Math.Min(g, b));
        var luminance = (max + min) / 2;
        var delta = max - min;
        var saturation = delta == 0 ? 0 : delta / (1 - Math.Abs(2 * luminance - 1));
        var hue = delta == 0 ? 0 : max == r ? ((g - b) / delta + 6) % 6 : max == g ? (b - r) / delta + 2 : (r - g) / delta + 4;
        luminance = tint < 0 ? luminance * (1 + tint) : luminance * (1 - tint) + tint;
        var chroma = (1 - Math.Abs(2 * luminance - 1)) * saturation;
        var x = chroma * (1 - Math.Abs(hue % 2 - 1));
        var m = luminance - chroma / 2;
        (r, g, b) = hue switch { < 1 => (chroma, x, 0d), < 2 => (x, chroma, 0d), < 3 => (0d, chroma, x),
            < 4 => (0d, x, chroma), < 5 => (x, 0d, chroma), _ => (chroma, 0d, x) };
        string Channel(double channel) => Math.Clamp((int)Math.Round((channel + m) * 255), 0, 255).ToString("X2", CultureInfo.InvariantCulture);
        return argb[..2] + Channel(r) + Channel(g) + Channel(b);
    }
}
