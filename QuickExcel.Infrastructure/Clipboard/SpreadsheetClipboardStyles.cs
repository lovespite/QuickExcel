using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Clipboard;

internal static class SpreadsheetClipboardStyles
{
    internal static readonly XNamespace Ss = "urn:schemas-microsoft-com:office:spreadsheet";
    internal static string? A(XElement? e, string name) => (string?)e?.Attribute(Ss + name);
    internal static int Int(XElement e, string name, int fallback = 0) => A(e, name) is { } text ? int.Parse(text, CultureInfo.InvariantCulture) : fallback;
    internal static string NumberFormat(string? value) => value switch
    {
        null or "General" => "General", "General Number" => "General", "General Date" => "m/d/yy h:mm",
        "Short Date" => "m/d/yy", "Long Date" => "dddd, mmmm dd, yyyy", "Medium Date" => "dd-mmm-yy",
        "Long Time" => "h:mm:ss AM/PM", "Medium Time" => "h:mm AM/PM", "Short Time" => "h:mm",
        "Currency" => "\"$\"#,##0.00", "Euro Currency" => "\"€\"#,##0.00", "Fixed" => "0.00", "Standard" => "#,##0.00",
        "Percent" => "0.00%", "Scientific" => "0.00E+00", "Yes/No" => "[=0]\"No\";[=1]\"Yes\";General",
        "True/False" => "[=0]\"False\";[=1]\"True\";General", "On/Off" => "[=0]\"Off\";[=1]\"On\";General", _ => value
    };
    internal static StyleDefinition Read(XElement element, StyleDefinition parent, ISet<string> limitations)
    {
        var s = parent;
        foreach (var e in element.Elements())
        {
            var known = e.Name.LocalName switch
            {
                "Font" => new[] { "FontName", "Size", "Color", "Bold", "Italic", "Underline", "StrikeThrough", "CharSet", "Family" },
                "Alignment" => ["Horizontal", "Vertical", "WrapText", "Rotate", "VerticalText", "Indent", "ShrinkToFit"],
                "Interior" => ["Color", "Pattern", "PatternColor"], "NumberFormat" => ["Format"], "Protection" => ["Protected", "HideFormula"], "Borders" => [], _ => null
            };
            if (known is null || e.Attributes().Any(a => !a.IsNamespaceDeclaration && !known.Contains(a.Name.LocalName))) limitations.Add("含未支持的单元格样式属性");
            switch (e.Name.LocalName)
            {
                case "Font":
                    s = s with { FontName = A(e, "FontName") ?? parent.FontName, FontSize = A(e, "Size") is { } size ? double.Parse(size, CultureInfo.InvariantCulture) : parent.FontSize,
                        FontColor = Color(A(e, "Color")) ?? parent.FontColor, Bold = Flag(e, "Bold", parent.Bold), Italic = Flag(e, "Italic", parent.Italic),
                        Strike = Flag(e, "StrikeThrough", parent.Strike), Underline = A(e, "Underline") is { } u ? u switch { "Single" => "single", "Double" => "double", "SingleAccounting" => "singleAccounting", "DoubleAccounting" => "doubleAccounting", "None" => null, _ => Unsupported<string?>(limitations, null) } : parent.Underline };
                    break;
                case "Alignment":
                    s = s with { Horizontal = EnumValue(A(e, "Horizontal"), ["Automatic", "Left", "Center", "Right", "Justify"], parent.Horizontal, limitations),
                        Vertical = EnumValue(A(e, "Vertical"), ["Top", "Center", "Bottom", "Justify"], parent.Vertical, limitations),
                        WrapText = Flag(e, "WrapText", parent.WrapText), ShrinkToFit = Flag(e, "ShrinkToFit", parent.ShrinkToFit), Indent = Int(e, "Indent", parent.Indent),
                        TextRotation = Flag(e, "VerticalText", false) ? 255 : A(e, "Rotate") is { } rot ? Rotation(int.Parse(rot, CultureInfo.InvariantCulture)) : parent.TextRotation };
                    break;
                case "Interior":
                    if (A(e, "Pattern") is not (null or "Solid" or "None")) limitations.Add("含未支持的图案填充");
                    s = s with { FillColor = A(e, "Pattern") == "Solid" ? Color(A(e, "Color")) : null }; break;
                case "NumberFormat": s = s with { NumberFormatId = 0, NumberFormatCode = NumberFormat(A(e, "Format")) }; break;
                case "Protection": s = s with { Locked = Flag(e, "Protected", true), FormulaHidden = Flag(e, "HideFormula", false) }; break;
                case "Borders":
                    s = s with { DiagonalBorder = 0, DiagonalBorderColor = null, DiagonalUp = false, DiagonalDown = false, LeftBorder = 0, RightBorder = 0, TopBorder = 0, BottomBorder = 0, LeftBorderColor = null, RightBorderColor = null, TopBorderColor = null, BottomBorderColor = null };
                    foreach (var b in e.Elements())
                    {
                        var kind = (A(b, "LineStyle"), Int(b, "Weight")) switch {
                            (null or "None", _) => 0, ("Continuous", 0) => 7, ("Continuous", 1) => 1, ("Continuous", 2) => 2, ("Continuous", 3) => 5,
                            ("Dash", 2) => 8, ("DashDot", 1) => 9, ("DashDot", 2) => 10, ("DashDotDot", 1) => 11, ("DashDotDot", 2) => 12, ("SlantDashDot", 2) => 13, ("Dash", 1) => 3, ("Dot", 1) => 4, ("Double", 3) => 6, _ => Unsupported(limitations, 0) };
                        var color = Color(A(b, "Color"));
                        s = A(b, "Position") switch { "DiagonalLeft" => s with { DiagonalBorder = kind, DiagonalBorderColor = color, DiagonalDown = kind != 0 }, "DiagonalRight" => s with { DiagonalBorder = kind, DiagonalBorderColor = color, DiagonalUp = kind != 0 }, "Left" => s with { LeftBorder = kind, LeftBorderColor = color }, "Right" => s with { RightBorder = kind, RightBorderColor = color },
                            "Top" => s with { TopBorder = kind, TopBorderColor = color }, "Bottom" => s with { BottomBorder = kind, BottomBorderColor = color }, _ => Unsupported(limitations, s) };
                    }
                    break;
            }
        }
        return s;
    }
    private static T Unsupported<T>(ISet<string> d, T fallback) { d.Add("含未支持的样式或边框"); return fallback; }
    internal static int Rotation(int r) => r is < -90 or > 90 ? throw new FormatException("文字旋转须在 -90 至 90 度之间。") : r >= 0 ? r : 90 - r;
    private static int EnumValue(string? text, string[] values, int fallback, ISet<string> d) => text is null ? fallback : Array.IndexOf(values, text) is var i && i >= 0 ? i : Unsupported(d, fallback);
    private static bool Flag(XElement e, string name, bool fallback) => A(e, name) switch { null => fallback, "1" => true, "0" => false, _ => throw new FormatException("无效样式布尔值。") };
    internal static string? Color(string? color)
    {
        if (color is null) return null;
        if (color.Length != 7 || color[0] != '#' || !uint.TryParse(color.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) throw new FormatException("无效剪贴板颜色。");
        return "FF" + color[1..].ToUpperInvariant();
    }
    internal static string Rgb(string color)
    {
        if (color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c))) throw new FormatException("内部颜色必须是 ARGB。");
        return "#" + color[2..];
    }
    internal static XElement Write(StyleDefinition s, string id)
    {
        XAttribute Attr(string name, object value) => new(Ss + name, value);
        var format = s.NumberFormatCode ?? CellDisplayFormatter.BuiltIn(s.NumberFormatId) ?? (s.NumberFormatId == 0 ? "General" : throw new NotSupportedException("数字格式无法导出。"));
        XElement Border(string side, int kind, string? color) => new(Ss + "Border", Attr("Position", side),
            Attr("LineStyle", kind switch { 3 or 8 => "Dash", 9 or 10 => "DashDot", 11 or 12 => "DashDotDot", 13 => "SlantDashDot", 4 => "Dot", 6 => "Double", _ => "Continuous" }),
            Attr("Weight", kind switch { 2 or 8 or 10 or 12 or 13 => 2, 5 or 6 => 3, 7 => 0, _ => 1 }), color is null ? null : Attr("Color", Rgb(color)));
        return new(Ss + "Style", Attr("ID", id),
            new XElement(Ss + "Alignment", Attr("Horizontal", new[] { "Automatic", "Left", "Center", "Right", "Justify" }[s.Horizontal]),
                Attr("Vertical", new[] { "Top", "Center", "Bottom", "Justify" }[s.Vertical]), Attr("WrapText", s.WrapText ? 1 : 0), Attr("ShrinkToFit", s.ShrinkToFit ? 1 : 0),
                Attr("Indent", s.Indent), s.TextRotation == 255 ? Attr("VerticalText", 1) : Attr("Rotate", s.TextRotation <= 90 ? s.TextRotation : 90 - s.TextRotation)),
            new XElement(Ss + "Font", Attr("FontName", s.FontName ?? "Calibri"), Attr("Size", s.FontSize ?? 11), Attr("Color", Rgb(s.FontColor ?? "FF000000")),
                Attr("Bold", s.Bold ? 1 : 0), Attr("Italic", s.Italic ? 1 : 0), Attr("StrikeThrough", s.Strike ? 1 : 0),
                s.Underline is null or "none" ? null : Attr("Underline", char.ToUpperInvariant(s.Underline[0]) + s.Underline[1..])),
            new XElement(Ss + "Interior", s.FillColor is null ? Array.Empty<XAttribute>() : new[] { Attr("Pattern", "Solid"), Attr("Color", Rgb(s.FillColor)) }),
            new XElement(Ss + "NumberFormat", Attr("Format", format)), new XElement(Ss + "Protection", Attr("Protected", s.Locked ? 1 : 0), Attr("HideFormula", s.FormulaHidden ? 1 : 0)),
            new XElement(Ss + "Borders", s.LeftBorder == 0 ? null : Border("Left", s.LeftBorder, s.LeftBorderColor), s.RightBorder == 0 ? null : Border("Right", s.RightBorder, s.RightBorderColor),
                s.TopBorder == 0 ? null : Border("Top", s.TopBorder, s.TopBorderColor), s.BottomBorder == 0 ? null : Border("Bottom", s.BottomBorder, s.BottomBorderColor),
                !s.DiagonalDown || s.DiagonalBorder == 0 ? null : Border("DiagonalLeft", s.DiagonalBorder, s.DiagonalBorderColor),
                !s.DiagonalUp || s.DiagonalBorder == 0 ? null : Border("DiagonalRight", s.DiagonalBorder, s.DiagonalBorderColor)));
    }
}
