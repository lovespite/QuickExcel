using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Clipboard;

/// <summary>CF_HTML context and fragment offsets are UTF-8 byte offsets. No browser, script or resource loader.</summary>
public static partial class HtmlClipboardCodec
{
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private const string Header = "Version:1.0\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
    public static byte[] Wrap(string html, int fragmentStart, int fragmentEnd)
    {
        var prefix = string.Format(CultureInfo.InvariantCulture, Header, 0, 0, 0, 0).Length;
        var header = string.Format(CultureInfo.InvariantCulture, Header, prefix, prefix + Utf8.GetByteCount(html),
            prefix + Utf8.GetByteCount(html.AsSpan(0, fragmentStart)), prefix + Utf8.GetByteCount(html.AsSpan(0, fragmentEnd)));
        var bytes = Utf8.GetBytes(header + html);
        if (bytes.Length > SpreadsheetClipboardCodec.MaxBytes) throw new InvalidDataException("HTML 剪贴板超过 8 MiB。");
        return bytes;
    }
    public static string Unwrap(byte[] bytes)
    {
        if (bytes.Length == 0 || bytes.Length > SpreadsheetClipboardCodec.MaxBytes) throw new InvalidDataException("HTML 剪贴板为空或超过 8 MiB。");
        var header = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 4096));
        int Offset(string key)
        {
            var match = Regex.Match(header, @"(?m)^" + key + @":(-?\d+)\r?$", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
            if (!match.Success || !int.TryParse(match.Groups[1].Value, out var value)) throw new FormatException("CF_HTML 缺少有效字节偏移。");
            return value;
        }
        var start = Offset("StartHTML"); var end = Offset("EndHTML"); var fs = Offset("StartFragment"); var fe = Offset("EndFragment");
        if (start == -1 && end == -1) { start = fs; end = fe; }
        if (start < 0 || end > bytes.Length || fs < start || fe > end || fs > fe || start >= end)
            throw new FormatException("CF_HTML 字节偏移越界。");
        // Context may contain styles and the outer table, but must not add unselected cells.
        var outside = Utf8.GetString(bytes, start, fs - start) + Utf8.GetString(bytes, fe, end - fe);
        if (Regex.IsMatch(outside, @"<\s*(tr|td|th)(?:\s|>)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100)))
            throw new NotSupportedException("HTML 片段未包含完整表格，不能扩大粘贴范围。");
        // Decode both slices to reject offsets in the middle of a multibyte character.
        _ = Utf8.GetString(bytes, fs, fe - fs);
        return Utf8.GetString(bytes, start, end - start);
    }
    public static ExternalClipboardRegion Read(byte[] bytes)
    {
        var html = Unwrap(bytes); var count = 0;
        var parser = new HtmlParser(new HtmlParserOptions { IsScripting = false, MaximumDepth = SpreadsheetClipboardCodec.MaxDepth, OnCreated = (e, _) =>
        {
            if ((count += 1 + e.Attributes.Length) > SpreadsheetClipboardCodec.MaxNodes) throw new InvalidDataException("HTML 节点超过上限。");
        } });
        using var document = parser.ParseDocument(html);
        var stack = new Stack<(INode Node, int Depth)>(); stack.Push((document, 0)); count = 0;
        while (stack.TryPop(out var item))
        {
            if (++count > SpreadsheetClipboardCodec.MaxNodes || item.Depth > SpreadsheetClipboardCodec.MaxDepth) throw new InvalidDataException("HTML 节点/深度超过上限。");
            foreach (var node in item.Node.ChildNodes) stack.Push((node, item.Depth + 1));
        }
        var tables = document.QuerySelectorAll("table");
        if (tables.Length != 1) throw new NotSupportedException("HTML 必须只包含一个无嵌套表格。");
        var limitations = new HashSet<string>();
        if (document.QuerySelector("script,iframe,object,img,svg,a,span[style],font") is not null) limitations.Add("HTML 含脚本、对象、链接或富文本，未执行或加载");
        var rules = new List<(string Selector, string Css)>();
        foreach (var style in document.QuerySelectorAll("style"))
            foreach (Match match in Rules().Matches(style.TextContent.Replace("<!--", "").Replace("-->", "")))
                foreach (var selector in match.Groups[1].Value.Split(','))
                {
                    var s = selector.Trim();
                    if (s is "td" or "th" || SimpleClass().IsMatch(s)) rules.Add((s, match.Groups[2].Value));
                    else if (!s.StartsWith('@') && s is not ("table" or "tr" or "col" or "br" or "ruby" or "rt")) limitations.Add("HTML 含未支持的样式选择器");
                    if (rules.Count > SpreadsheetClipboardCodec.MaxStyles) throw new InvalidDataException("HTML 样式超过上限。");
                }
        var rows = tables[0].QuerySelectorAll("tr");
        if (rows.Length == 0 || rows.Length > WorkbookClipboard.MaxCells) throw new InvalidDataException("HTML 表格行数无效。");
        var cells = new Dictionary<CellPosition, ExternalClipboardCell>(); var occupied = new HashSet<CellPosition>(); var merges = new List<CellRectangle>(); var columns = 0;
        for (var r = 0; r < rows.Length; r++)
        {
            var c = 0;
            foreach (var td in rows[r].Children.Where(e => e.LocalName is "td" or "th"))
            {
                while (occupied.Contains(new(r, c))) c++;
                int Span(string name) => td.GetAttribute(name) is { } s ? int.Parse(s, CultureInfo.InvariantCulture) : 1;
                var rs = Span("rowspan"); var cs = Span("colspan");
                if (rs < 1 || cs < 1 || (long)r + rs > rows.Length || (long)(c + (long)cs) * rows.Length > WorkbookClipboard.MaxCells) throw new FormatException("HTML 合并/矩形尺寸越界。");
                for (var rr = r; rr < r + rs; rr++) for (var cc = c; cc < c + cs; cc++) if (!occupied.Add(new(rr, cc))) throw new FormatException("HTML 合并区域重叠。");
                if (rs > 1 || cs > 1) merges.Add(new(r, c, r + rs - 1, c + cs - 1));
                var css = string.Join(";", rules.Where(x => x.Selector == td.LocalName || x.Selector.StartsWith('.') && td.ClassList.Contains(x.Selector[1..])).OrderBy(x => x.Selector.StartsWith('.') ? 1 : 0).Select(x => x.Css)) + ";" + td.GetAttribute("style");
                var style = Css(css, limitations);
                foreach (var decoration in td.QuerySelectorAll("s,strike,u"))
                {
                    if (decoration.TextContent != td.TextContent) { limitations.Add("HTML 含单元格内局部文字格式"); continue; }
                    style = decoration.LocalName == "u" ? style with { Underline = style.Underline ?? "single" } : style with { Strike = true };
                }
                var text = new StringBuilder();
                void Text(INode n) { if (n is IText t) text.Append(t.Data); else if (n is IElement { LocalName: "br" }) text.Append('\n'); else if (n is not IElement { LocalName: "script" or "style" }) foreach (var child in n.ChildNodes) Text(child); }
                Text(td);
                var value = text.ToString();
                if (value.Length > WorkbookLimits.MaxCellTextLength) throw new InvalidDataException("HTML 单元格文本过长。");
                cells.Add(new(r, c), new(value.Length == 0 ? CellValue.Blank : CellValue.FromText(value), style));
                c += cs; columns = Math.Max(columns, c);
            }
        }
        if (columns == 0) throw new FormatException("HTML 表格没有单元格。");
        var region = new ExternalClipboardRegion(rows.Length, columns, Enumerable.Range(0, rows.Length).SelectMany(r => Enumerable.Range(0, columns).Select(c => cells.GetValueOrDefault(new(r, c)) ?? new(CellValue.Blank, new()))), merges, limitations, valuesAreDisplayText: true);
        _ = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new WorksheetSnapshot(SheetId.New(), "Validation")], styles: region.Cells.Select(c => c.Style!).Distinct());
        return region;
    }
    [GeneratedRegex(@"([^{}]+)\{([^{}]*)\}", RegexOptions.CultureInvariant, 100)] private static partial Regex Rules();
    [GeneratedRegex(@"^\.[a-zA-Z_][\w-]*$", RegexOptions.CultureInvariant, 100)] private static partial Regex SimpleClass();
    private static StyleDefinition Css(string text, ISet<string> limitations)
    {
        var style = new StyleDefinition(); var declarations = new List<string>(); var start = 0; char quote = '\0';
        for (var i = 0; i <= text.Length; i++)
        {
            if (i == text.Length || text[i] == ';' && quote == '\0') { declarations.Add(text[start..i]); start = i + 1; continue; }
            if (text[i] == '\\') { i++; continue; }
            if (text[i] is '\'' or '"') quote = quote == '\0' ? text[i] : quote == text[i] ? '\0' : quote;
        }
        foreach (var declaration in declarations)
        {
            var parts = declaration.Split(':', 2); if (parts.Length != 2) continue;
            var key = parts[0].Trim().ToLowerInvariant(); var value = parts[1].Trim();
            string Unquote(string s) => s.Trim('"', '\'').Replace("\\:", ":").Replace("\\@", "@").Replace("\\\\", "\\");
            switch (key)
            {
                case "font-family": style = style with { FontName = Unquote(value.Split(',')[0]) }; break;
                case "font-size":
                    if (!value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) { limitations.Add("HTML 字号单位尚未支持"); break; }
                    style = style with { FontSize = double.Parse(value[..^2], CultureInfo.InvariantCulture) }; break;
                case "font-weight": style = style with { Bold = value is "bold" or "700" or "800" or "900" }; break;
                case "font-style": style = style with { Italic = value == "italic" }; break;
                case "color": style = style with { FontColor = CssColor(value, limitations) }; break;
                case "background": case "background-color": style = style with { FillColor = CssColor(value, limitations) }; break;
                case "text-align": style = style with { Horizontal = value switch { "left" => 1, "center" => 2, "right" => 3, "justify" => 4, _ => 0 } }; break;
                case "vertical-align": style = style with { Vertical = value switch { "top" => 0, "middle" => 1, "bottom" => 2, _ => 2 } }; break;
                case "white-space": style = style with { WrapText = value is "normal" or "pre-wrap" }; break;
                case "mso-number-format": style = style with { NumberFormatCode = SpreadsheetClipboardStyles.NumberFormat(Unquote(value)) }; break;
                case "text-decoration": style = style with { Underline = value.Contains("underline") ? "single" : null, Strike = value.Contains("line-through") }; break;
                case "text-underline-style":
                    if (value is not ("single" or "double" or "single-accounting" or "double-accounting" or "none")) { limitations.Add("HTML 下划线类型尚未支持"); break; }
                    style = style with { Underline = value switch { "single-accounting" => "singleAccounting", "double-accounting" => "doubleAccounting", "none" => null, _ => value } }; break;
                case "mso-char-indent-count": style = style with { Indent = int.Parse(value, CultureInfo.InvariantCulture) }; break;
                case "mso-rotate": style = style with { TextRotation = SpreadsheetClipboardStyles.Rotation(int.Parse(value, CultureInfo.InvariantCulture)) }; break;
                case "layout-flow":
                    if (value == "vertical-ideographic") style = style with { TextRotation = 255 };
                    else if (value != "horizontal") limitations.Add("HTML 文字排列尚未支持");
                    break;
                case "border": case "border-left": case "border-right": case "border-top": case "border-bottom":
                    var bits = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    var kind = BorderKind(bits);
                    if (kind < 0) { limitations.Add("HTML 边框尚未支持"); break; }
                    var color = bits.Length > 2 ? CssColor(bits[^1], limitations) : null;
                    if (key is "border" or "border-left") style = style with { LeftBorder = kind, LeftBorderColor = color };
                    if (key is "border" or "border-right") style = style with { RightBorder = kind, RightBorderColor = color };
                    if (key is "border" or "border-top") style = style with { TopBorder = kind, TopBorderColor = color };
                    if (key is "border" or "border-bottom") style = style with { BottomBorder = kind, BottomBorderColor = color };
                    break;
                default:
                    if (!key.StartsWith("padding") && key is not ("height" or "width" or "mso-ignore" or "mso-generic-font-family" or "mso-font-charset" or "mso-background-source" or "mso-pattern" or "mso-protection")) limitations.Add("HTML 含未支持的样式：" + key);
                    break;
            }
        }
        return style;
    }
    private static int BorderKind(string[] bits)
    {
        if (bits is ["none"]) return 0;
        if (bits.Length < 2 || !bits[0].EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
            !double.TryParse(bits[0].AsSpan(0, bits[0].Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out var width)) return -1;
        return (bits[1], width) switch {
            ("hairline", .5) => 7, ("solid", .5) => 1, ("solid", 1) => 2, ("solid", 1.5) => 5,
            ("dashed", .5) => 3, ("dotted", .5) => 4, ("double", 2) => 6, _ => -1 };
    }
    private static string BorderCss(int kind) => kind switch {
        1 => ".5pt solid", 2 => "1.0pt solid", 3 => ".5pt dashed", 4 => ".5pt dotted", 5 => "1.5pt solid", 6 => "2.0pt double", 7 => ".5pt hairline",
        _ => throw new NotSupportedException("HTML 边框类型尚未支持。") };
    private static string? CssColor(string value, ISet<string> limitations)
    {
        value = value.ToLowerInvariant() switch { "black" or "windowtext" => "#000000", "white" => "#FFFFFF", "red" => "#FF0000", "yellow" => "#FFFF00", "blue" => "#0000FF", "green" => "#008000", "transparent" => "", _ => value };
        if (value == "") return null;
        try { return SpreadsheetClipboardStyles.Color(value); } catch (FormatException) { limitations.Add("HTML 颜色尚未支持"); return null; }
    }
    public static byte[] Write(ExternalClipboardRegion region, bool formulas = false)
    {
        region.RequireExport(formulas);
        string E(string s) => WebUtility.HtmlEncode(s).Replace("\r", "&#13;").Replace("\n", "&#10;").Replace("\t", "&#9;");
        var b = new StringBuilder("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><head><meta charset=\"utf-8\"></head><body><!--StartFragment--><table>"); var start = b.Length - "<table>".Length;
        var covered = new HashSet<CellPosition>(); var anchors = region.Merges.ToDictionary(m => new CellPosition(m.Top, m.Left));
        foreach (var m in region.Merges) for (var r = m.Top; r <= m.Bottom; r++) for (var c = m.Left; c <= m.Right; c++) if (r != m.Top || c != m.Left) covered.Add(new(r, c));
        for (var r = 0; r < region.Rows; r++)
        {
            b.Append("<tr>");
            for (var c = 0; c < region.Columns; c++)
            {
                if (covered.Contains(new(r, c))) continue;
                var cell = region.Cells[r * region.Columns + c]; var s = cell.Style ?? new();
                var format = s.NumberFormatCode ?? CellDisplayFormatter.BuiltIn(s.NumberFormatId) ?? "General";
                var css = $"font-family:'{s.FontName ?? "Calibri"}';font-size:{(s.FontSize ?? 11).ToString(CultureInfo.InvariantCulture)}pt;font-weight:{(s.Bold ? 700 : 400)};font-style:{(s.Italic ? "italic" : "normal")};color:{SpreadsheetClipboardStyles.Rgb(s.FontColor ?? "FF000000")};white-space:{(s.WrapText ? "pre-wrap" : "nowrap")};mso-number-format:'{format.Replace("'", "\\'")}'";
                if (s.FillColor is { } fill) css += ";background-color:" + SpreadsheetClipboardStyles.Rgb(fill);
                css += ";text-decoration:" + (s.Underline is not (null or "none") ? (s.Strike ? "underline line-through" : "underline") : s.Strike ? "line-through" : "none");
                if (s.Underline is not (null or "none")) css += ";text-underline-style:" + (s.Underline switch { "singleAccounting" => "single-accounting", "doubleAccounting" => "double-accounting", _ => s.Underline });
                css += ";mso-char-indent-count:" + s.Indent.ToString(CultureInfo.InvariantCulture);
                css += s.TextRotation == 255 ? ";layout-flow:vertical-ideographic" : ";mso-rotate:" + (s.TextRotation <= 90 ? s.TextRotation : 90 - s.TextRotation).ToString(CultureInfo.InvariantCulture);
                css += ";text-align:" + new[] { "general", "left", "center", "right", "justify" }[s.Horizontal] + ";vertical-align:" + new[] { "top", "middle", "bottom", "middle" }[s.Vertical];
                foreach (var (side, kind, color) in new[] { ("left", s.LeftBorder, s.LeftBorderColor), ("right", s.RightBorder, s.RightBorderColor), ("top", s.TopBorder, s.TopBorderColor), ("bottom", s.BottomBorder, s.BottomBorderColor) })
                    css += $";border-{side}:" + (kind == 0 ? "none" : BorderCss(kind) + " " + SpreadsheetClipboardStyles.Rgb(color ?? "FF000000"));
                b.Append("<td style=\"").Append(E(css)).Append('"');
                if (anchors.TryGetValue(new(r, c), out var m)) b.Append($" rowspan=\"{m.Bottom - m.Top + 1}\" colspan=\"{m.Right - m.Left + 1}\"");
                var display = CellInput.DisplayText(cell.Value);
                if (formulas && cell.Formula is { } formula)
                {
                    var expression = ExternalClipboardFormula.Validate(formula);
                    expression = FormulaSyntax.Map(expression, n =>
                    {
                        if (n is not ReferenceNode reference) return n;
                        var row = checked(reference.Row - cell.FormulaOrigin.Row + r); var col = checked(reference.Column - cell.FormulaOrigin.Column + c);
                        var er = checked(reference.EndRow - cell.FormulaOrigin.Row + r); var ec = checked(reference.EndColumn - cell.FormulaOrigin.Column + c);
                        if (row < 0 || col < 0 || er < 0 || ec < 0) throw new NotSupportedException("HTML 公式引用超出复制矩形上方/左侧；请扩大复制范围包含依赖，或复制结果值。");
                        return reference with { Row = row, Column = col, EndRow = er, EndColumn = ec };
                    });
                    b.Append(" x:fmla=\"").Append(E("=" + FormulaSyntax.Print(expression))).Append('"');
                }
                if (cell.Date is { } date)
                {
                    if (date.IsFictitiousLeapDay || !date.TryToDateTime(out var actual) || actual.Year < 1900)
                        throw new NotSupportedException("此边界日期不能通过 Excel HTML 安全交换；可将日期显式转换为数值后复制。");
                    // A calendar string is interpreted in the destination's own date system. x:num would
                    // copy a 1900 serial unchanged into a 1904 workbook. Never publish workbook-level options.
                    display = actual.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                }
                else if (cell.Value.Kind != CellValueKind.Blank)
                {
                    if (cell.Value.Kind == CellValueKind.Error && cell.Value.ErrorCode is not (CellErrorCode.Null or CellErrorCode.DivisionByZero or CellErrorCode.Value or CellErrorCode.Reference or CellErrorCode.Name or CellErrorCode.Number or CellErrorCode.NotAvailable))
                        throw new NotSupportedException("此 QuickExcel 错误类型没有已验证的 Excel 剪贴板表示。");
                    var attribute = cell.Value.Kind switch { CellValueKind.Text => "str", CellValueKind.Number => "num", CellValueKind.Boolean => "bool", _ => "err" };
                    var raw = cell.Value.Kind == CellValueKind.Error ? "" : cell.Value.Kind == CellValueKind.Boolean ? (cell.Value.BooleanValue ? "TRUE" : "FALSE") : display;
                    b.Append(" x:").Append(attribute).Append("=\"").Append(E(raw)).Append('"');
                }
                b.Append('>').Append(E(display)).Append("</td>");
            }
            b.Append("</tr>");
        }
        b.Append("</table>"); var end = b.Length; b.Append("<!--EndFragment--></body></html>");
        return Wrap(b.ToString(), start, end);
    }
}
