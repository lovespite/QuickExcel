using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using static QuickExcel.Clipboard.SpreadsheetClipboardStyles;

namespace QuickExcel.Clipboard;

/// <summary>SpreadsheetML clipboard data, not XLSX or an internal WorkbookClipboard serialization.</summary>
public static class SpreadsheetClipboardCodec
{
    public const int MaxBytes = 8 * 1024 * 1024;
    public const int MaxNodes = 200_000;
    public const int MaxDepth = 32;
    public const int MaxStyles = 2048;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    public static ExternalClipboardRegion Read(byte[] bytes, byte[]? biff8 = null)
    {
        if (bytes.Length == 0 || bytes.Length > MaxBytes) throw new InvalidDataException("XML 剪贴板载荷为空或超过 8 MiB。");
        // Excel publishes UTF-8 with one NUL; tolerate UTF-16 BOM only when explicitly declared by its bytes.
        var length = bytes.Length;
        var utf16 = length >= 2 && (bytes[0] == 255 && bytes[1] == 254 || bytes[0] == 254 && bytes[1] == 255);
        if (utf16) { while (length >= 2 && bytes[length - 1] == 0 && bytes[length - 2] == 0) length -= 2; }
        else while (length > 0 && bytes[length - 1] == 0) length--;
        using var stream = new MemoryStream(bytes, 0, length, false);
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxBytes, IgnoreComments = true };
        // Preflight before materializing the tree: depth/node budgets include attributes.
        using (var reader = XmlReader.Create(stream, settings))
        {
            var nodes = 0;
            while (reader.Read()) if ((nodes += 1 + reader.AttributeCount) > MaxNodes || reader.Depth > MaxDepth) throw new InvalidDataException("XML 剪贴板节点/深度超过上限。");
        }
        stream.Position = 0;
        using var parsed = XmlReader.Create(stream, settings);
        var root = XDocument.Load(parsed, LoadOptions.PreserveWhitespace).Root;
        if (root?.Name != Ss + "Workbook") throw new InvalidDataException("不是 XML Spreadsheet 工作簿。");
        var sheets = root.Elements(Ss + "Worksheet").ToArray();
        if (sheets.Length != 1 || sheets[0].Elements(Ss + "Table").Count() != 1) throw new NotSupportedException("剪贴板只支持单个连续矩形。");
        var limitations = new HashSet<string>();
        var excel = (XNamespace)"urn:schemas-microsoft-com:office:excel";
        // Excel can publish only the reserved name/NamedCell markers, without an AutoFilter element.
        // These identify source filtering; they do not authorize importing criteria or hiding rows.
        bool FilterName(XElement e) => (e.Name == Ss + "NamedRange" || e.Name == Ss + "NamedCell") &&
            A(e, "Name") is "_FilterDatabase" or "_xlnm._FilterDatabase";
        bool FilterMetadata(XElement e) => FilterName(e) || e.Name == excel + "AutoFilter";
        var hasFilterMetadata = root.Descendants().Any(FilterMetadata);
        if (root.Descendants().Any(e => !FilterMetadata(e) &&
            (e.Name.LocalName is "Comment" or "NamedRange" or "NamedCell" or "DataValidation" or "ConditionalFormatting" or "AutoFilter" or "PivotTable" ||
             e.Name.LocalName == "Names" && (e.Name != Ss + "Names" || e.Elements().Any(n => !FilterName(n)))) || e.Attribute(Ss + "HRef") is not null))
            limitations.Add("来源含批注、链接、名称或规则，完整互粘尚未支持");
        var definitions = root.Element(Ss + "Styles")?.Elements(Ss + "Style").ToArray() ?? [];
        if (definitions.Length > MaxStyles) throw new InvalidDataException("剪贴板样式超过 2048 项。");
        var raw = definitions.ToDictionary(e => A(e, "ID") ?? throw new FormatException("样式缺少 ID。"));
        var resolved = new Dictionary<string, StyleDefinition>(); var visiting = new HashSet<string>();
        StyleDefinition Style(string? id, int depth = 0)
        {
            id ??= "Default";
            if (resolved.TryGetValue(id, out var result)) return result;
            if (depth > MaxDepth || !visiting.Add(id)) throw new FormatException("样式继承循环或过深。");
            if (!raw.TryGetValue(id, out var e))
            {
                if (id != "Default") throw new FormatException("剪贴板引用未知样式。");
                result = new() { FontName = "Calibri", FontSize = 11, FontColor = "FF000000", NumberFormatCode = "General" };
            }
            else result = SpreadsheetClipboardStyles.Read(e, id == "Default" ? new() : Style(A(e, "Parent") ?? "Default", depth + 1), limitations);
            visiting.Remove(id); resolved.Add(id, result); return result;
        }
        var table = sheets[0].Element(Ss + "Table")!;
        var rows = Int(table, "ExpandedRowCount"); var columns = Int(table, "ExpandedColumnCount");
        if (rows < 1 || columns < 1 || (long)rows * columns > WorkbookClipboard.MaxCells) throw new InvalidDataException("剪贴板缺少矩形尺寸或超过 20,000 格。");
        // Parse the optional companion only when calendar/time values need its original numeric precision/epoch.
        var numeric = biff8 is not null && table.Descendants(Ss + "Data").Any(d => A(d, "Type") == "DateTime") ? Biff8ClipboardNumbers.Read(biff8) : null;
        if (numeric is not null && (numeric.Rows != rows || numeric.Columns != columns)) throw new InvalidDataException("同一剪贴板的 XML 与 Biff8 矩形不一致。");
        var columnStyles = new string?[columns]; var ci = 0;
        foreach (var col in table.Elements(Ss + "Column"))
        {
            ci = Int(col, "Index", ci + 1) - 1; var span = Int(col, "Span");
            if (ci < 0 || span < 0 || (long)ci + span >= columns) throw new FormatException("列范围越界。");
            for (var k = 0; k <= span; k++) columnStyles[ci + k] = A(col, "StyleID");
            ci += span + 1;
        }
        var cells = new ExternalClipboardCell[rows * columns];
        for (var r = 0; r < rows; r++) for (var c = 0; c < columns; c++) cells[r * columns + c] = new(CellValue.Blank, Style(columnStyles[c]));
        var occupied = new bool[cells.Length]; var merges = new List<CellRectangle>(); var ri = 0;
        foreach (var row in table.Elements(Ss + "Row"))
        {
            var next = Int(row, "Index", ri + 1) - 1;
            if (next < ri || next >= rows) throw new FormatException("行索引重复或越界。");
            ri = next;
            for (var c = 0; c < columns; c++) cells[ri * columns + c] = new(CellValue.Blank, Style(A(row, "StyleID") ?? columnStyles[c]));
            ci = 0;
            foreach (var cell in row.Elements(Ss + "Cell"))
            {
                var nextColumn = Int(cell, "Index", ci + 1) - 1;
                if (nextColumn < ci || nextColumn >= columns || occupied[ri * columns + nextColumn]) throw new FormatException("单元格索引重复、覆盖合并或越界。");
                ci = nextColumn;
                var across = Int(cell, "MergeAcross"); var down = Int(cell, "MergeDown");
                if (across < 0 || down < 0 || (long)ci + across >= columns || (long)ri + down >= rows) throw new FormatException("合并区域越界。");
                for (var mr = ri; mr <= ri + down; mr++) for (var mc = ci; mc <= ci + across; mc++)
                { if (occupied[mr * columns + mc]) throw new FormatException("合并区域重叠。"); occupied[mr * columns + mc] = true; }
                if (across + down > 0) merges.Add(new(ri, ci, ri + down, ci + across));
                if (cell.Elements(Ss + "Data").Count() > 1) throw new FormatException("单元格有多个数据元素。");
                var data = cell.Element(Ss + "Data"); var style = Style(A(cell, "StyleID") ?? A(row, "StyleID") ?? columnStyles[ci]);
                if (data?.HasElements == true) limitations.Add("来源含单元格内富文本");
                var formula = A(cell, "Formula"); string? a1 = null, rejection = null;
                var origin = SpreadsheetClipboardFormula.Origin;
                if (formula is not null)
                {
                    try { a1 = SpreadsheetClipboardFormula.Read(formula); }
                    catch (Exception ex) when (ex is FormatException or NotSupportedException or OverflowException or ArgumentException) { rejection = ex.Message; }
                }
                if (A(cell, "ArrayRange") is not null) rejection = "来源含数组公式";
                var (value, date) = ReadValue(data);
                if (date is { } calendar)
                {
                    if (numeric is not null && numeric.Numbers.TryGetValue(new(ri, ci), out var rawValue))
                    {
                        // SpreadsheetML rounds DateTime to milliseconds; preserve the original raw double.
                        const double tolerance = 1d / 86400000 + 1e-9;
                        ExcelDateSystem system;
                        if (Math.Abs(calendar.Serial - rawValue) < tolerance) system = ExcelDateSystem.Windows1900;
                        else if (new ExcelSerialDate(rawValue, ExcelDateSystem.Mac1904).TryConvertTo(ExcelDateSystem.Windows1900, out var converted) && Math.Abs(calendar.Serial - converted.Serial) < tolerance) system = ExcelDateSystem.Mac1904;
                        else throw new InvalidDataException("XML 日期与 Biff8 原始值不一致。");
                        value = CellValue.FromNumber(rawValue); date = IsTimeOnlyFormat(style, rawValue) ? null : new(rawValue, system);
                    }
                    else if (IsTimeOnlyFormat(style, calendar.Serial))
                        throw new NotSupportedException("时间/累计时长缺少 Biff8 原始数值，无法确认源日期系统；请在 Excel 重新复制或明确转换为数字。");
                }
                cells[ri * columns + ci] = new(value, style, a1, origin, formula is not null && data is null, date, rejection);
                ci += across + 1;
            }
            ri++;
        }
        // A WorkbookSnapshot validates all supported style properties, including sizes and rotation, before use.
        _ = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new WorksheetSnapshot(SheetId.New(), "Validation")], styles: cells.Select(c => c.Style!).Distinct());
        return new(rows, columns, cells, merges, limitations, hasFilterMetadata: hasFilterMetadata);
    }
    private static bool IsTimeOnlyFormat(StyleDefinition style, double number) => CellDisplayFormatter.IsTimeOnlyFormat(style, number);

    private static (CellValue Value, ExcelSerialDate? Date) ReadValue(XElement? data)
    {
        if (data is null) return (CellValue.Blank, null);
        var text = data.Value;
        if (text.Length > WorkbookLimits.MaxCellTextLength) throw new InvalidDataException("剪贴板单元格文本过长。");
        switch (A(data, "Type"))
        {
            case "String": return (CellValue.FromText(text), null);
            case "Number":
                var number = double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
                if (!double.IsFinite(number)) throw new FormatException("剪贴板数值不是有限数。");
                return (CellValue.FromNumber(number), null);
            case "Boolean": return (CellValue.FromBoolean(text switch { "1" => true, "0" => false, _ => throw new FormatException("无效布尔值。") }), null);
            case "Error":
                var error = Enum.GetValues<CellErrorCode>().FirstOrDefault(e => CellInput.DisplayText(CellValue.FromError(e)) == text, (CellErrorCode)(-1));
                if (!Enum.IsDefined(error)) throw new NotSupportedException("未知的 Excel 错误类型。");
                return (CellValue.FromError(error), null);
            case "DateTime":
                var leap = text.StartsWith("1900-02-29T", StringComparison.Ordinal);
                var normalized = leap ? "1900-02-28" + text[10..] : text;
                if (!DateTime.TryParseExact(normalized, ["yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", "yyyy-MM-dd'T'HH:mm:ss"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) throw new FormatException("无效或含时区的 Excel 日期。");
                var serial = ExcelSerialDate.FromDateTime(DateTime.SpecifyKind(date, DateTimeKind.Unspecified), ExcelDateSystem.Windows1900);
                if (leap) serial = new(serial.Serial + 1, ExcelDateSystem.Windows1900);
                return (CellValue.FromNumber(serial.Serial), serial);
            default: throw new NotSupportedException("未知 XML Spreadsheet 数据类型。");
        }
    }

    public static ExternalClipboardRegion Capture(WorkbookSnapshot snapshot, WorkbookClipboard capture, bool formulas = false, bool valuesOnly = false)
    {
        if (capture.DocumentId != snapshot.DocumentId || capture.Revision != snapshot.Revision) throw new InvalidOperationException("复制捕获已过期。");
        var sheet = snapshot.GetSheet(capture.SheetId);
        if (!valuesOnly && (capture.HasAnnotations || capture.HasConditionalFormats || capture.Cells.Any(c => c.Content.RichText is not null) ||
            sheet.Data.Tables.Any(t => Intersects(t.Range, capture.Range)) || sheet.Data.Validations.Any(v => v.Ranges.Any(r => Intersects(r, capture.Range)))))
            throw new NotSupportedException("来源含富文本、附件或规则；对外复制请明确选择“复制结果值”。");
        if (capture.HasOldResults) throw new InvalidOperationException("源公式结果缺失或已过期，请先完成重算再对外复制。");
        var cells = capture.Cells.Select((c, i) =>
        {
            var style = snapshot.Styles[c.Content.StyleId] with { FontScheme = null, NamedStyleId = 0 };
            if (style.TextRotation == 255) throw new NotSupportedException("竖排文字尚不支持对外复制。");
            var formula = formulas && c.Content.Kind == CellContentKind.Formula ? c.Content.FormulaText : null;
            if (formula is not null) ExternalClipboardFormula.Validate(formula);
            var date = c.Value.Kind == CellValueKind.Number && CellDisplayFormatter.IsDateFormat(style, c.Value.NumberValue)
                ? new ExcelSerialDate(c.Value.NumberValue, snapshot.DateSystem) : (ExcelSerialDate?)null;
            return new ExternalClipboardCell(c.Value, valuesOnly ? new() { NumberFormatCode = c.Value.Kind == CellValueKind.Text ? "@" : "General" } : style,
                formula, new(capture.Range.Top + i / capture.Columns, capture.Range.Left + i % capture.Columns), Date: date);
        });
        return new(capture.Rows, capture.Columns, cells, valuesOnly ? [] : capture.MergedCells.Select(m => new CellRectangle(m.Top - capture.Range.Top, m.Left - capture.Range.Left, m.Bottom - capture.Range.Top, m.Right - capture.Range.Left)));
    }
    private static bool Intersects(CellRectangle a, CellRectangle b) => a.Top <= b.Bottom && a.Bottom >= b.Top && a.Left <= b.Right && a.Right >= b.Left;

    public static byte[] Write(ExternalClipboardRegion region, bool formulas = false)
    {
        region.RequireExport(formulas);
        var styles = region.Cells.Select(c => c.Style!).Distinct().ToArray();
        if (styles.Length > MaxStyles) throw new InvalidDataException("剪贴板样式超过上限。");
        var styleIds = styles.Select((s, i) => (s, id: "s" + i)).ToDictionary(p => p.s, p => p.id);
        var table = new XElement(Ss + "Table", new XAttribute(Ss + "ExpandedRowCount", region.Rows), new XAttribute(Ss + "ExpandedColumnCount", region.Columns));
        var covered = new HashSet<CellPosition>(); var anchors = region.Merges.ToDictionary(m => new CellPosition(m.Top, m.Left));
        foreach (var m in region.Merges) for (var r = m.Top; r <= m.Bottom; r++) for (var c = m.Left; c <= m.Right; c++) if (r != m.Top || c != m.Left) covered.Add(new(r, c));
        for (var r = 0; r < region.Rows; r++)
        {
            var row = new XElement(Ss + "Row"); table.Add(row);
            for (var c = 0; c < region.Columns; c++)
            {
                if (covered.Contains(new(r, c))) continue;
                var cell = region.Cells[r * region.Columns + c];
                var e = new XElement(Ss + "Cell", new XAttribute(Ss + "Index", c + 1), new XAttribute(Ss + "StyleID", styleIds[cell.Style!]));
                if (anchors.TryGetValue(new(r, c), out var m)) { e.Add(new XAttribute(Ss + "MergeAcross", m.Right - m.Left), new XAttribute(Ss + "MergeDown", m.Bottom - m.Top)); }
                if (formulas && cell.Formula is { } f) e.Add(new XAttribute(Ss + "Formula", SpreadsheetClipboardFormula.Write(f, cell.FormulaOrigin)));
                if (cell.Value.Kind != CellValueKind.Blank)
                {
                    var type = cell.Value.Kind switch { CellValueKind.Text => "String", CellValueKind.Number => "Number", CellValueKind.Boolean => "Boolean", _ => "Error" };
                    var value = cell.Value.Kind == CellValueKind.Boolean ? (cell.Value.BooleanValue ? "1" : "0") : CellInput.DisplayText(cell.Value);
                    if (cell.Date is { } date)
                    {
                        type = "DateTime";
                        if (date.IsFictitiousLeapDay) value = "1900-02-29T" + DateTime.MinValue.AddDays(date.Serial - Math.Floor(date.Serial)).ToString("HH:mm:ss.fffffff", CultureInfo.InvariantCulture);
                        else if (date.TryToDateTime(out var actual)) value = actual.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", CultureInfo.InvariantCulture);
                        else throw new NotSupportedException("日期超出可互通范围。");
                    }
                    e.Add(new XElement(Ss + "Data", new XAttribute(Ss + "Type", type), value));
                }
                row.Add(e);
            }
        }
        var root = new XElement(Ss + "Workbook", new XAttribute(XNamespace.Xmlns + "ss", Ss),
            new XElement(Ss + "Styles", styles.Select((s, i) => SpreadsheetClipboardStyles.Write(s, "s" + i))),
            new XElement(Ss + "Worksheet", new XAttribute(Ss + "Name", "Clipboard"), table));
        // XmlWriter entitizes CR; literal CR/CRLF would be normalized by every XML reader.
        using var buffer = new MemoryStream();
        using (var writer = XmlWriter.Create(buffer, new XmlWriterSettings { Encoding = Utf8, NewLineHandling = NewLineHandling.Entitize, CloseOutput = false })) root.Save(writer);
        var encoded = Utf8.GetString(buffer.ToArray()).Replace("\n", "&#10;").Replace("\t", "&#9;");
        var bytes = Utf8.GetBytes(encoded + "\0");
        if (bytes.Length > MaxBytes) throw new InvalidDataException("XML 剪贴板超过 8 MiB。");
        return bytes;
    }
}
