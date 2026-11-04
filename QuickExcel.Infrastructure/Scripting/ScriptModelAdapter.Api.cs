using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private ScriptWorkbookModel Model => request.Model!;
    private static readonly CellRectangle ScalarRange = new(0, 0, 0, 0);
    private static object Handle(string type, object data) => new { type, data, scriptHandle = true };
    private static object RangeHandle(SheetId sheet, CellRectangle range) => Handle("Range", new { sheetId = sheet.ToString(), address = range.ToString() });
    private static object TableHandle(SheetId sheet, int table) => Handle("Table", new { sheetId = sheet.ToString(), tableId = table });
    private static string Text(JsonElement value) => value.ValueKind == JsonValueKind.String ? value.GetString()! : throw Invalid("需要字符串。");
    private static int Int(JsonElement value) => value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var n) ? n : throw Invalid("需要整数。");
    private static bool Bool(JsonElement value) => value.ValueKind is JsonValueKind.True or JsonValueKind.False ? value.GetBoolean() : throw Invalid("需要布尔值。");
    private static double Number(JsonElement value) => value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var n) && double.IsFinite(n) ? n : throw Invalid("需要有限数值。");
    private static ModelFailure Invalid(string message) => new("InvalidArgument", message);
    private static ModelFailure Unsupported(string message) => new("UnsupportedApi", message);
    private static void Arity(JsonElement[] args, int min, int? max = null)
    { if (args.Length < min || args.Length > (max ?? min)) throw Invalid("参数个数不正确。"); }
    private static string Argb(JsonElement value)
    {
        var color = Text(value);
        if (color.Length != 7 || color[0] != '#' || color[1..].Any(c => !char.IsAsciiHexDigit(c))) throw Invalid("颜色必须是 #RRGGBB。");
        return "FF" + color[1..].ToUpperInvariant();
    }
    private static string Rgb(string? color) => color is { Length: 8 } ? "#" + color[2..] : "";
    private static CellRectangle Address(JsonElement value, SheetId sheet)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            if (value.GetProperty("sheetId").GetString() != sheet.ToString()) throw Invalid("区域必须属于同一工作表。");
            value = value.GetProperty("address");
        }
        var text = Text(value);
        if (text.Length > 40 || !System.Text.RegularExpressions.Regex.IsMatch(text, "^[A-Za-z]+[1-9][0-9]*(?::[A-Za-z]+[1-9][0-9]*)?$"))
            throw Invalid("需要 A1 单元格或矩形区域。");
        var range = CellRectangle.Parse(text);
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Top > range.Bottom || range.Left > range.Right) throw Invalid("区域起止顺序颠倒。");
        return range;
    }
    private void Write(SheetId sheet, CellRectangle range, DocumentCommand command)
    {
        if (range.Count > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "单次模型写入超过 20000 格。");
        Spend(checked((int)range.Count), true); Model.ExecuteCommand(sheet, range, command);
    }
    private static IEnumerable<CellAddress> Cells(SheetId sheet, CellRectangle range)
    {
        for (var r = range.Top; r <= range.Bottom; r++)
        for (var c = range.Left; c <= range.Right; c++) yield return new(sheet, r, c);
    }
    private void Style(SheetId sheet, CellRectangle range, CellStylePatch patch) => Write(sheet, range, new SetCellStyleCommand(Cells(sheet, range), patch));
    private object? Uniform(WorksheetSnapshot sheet, CellRectangle range, Func<StyleDefinition, object?> read)
    {
        if (range.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "读取范围超过预算。");
        Spend((int)range.Count, false); object? first = null; var seen = false;
        foreach (var cell in Cells(sheet.Id, range))
        {
            var value = read(TablePresentation.GetStyle(Model.View, sheet, cell.Row, cell.Column));
            if (seen && !Equals(first, value)) return null;
            first = value; seen = true;
        }
        return first;
    }

    private object? CallApi(string method, JsonElement[] transport)
    {
        var name = method[7..]; // getApi: / setApi:
        if (transport.Length != 2 || transport[0].ValueKind != JsonValueKind.Object || transport[1].ValueKind != JsonValueKind.Array)
            throw Invalid("对象调用参数无效。");
        var data = transport[0]; var args = transport[1].EnumerateArray().ToArray();
        if (args.Length > 8) throw Invalid("对象调用参数过多。");
        if (name.StartsWith("Workbook.", StringComparison.Ordinal)) return WorkbookApi(name[9..], args);
        if (!Guid.TryParse(Text(data.GetProperty("sheetId")), out var id)) throw Invalid("工作表 ID 无效。");
        var sheet = Model.View.Sheets.FirstOrDefault(s => s.Id.Value == id) ?? throw new ModelFailure("InvalidReference", "工作表句柄已失效。");
        if (name.StartsWith("Worksheet.", StringComparison.Ordinal)) return WorksheetApi(sheet, name[10..], args);
        if (name.StartsWith("WorksheetProtection.", StringComparison.Ordinal)) return ProtectionApi(sheet, name[20..], args);
        if (name.StartsWith("Chart", StringComparison.Ordinal)) return ChartApi(sheet, data, name, args);
        if (name.StartsWith("Comment", StringComparison.Ordinal)) return CommentApi(sheet, data, name.Split('.')[1], args);
        if (data.TryGetProperty("conditionalId", out _)) return ConditionalApi(sheet, data, name, args);
        if (name.StartsWith("Table", StringComparison.Ordinal) || name.StartsWith("AutoFilter.", StringComparison.Ordinal) || name.StartsWith("Filter.", StringComparison.Ordinal))
            return TableApi(sheet, data, name, args);
        var range = Address(data.GetProperty("address"), sheet.Id);
        var split = name.Split('.'); var type = split[0]; var member = split[1];
        if (type == "Range") return RangeApi(sheet, range, member, args);
        if (type == "DataValidation") return ValidationApi(sheet, range, member, args);
        if (type == "RangeSort") { ApplySort(sheet, range, args, false); return null; }
        if (type == "RangeBorder") return BorderApi(sheet, range, Int(data.GetProperty("side")), member, args);
        if (type == "RangeFormat")
        {
            if (member is "getFont" or "getProtection")
            { Arity(args, 0); return Handle(member == "getFont" ? "RangeFont" : "FormatProtection", data); }
            if (member is "getBorders" or "getRangeBorder")
            {
                Arity(args, member == "getBorders" ? 0 : 1);
                object Border(int side) => Handle("RangeBorder", new { sheetId = sheet.Id.ToString(), address = range.ToString(), side });
                if (member == "getBorders") return Enumerable.Range(0, 6).Select(Border).ToArray();
                var side = Int(args[0]); if (side is < 0 or > 5) throw Unsupported("对角线边框尚未建模。");
                return Border(side);
            }
            if (member is "getRowHeight" or "getColumnWidth" or "setRowHeight" or "setColumnWidth")
            {
                var row = member.EndsWith("RowHeight", StringComparison.Ordinal); var write = member.StartsWith("set", StringComparison.Ordinal);
                Arity(args, write ? 1 : 0);
                if (write)
                {
                    var points = Number(args[0]); if (points <= 0) throw Invalid("尺寸必须大于零。");
                    Write(sheet.Id, range, new SetAxisMetadataCommand(sheet.Id, row ? SheetAxis.Row : SheetAxis.Column,
                        row ? range.Top : range.Left, row ? range.Bottom - range.Top + 1 : range.Right - range.Left + 1, points * 4 / 3));
                    return null;
                }
                Spend(row ? range.Bottom - range.Top + 1 : range.Right - range.Left + 1, false);
                var sizes = Enumerable.Range(row ? range.Top : range.Left, row ? range.Bottom - range.Top + 1 : range.Right - range.Left + 1)
                    .Select(i => row ? sheet.Rows.GetValueOrDefault(i)?.Height ?? sheet.DefaultRowHeight : sheet.Columns.GetValueOrDefault(i)?.Width ?? sheet.Data.DefaultColumnWidth ?? 64).Distinct().ToArray();
                return sizes.Length == 1 ? sizes[0] * 3 / 4 : null;
            }
        }
        var setter = member.StartsWith("set", StringComparison.Ordinal);
        Arity(args, setter ? 1 : 0);
        if (!setter)
        {
            return Uniform(sheet, range, s => (type, member) switch
            {
                ("RangeFont", "getBold") => s.Bold, ("RangeFont", "getItalic") => s.Italic,
                ("RangeFont", "getColor") => Rgb(s.FontColor), ("RangeFont", "getName") => s.FontName ?? "Calibri",
                ("RangeFont", "getSize") => s.FontSize ?? 11, ("RangeFont", "getStrikethrough") => s.Strike,
                ("RangeFont", "getUnderline") => s.Underline switch { "single" => 1, "double" => 2, "singleAccounting" => 3, "doubleAccounting" => 4, _ => 0 },
                ("RangeFormat", "getHorizontalAlignment") => s.Horizontal, ("RangeFormat", "getVerticalAlignment") => s.Vertical,
                ("RangeFormat", "getWrapText") => s.WrapText, ("RangeFormat", "getShrinkToFit") => s.ShrinkToFit,
                ("RangeFormat", "getIndentLevel") => s.Indent, ("RangeFormat", "getTextOrientation") => s.TextRotation == 255 ? 180 : s.TextRotation > 90 ? 90 - s.TextRotation : s.TextRotation,
                ("FormatProtection", "getLocked") => s.Locked, ("FormatProtection", "getFormulaHidden") => s.FormulaHidden,
                _ => throw Unsupported(name)
            });
        }
        var v = args[0];
        CellStylePatch patch = (type, member) switch
        {
            ("RangeFont", "setBold") => new(Bold: Bool(v)), ("RangeFont", "setItalic") => new(Italic: Bool(v)),
            ("RangeFont", "setColor") => new(FontColor: Argb(v)), ("RangeFont", "setName") => new(FontName: Text(v)),
            ("RangeFont", "setSize") => new(FontSize: Number(v)), ("RangeFont", "setStrikethrough") => new(Strike: Bool(v)),
            ("RangeFont", "setUnderline") => new(Underline: Int(v) switch { 0 => "none", 1 => "single", 2 => "double", 3 => "singleAccounting", 4 => "doubleAccounting", _ => throw Invalid("下划线枚举无效。") }),
            ("RangeFormat", "setHorizontalAlignment") => new(Horizontal: Int(v)), ("RangeFormat", "setVerticalAlignment") => new(Vertical: Int(v)),
            ("RangeFormat", "setWrapText") => new(WrapText: Bool(v)), ("RangeFormat", "setShrinkToFit") => new(ShrinkToFit: Bool(v)),
            ("RangeFormat", "setIndentLevel") => new(Indent: Int(v)), ("RangeFormat", "setTextOrientation") => new(TextRotation: Orientation(Int(v))),
            ("FormatProtection", "setLocked") => new(Locked: Bool(v)), ("FormatProtection", "setFormulaHidden") => new(FormulaHidden: Bool(v)),
            _ => throw Unsupported(name)
        };
        Style(sheet.Id, range, patch); return null;
    }
    private static int Orientation(int n) => n == 180 ? 255 : n is < -90 or > 90 ? throw Invalid("文本方向必须在 -90 至 90，或为 180。") : n < 0 ? 90 - n : n;
}
