using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? WorkbookApi(string member, JsonElement[] args)
    {
        if (member.Contains("Comment", StringComparison.Ordinal)) return Comments(null, member, args);
        if (member is "getTables" or "getTable")
        {
            Arity(args, member == "getTable" ? 1 : 0); Spend(1, false);
            var tables = Model.View.Sheets.SelectMany(s => s.Data.Tables.Select(t => (Sheet: s, Table: t)));
            if (member == "getTables") return tables.Select(t => TableHandle(t.Sheet.Id, t.Table.Id)).ToArray();
            var key = Text(args[0]); var item = tables.FirstOrDefault(t => t.Table.Name.Equals(key, StringComparison.OrdinalIgnoreCase) || t.Table.Id.ToString() == key);
            return item.Table is null ? null : TableHandle(item.Sheet.Id, item.Table.Id);
        }
        throw Unsupported(member);
    }
    private object? WorksheetApi(WorksheetSnapshot sheet, string member, JsonElement[] args)
    {
        if (member.Contains("Comment", StringComparison.Ordinal)) return Comments(sheet, member, args);
        if (member is "getCharts" or "getChart" or "addChart") return WorksheetCharts(sheet, member, args);
        switch (member)
        {
            case "getProtection": case "getAutoFilter":
                Arity(args, 0); return Handle(member == "getProtection" ? "WorksheetProtection" : "AutoFilter", new { sheetId = sheet.Id.ToString() });
            case "getUsedRange":
                Arity(args, 0, 1); var valuesOnly = args.Length == 1 && Bool(args[0]);
                if (sheet.Cells.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "已用区域扫描超过读取预算。");
                Spend(Math.Max(1, sheet.Cells.Count), false);
                var positions = sheet.Cells.Where(p => !valuesOnly || p.Value.Content.Kind == CellContentKind.Formula || p.Value.Content.LiteralValue.Kind != CellValueKind.Blank).Select(p => p.Key).ToArray();
                return positions.Length == 0 ? null : RangeHandle(sheet.Id, new(positions.Min(p => p.Row), positions.Min(p => p.Column), positions.Max(p => p.Row), positions.Max(p => p.Column)));
            case "getTables": case "getTable":
                Arity(args, member == "getTable" ? 1 : 0); Spend(1, false);
                if (member == "getTables") return sheet.Data.Tables.Select(t => TableHandle(sheet.Id, t.Id)).ToArray();
                var key = Text(args[0]); var table = sheet.Data.Tables.FirstOrDefault(t => t.Name.Equals(key, StringComparison.OrdinalIgnoreCase) || t.Id.ToString() == key);
                return table is null ? null : TableHandle(sheet.Id, table.Id);
            case "addTable":
                Arity(args, 2); var range = Address(args[0], sheet.Id);
                if (!Bool(args[1])) throw Unsupported("创建无表头表尚未支持；请显式提供唯一文本表头。");
                var names = Model.View.Sheets.SelectMany(s => s.Data.Tables).Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var index = 1; while (names.Contains("Table" + index)) index++;
                Write(sheet.Id, range, new CreateTableCommand(sheet.Id, range, "Table" + index, reservedId: Model.ReserveTableId()));
                return TableHandle(sheet.Id, Model.View.GetSheet(sheet.Id).Data.Tables.Single(t => t.Name == "Table" + index).Id);
            default: throw Unsupported(member);
        }
    }
    private object? RangeApi(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        if (member == "getDataValidation") { Arity(args, 0); return Handle("DataValidation", new { sheetId = sheet.Id.ToString(), address = range.ToString() }); }
        if (member.Contains("ConditionalFormat", StringComparison.Ordinal)) return RangeConditional(sheet, range, member, args);
        if (member == "clear")
        {
            Arity(args, 0, 1); var mode = args.Length == 0 ? 0 : Int(args[0]);
            if (mode is < 0 or > 2) throw Unsupported("clear 目前支持 all、formats、contents。");
            Write(sheet.Id, range, new ClearScriptRangeCommand(sheet.Id, range, mode)); return null;
        }
        if (member is "getRowIndex" or "getColumnIndex" or "getRowCount" or "getColumnCount" or "getCellCount" or "getSort")
        {
            Arity(args, 0);
            return member switch { "getRowIndex" => range.Top, "getColumnIndex" => range.Left, "getRowCount" => range.Bottom - range.Top + 1,
                "getColumnCount" => range.Right - range.Left + 1, "getCellCount" => range.Count, _ => Handle("RangeSort", new { sheetId = sheet.Id.ToString(), address = range.ToString() }) };
        }
        if (member is "setNumberFormat" or "setNumberFormats")
        {
            Arity(args, 1);
            if (member == "setNumberFormat") { Style(sheet.Id, range, new(NumberFormatId: 0, NumberFormatCode: FormatCode(args[0]))); return null; }
            var rows = Matrix(args[0], range.Bottom - range.Top + 1, range.Right - range.Left + 1);
            // Validate the complete input before applying the first patch; repeated styles are grouped.
            var edits = Cells(sheet.Id, range).Select(a => (Address: a, Format: FormatCode(rows[a.Row - range.Top][a.Column - range.Left]))).ToArray();
            Write(sheet.Id, range, new SetScriptNumberFormatsCommand(edits.Select(e => (e.Address, e.Format)))); return null;
        }
        if (member is "getNumberFormat" or "getNumberFormats" or "getNumberFormatCategory" or "getNumberFormatCategories" or "getFormulas" or "getValueType" or "getValueTypes" or "getText" or "getTexts")
        {
            Arity(args, 0); var scalar = member is "getNumberFormat" or "getNumberFormatCategory" or "getValueType" or "getText";
            var target = scalar ? new CellRectangle(range.Top, range.Left, range.Top, range.Left) : range;
            if (target.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "读取超过预算。");
            Spend((int)target.Count, false);
            if (member is "getValueType" or "getValueTypes" or "getText" or "getTexts") Model.PrepareValueRead(sheet.Id, target, cancellation);
            var result = new List<object?[]>();
            for (var r = target.Top; r <= target.Bottom; r++)
            {
                cancellation.ThrowIfCancellationRequested(); var row = new List<object?>();
                for (var c = target.Left; c <= target.Right; c++)
                {
                    var cell = Model.ReadCell(sheet.Id, new(r, c), calculated: true);
                    var style = TablePresentation.NumberStyle(sheet, Model.View.Styles, r, c); var value = cell?.Value ?? CellValue.Blank;
                    if (member is "getValueType" or "getValueTypes" or "getText" or "getTexts") EnsureReadable(cell);
                    row.Add(member switch
                    {
                        "getNumberFormat" or "getNumberFormats" => style.NumberFormatCode ?? CellDisplayFormatter.BuiltIn(style.NumberFormatId) ?? "General",
                        "getNumberFormatCategory" or "getNumberFormatCategories" => FormatCategory(style),
                        "getFormulas" => cell?.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText : Scalar(value),
                        "getText" or "getTexts" => CellDisplayFormatter.Format(value, style, Model.View.DateSystem),
                        _ => value.Kind switch { CellValueKind.Blank => 1, CellValueKind.Text => 2, CellValueKind.Number => Math.Truncate(value.NumberValue) == value.NumberValue ? 3 : 4, CellValueKind.Boolean => 5, CellValueKind.Error => 6, _ => 0 }
                    });
                }
                result.Add(row.ToArray());
            }
            return scalar ? result[0][0] : result.ToArray();
        }
        throw Unsupported(member);
    }
    private static string FormatCode(JsonElement input)
    {
        var code = Text(input); if (code.Length is 0 or > 32767 || code.Any(c => c is '\0' or '\r' or '\n')) throw Invalid("数字格式代码无效。");
        return code;
    }
    private static int FormatCategory(StyleDefinition style)
    {
        var id = style.NumberFormatCode is { } code ? Enumerable.Range(0, 50).FirstOrDefault(i => CellDisplayFormatter.BuiltIn(i) == code, -1) : style.NumberFormatId;
        return id switch { 0 => 0, >= 1 and <= 4 or >= 37 and <= 40 => 1, >= 5 and <= 8 => 2, >= 41 and <= 44 => 3,
            >= 14 and <= 17 or 22 => 4, >= 18 and <= 21 or >= 45 and <= 47 => 5, 9 or 10 => 6, 12 or 13 => 7, 11 or 48 => 8, 49 => 9, _ => 11 };
    }
    private object? BorderApi(WorksheetSnapshot sheet, CellRectangle range, int side, string member, JsonElement[] args)
    {
        if (side is < 0 or > 5) throw Unsupported("对角线边框尚未支持。");
        var setter = member.StartsWith("set", StringComparison.Ordinal); Arity(args, setter ? 1 : 0);
        if (member == "getSideIndex") return side;
        // Internal borders apply to both sides of the boundary; the first edge is the read projection.
        var selected = side switch { 0 => range with { Bottom = range.Top }, 1 => range with { Top = range.Bottom },
            2 => range with { Right = range.Left }, 3 => range with { Left = range.Right },
            4 => range with { Right = range.Right - 1 }, _ => range with { Bottom = range.Bottom - 1 } };
        if (selected.Top > selected.Bottom || selected.Left > selected.Right) return setter ? null : member == "getColor" ? "" : 0;
        int Line(StyleDefinition s) => side switch { 0 => s.TopBorder, 1 => s.BottomBorder, 2 => s.LeftBorder, _ when side == 3 || side == 4 => s.RightBorder, _ => s.BottomBorder };
        string? Color(StyleDefinition s) => side switch { 0 => s.TopBorderColor, 1 => s.BottomBorderColor, 2 => s.LeftBorderColor, _ when side == 3 || side == 4 => s.RightBorderColor, _ => s.BottomBorderColor };
        if (!setter) return Uniform(sheet, selected, s => member switch
        {
            "getColor" => Rgb(Color(s)), "getStyle" => Line(s) switch { 0 => 0, 3 => 2, 4 => 5, 6 => 6, _ => 1 },
            "getWeight" => Line(s) switch { 7 => 0, 2 => 2, 5 => 3, _ => 1 }, _ => throw Unsupported(member)
        });
        int? style = null; string? color = null; int? weight = null;
        if (member == "setColor") color = Argb(args[0]);
        else if (member == "setStyle") style = Int(args[0]) switch { 0 => 0, 1 => 1, 2 => 3, 5 => 4, 6 => 6, _ => throw Unsupported("此边框线型尚未建模。") };
        else if (member == "setWeight") weight = Int(args[0]) switch { 0 => 7, 1 => 1, 2 => 2, 3 => 5, _ => throw Invalid("边框粗细枚举无效。") };
        else throw Unsupported(member);
        Write(sheet.Id, range, new SetScriptBorderCommand(sheet.Id, range, side, style, color, weight)); return null;
    }
}
