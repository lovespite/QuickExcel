using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private static object ConditionalHandle(SheetId sheet, ConditionalFormatRule rule, string type = "ConditionalFormat") => Handle(type, new { sheetId = sheet.ToString(), conditionalId = rule.ScriptId.ToString() });
    private object? RangeConditional(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        var rules = sheet.Data.ConditionalFormats.Where(r => r.Ranges.Any(a => Overlap(a, range))).OrderBy(r => r.Priority).ToArray();
        Spend(Math.Max(1, sheet.Data.ConditionalFormats.Length), false);
        if (member == "getConditionalFormats") { Arity(args, 0); return rules.Select(r => ConditionalHandle(sheet.Id, r)).ToArray(); }
        if (member == "getConditionalFormat") { Arity(args, 1); var found = rules.FirstOrDefault(r => r.ScriptId.ToString() == Text(args[0])); return found is null ? null : ConditionalHandle(sheet.Id, found); }
        if (member == "clearAllConditionalFormats")
        {
            Arity(args, 0);
            if (rules.Any(r => r.Ranges.Any(a => !range.Contains(a.Top, a.Left) || !range.Contains(a.Bottom, a.Right)))) throw Unsupported("清除区域只覆盖条件格式的一部分，请使用完整规则范围。");
            WriteRules(sheet, rules.SelectMany(r => r.Ranges), new SetConditionalFormatsCommand(sheet.Id, sheet.Data.ConditionalFormats.Except(rules))); return null;
        }
        Arity(args, 1); var kind = EnumName("ConditionalFormatType", args[0]) switch
        { "custom" => ConditionalFormatKind.Expression, "cellValue" => ConditionalFormatKind.CellIs, "presetCriteria" => ConditionalFormatKind.DuplicateValues, "containsText" => ConditionalFormatKind.ContainsText, _ => throw Unsupported("此条件格式类型尚未建模。") };
        var priority = sheet.Data.ConditionalFormats.Select(r => r.Priority).DefaultIfEmpty().Max() + 1;
        var added = new ConditionalFormatRule([range], kind, priority, new(), kind is ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues ? null : "=0", comparison: ConditionalComparison.Equal, text: kind == ConditionalFormatKind.ContainsText ? "" : null);
        WriteRules(sheet, [range], new SetConditionalFormatsCommand(sheet.Id, sheet.Data.ConditionalFormats.Add(added))); return ConditionalHandle(sheet.Id, added);
    }
    private static string ComparisonName(ConditionalComparison comparison) => comparison switch
    { ConditionalComparison.Equal => "equalTo", ConditionalComparison.NotEqual => "notEqualTo", ConditionalComparison.GreaterThan => "greaterThan", ConditionalComparison.LessThan => "lessThan", ConditionalComparison.GreaterThanOrEqual => "greaterThanOrEqual", ConditionalComparison.LessThanOrEqual => "lessThanOrEqual", ConditionalComparison.Between => "between", _ => "notBetween" };
    private static ConditionalComparison Comparison(string name) => name switch
    { "between" => ConditionalComparison.Between, "notBetween" => ConditionalComparison.NotBetween, "equalTo" => ConditionalComparison.Equal, "notEqualTo" => ConditionalComparison.NotEqual, "lessThan" => ConditionalComparison.LessThan, "lessThanOrEqual" or "lessThanOrEqualTo" => ConditionalComparison.LessThanOrEqual, "greaterThan" => ConditionalComparison.GreaterThan, "greaterThanOrEqual" or "greaterThanOrEqualTo" => ConditionalComparison.GreaterThanOrEqual, _ => throw Invalid("不支持的比较运算符。") };
    private object? ConditionalApi(WorksheetSnapshot sheet, JsonElement data, string name, JsonElement[] args)
    {
        var rule = sheet.Data.ConditionalFormats.FirstOrDefault(r => r.ScriptId.ToString() == Text(data.GetProperty("conditionalId"))) ?? throw new ModelFailure("InvalidReference", "条件格式句柄已失效。");
        var type = name.Split('.')[0]; var member = name.Split('.')[1];
        object Child(string t) => ConditionalHandle(sheet.Id, rule, t);
        ConditionalFormatRule Copy(DifferentialStyle? style = null, bool? stop = null, IEnumerable<CellRectangle>? ranges = null) => new(ranges ?? rule.Ranges, rule.Kind, rule.Priority, style ?? rule.Style, rule.Formula1, rule.Formula2, rule.Comparison, stop ?? rule.StopIfTrue, rule.Text, rule.ColorScale, rule.DataBar, rule.IconSet) { ScriptId = rule.ScriptId };
        void Save(ConditionalFormatRule next) => WriteRules(sheet, rule.Ranges.Concat(next.Ranges).Distinct(), new SetConditionalFormatsCommand(sheet.Id, sheet.Data.ConditionalFormats.Replace(rule, next)));
        if (type == "ConditionalFormat")
        {
            if (member == "delete") { Arity(args, 0); WriteRules(sheet, rule.Ranges, new SetConditionalFormatsCommand(sheet.Id, sheet.Data.ConditionalFormats.Remove(rule))); return null; }
            if (member == "setPriority")
            {
                Arity(args, 1); var at = Int(args[0]); if (at < 1 || at > sheet.Data.ConditionalFormats.Length) throw Invalid("优先级越界。");
                var ordered = sheet.Data.ConditionalFormats.OrderBy(r => r.Priority).Where(r => r != rule).ToList(); ordered.Insert(at - 1, rule);
                WriteRules(sheet, rule.Ranges, new SetConditionalFormatsCommand(sheet.Id, ordered.Select((r, i) => r.WithPriority(i + 1)))); return null;
            }
            if (member == "setRanges") { Arity(args, 1); Save(Copy(ranges: [Address(args[0], sheet.Id)])); return null; }
            if (member == "setStopIfTrue") { Arity(args, 1); Save(Copy(stop: Bool(args[0]))); return null; }
            Arity(args, 0); Spend(1, false);
            return member switch
            {
                "getId" => rule.ScriptId.ToString(), "getPriority" => rule.Priority, "getStopIfTrue" => rule.StopIfTrue,
                "getType" => EnumNumber("ConditionalFormatType", rule.Kind switch { ConditionalFormatKind.CellIs => "cellValue", ConditionalFormatKind.Expression => "custom", ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues => "presetCriteria", ConditionalFormatKind.ColorScale => "colorScale", ConditionalFormatKind.DataBar => "dataBar", ConditionalFormatKind.IconSet => "iconSet", _ => "containsText" }),
                "getCustom" => rule.Kind == ConditionalFormatKind.Expression ? Child("CustomConditionalFormat") : null,
                "getCellValue" => rule.Kind == ConditionalFormatKind.CellIs ? Child("CellValueConditionalFormat") : null,
                "getPreset" => rule.Kind is ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues ? Child("PresetCriteriaConditionalFormat") : null,
                "getTextComparison" => rule.Kind is >= ConditionalFormatKind.ContainsText and <= ConditionalFormatKind.EndsWith ? Child("TextConditionalFormat") : null,
                "getRange" => rule.Ranges.Length == 1 ? RangeHandle(sheet.Id, rule.Ranges[0]) : throw Unsupported("多区域规则请使用 getRanges。"),
                "getRanges" => Handle("RangeAreas", new { sheetId = sheet.Id.ToString(), addresses = rule.Ranges.Select(r => r.ToString()).ToArray() }), _ => throw Unsupported(member)
            };
        }
        if ((type is "CustomConditionalFormat" or "ConditionalFormatRule") && rule.Kind != ConditionalFormatKind.Expression || type == "CellValueConditionalFormat" && rule.Kind != ConditionalFormatKind.CellIs || type == "PresetCriteriaConditionalFormat" && rule.Kind is not (ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues) || type == "TextConditionalFormat" && (rule.Kind < ConditionalFormatKind.ContainsText || rule.Kind > ConditionalFormatKind.EndsWith))
            throw new ModelFailure("InvalidReference", "条件格式规则类型已改变。");
        if (type == "ConditionalRangeFormat" && (member == "getBorders" || member.StartsWith("getConditionalRangeBorder")))
        {
            Arity(args, member == "getConditionalRangeBorder" ? 1 : 0);
            object Border(int side) => Handle("ConditionalRangeBorder", new { sheetId = sheet.Id.ToString(), conditionalId = rule.ScriptId.ToString(), side });
            if (member == "getBorders") return Enumerable.Range(0, 4).Select(Border).ToArray();
            var side = member switch { "getConditionalRangeBorderTop" => 0, "getConditionalRangeBorderBottom" => 1, "getConditionalRangeBorderLeft" => 2, "getConditionalRangeBorderRight" => 3, _ => Int(args[0]) };
            if (side is < 0 or > 3) throw Invalid("条件边框方向无效。"); return Border(side);
        }
        Arity(args, member.StartsWith("set") ? 1 : 0); Spend(1, false);
        if (type == "ConditionalRangeBorder")
        {
            var side = Int(data.GetProperty("side")); var border = (side switch { 0 => rule.Style.Top, 1 => rule.Style.Bottom, 2 => rule.Style.Left, _ => rule.Style.Right }) ?? new TableBorder(0, null);
            if (member == "getSideIndex") return side;
            if (member == "getColor") return Rgb(border.Color);
            if (member == "getStyle") return EnumNumber("ConditionalRangeBorderLineStyle", border.Style switch { 0 => "none", 1 or 2 or 5 or 7 => "continuous", 3 => "dash", 4 => "dot", _ => throw Unsupported("此已导入线型不在条件边框枚举中。") });
            var next = member switch { "setColor" => border with { Color = Argb(args[0]) }, "setStyle" => border with { Style = EnumName("ConditionalRangeBorderLineStyle", args[0]) switch { "none" => 0, "continuous" => 1, "dash" => 3, "dot" => 4, _ => throw Unsupported("此线型尚未建模。") } }, _ => throw Unsupported(member) };
            Save(Copy(side switch { 0 => rule.Style with { Top = next }, 1 => rule.Style with { Bottom = next }, 2 => rule.Style with { Left = next }, _ => rule.Style with { Right = next } })); return null;
        }
        if (member == "getFormat") return Child("ConditionalRangeFormat");
        if (member == "getFont") return Child("ConditionalRangeFont");
        if (member == "getFill") return Child("ConditionalRangeFill");
        if (type == "CustomConditionalFormat" && member == "getRule") return Child("ConditionalFormatRule");
        if (type == "ConditionalFormatRule")
        {
            if (member == "getFormula") return rule.Formula1;
            if (member == "setFormula") { Save(new(rule.Ranges, rule.Kind, rule.Priority, rule.Style, Text(args[0]), stopIfTrue: rule.StopIfTrue) { ScriptId = rule.ScriptId }); return null; }
        }
        if (type is "CellValueConditionalFormat" or "TextConditionalFormat" or "PresetCriteriaConditionalFormat")
        {
            if (member == "getRule")
            {
                if (type == "CellValueConditionalFormat") { var result = new Dictionary<string, object> { ["formula1"] = rule.Formula1!, ["operator"] = EnumNumber("ConditionalCellValueOperator", ComparisonName(rule.Comparison)) }; if (rule.Formula2 is not null) result["formula2"] = rule.Formula2; return result; }
                if (type == "PresetCriteriaConditionalFormat") return new { criterion = EnumNumber("ConditionalFormatPresetCriterion", rule.Kind == ConditionalFormatKind.UniqueValues ? "uniqueValues" : "duplicateValues") };
                return new { text = rule.Text, @operator = EnumNumber("ConditionalTextOperator", rule.Kind switch { ConditionalFormatKind.ContainsText => "contains", ConditionalFormatKind.NotContainsText => "notContains", ConditionalFormatKind.BeginsWith => "beginsWith", _ => "endsWith" }) };
            }
            if (member != "setRule") throw Unsupported(member);
            var v = args[0]; ConditionalFormatRule next;
            if (type == "CellValueConditionalFormat")
            {
                Fields(v, "formula1", "formula2", "operator"); next = new(rule.Ranges, rule.Kind, rule.Priority, rule.Style, Text(v.GetProperty("formula1")), v.TryGetProperty("formula2", out var f2) ? Text(f2) : null, Comparison(EnumName("ConditionalCellValueOperator", v.GetProperty("operator"))), rule.StopIfTrue);
            }
            else if (type == "PresetCriteriaConditionalFormat")
            {
                Fields(v, "criterion"); var kind = EnumName("ConditionalFormatPresetCriterion", v.GetProperty("criterion")) switch { "uniqueValues" => ConditionalFormatKind.UniqueValues, "duplicateValues" => ConditionalFormatKind.DuplicateValues, _ => throw Unsupported("目前支持唯一值和重复值预设。") };
                next = new(rule.Ranges, kind, rule.Priority, rule.Style, stopIfTrue: rule.StopIfTrue);
            }
            else
            {
                Fields(v, "operator", "text"); var text = Text(v.GetProperty("text")); var literal = "\"" + text.Replace("\"", "\"\"") + "\"";
                var address = new CellRectangle(rule.Anchor.Row, rule.Anchor.Column, rule.Anchor.Row, rule.Anchor.Column).ToString().Split(':')[0];
                var (kind, formula) = EnumName("ConditionalTextOperator", v.GetProperty("operator")) switch {
                    "contains" => (ConditionalFormatKind.ContainsText, $"=NOT(ISERROR(SEARCH({literal},{address})))"), "notContains" => (ConditionalFormatKind.NotContainsText, $"=ISERROR(SEARCH({literal},{address}))"),
                    "beginsWith" => (ConditionalFormatKind.BeginsWith, $"=LEFT({address},LEN({literal}))={literal}"), "endsWith" => (ConditionalFormatKind.EndsWith, $"=RIGHT({address},LEN({literal}))={literal}"), _ => throw Invalid("无效文本比较运算符。") };
                next = new(rule.Ranges, kind, rule.Priority, rule.Style, formula, stopIfTrue: rule.StopIfTrue, text: text);
            }
            next = new(next.Ranges, next.Kind, next.Priority, next.Style, next.Formula1, next.Formula2, next.Comparison, next.StopIfTrue, next.Text) { ScriptId = rule.ScriptId }; Save(next); return null;
        }
        var style = rule.Style;
        if (type == "ConditionalRangeFormat")
        {
            if (member == "getNumberFormat") return style.NumberFormatCode ?? CellDisplayFormatter.BuiltIn(style.NumberFormatId ?? 0) ?? "General";
            if (member == "setNumberFormat") style = style with { NumberFormatId = 0, NumberFormatCode = FormatCode(args[0]) };
            else if (member == "clearFormat") style = new(); else throw Unsupported(member);
        }
        else if (type == "ConditionalRangeFill")
        {
            if (member == "getColor") return Rgb(style.FillColor);
            style = member switch { "setColor" => style with { FillColor = Argb(args[0]) }, "clear" => style with { FillColor = null }, _ => throw Unsupported(member) };
        }
        else if (type == "ConditionalRangeFont")
        {
            if (member.StartsWith("get")) return member switch { "getBold" => style.Bold, "getItalic" => style.Italic, "getStrikethrough" => style.Strike, "getColor" => Rgb(style.FontColor), "getUnderline" => EnumNumber("ConditionalRangeFontUnderlineStyle", style.Underline ?? "none"), _ => throw Unsupported(member) };
            style = member switch {
                "setBold" => style with { Bold = Bool(args[0]) }, "setItalic" => style with { Italic = Bool(args[0]) }, "setStrikethrough" => style with { Strike = Bool(args[0]) },
                "setColor" => style with { FontColor = Argb(args[0]) }, "setUnderline" => style with { Underline = EnumName("ConditionalRangeFontUnderlineStyle", args[0]) },
                "clear" => style with { Bold = null, Italic = null, Strike = null, FontColor = null, Underline = null }, _ => throw Unsupported(member) };
        }
        else throw Unsupported(name);
        Save(Copy(style)); return null;
    }
}
