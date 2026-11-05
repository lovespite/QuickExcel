using System.Globalization;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? ValidationApi(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        var matches = sheet.Data.Validations.Where(v => v.Ranges.Any(r => Overlap(r, range))).ToArray();
        var exact = matches.Length == 1 && matches[0].Ranges.Length == 1 && matches[0].Ranges[0].Contains(range.Top, range.Left) && matches[0].Ranges[0].Contains(range.Bottom, range.Right);
        var rule = exact ? matches[0] : new WorksheetValidation([range], ValidationKind.None);
        var kindName = rule.Kind switch { ValidationKind.Whole => "wholeNumber", ValidationKind.TextLength => "textLength", _ => rule.Kind.ToString().ToLowerInvariant() };
        Arity(args, member.StartsWith("set") ? 1 : 0); Spend(1, false);
        if (member == "getType") return EnumNumber("DataValidationType", matches.Length != 0 && !exact ? "inconsistent" : kindName);
        if (member is "getValid" or "getInvalidCells")
        {
            if (range.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "验证读取超过 100000 格。");
            Spend((int)range.Count, false); Model.PrepareValueRead(sheet.Id, range, cancellation);
            var failures = new WorkbookValidation().CheckAll(Model.ValueView, Cells(sheet.Id, range).ToArray());
            if (failures.Any(f => f.Title == MessageText.Get("无法确认验证结果"))) throw Unsupported(failures.First(f => f.Title == MessageText.Get("无法确认验证结果")).Message);
            if (member == "getValid") return failures.Count == 0;
            if (failures.Count == 0) return null;
            if (failures.Count > 1024) throw new ModelFailure("BudgetExceeded", "无效区域超过 1024 个，请分块查询。");
            return Handle("RangeAreas", new { sheetId = sheet.Id.ToString(), addresses = failures.Select(f => new CellRectangle(f.Address.Row, f.Address.Column, f.Address.Row, f.Address.Column).ToString()).ToArray() });
        }
        if (member == "clear") { Write(sheet.Id, range, new SetScriptValidationCommand(sheet.Id, range, null)); return null; }
        if (matches.Length != 0 && !exact && member != "setRule") throw Unsupported("混合验证规则请分块读取或使用完整范围替换。");
        WorksheetValidation Copy(ValidationKind? kind = null, string? first = null, string? second = null, ValidationOperator? comparison = null,
            bool? blanks = null, bool? dropdown = null, bool? showInput = null, bool? showError = null, ValidationAlert? alert = null,
            string? inputTitle = null, string? inputMessage = null, string? errorTitle = null, string? errorMessage = null) => new([range], kind ?? rule.Kind,
                kind.HasValue ? first : rule.Formula1, kind.HasValue ? second : rule.Formula2, comparison ?? rule.Comparison, blanks ?? rule.AllowBlank,
                dropdown ?? rule.ShowDropDown, showInput ?? rule.ShowInput, showError ?? rule.ShowError, alert ?? rule.Alert,
                inputTitle ?? rule.InputTitle, inputMessage ?? rule.InputMessage, errorTitle ?? rule.ErrorTitle, errorMessage ?? rule.ErrorMessage);
        if (member == "getIgnoreBlanks") return rule.AllowBlank;
        if (member == "getPrompt") return new { showPrompt = rule.ShowInput, title = rule.InputTitle ?? "", message = rule.InputMessage ?? "" };
        if (member == "getErrorAlert") return new { showAlert = rule.ShowError, title = rule.ErrorTitle ?? "", message = rule.ErrorMessage ?? "", style = (int)rule.Alert };
        if (member == "getRule")
        {
            if (rule.Kind == ValidationKind.None) return new { };
            object value;
            if (rule.Kind == ValidationKind.Custom) value = new { formula = rule.Formula1 };
            else if (rule.Kind == ValidationKind.List)
            {
                var source = QuickExcel.Formulas.FormulaSyntax.Parse(rule.Formula1!) is QuickExcel.Formulas.ValueNode { Value.Kind: CellValueKind.Text } literal ? literal.Value.TextValue : rule.Formula1;
                value = new { source, inCellDropDown = rule.ShowDropDown };
            }
            else
            {
                var op = ComparisonName((ConditionalComparison)rule.Comparison); if (op is "greaterThanOrEqual" or "lessThanOrEqual") op += "To";
                var fields = new Dictionary<string, object> { ["formula1"] = rule.Formula1!, ["operator"] = EnumNumber("DataValidationOperator", op) };
                if (rule.Formula2 is not null) fields["formula2"] = rule.Formula2; value = fields;
            }
            return new Dictionary<string, object> { [kindName] = value };
        }
        WorksheetValidation next;
        if (member == "setIgnoreBlanks") next = Copy(blanks: Bool(args[0]));
        else if (member == "setPrompt")
        { Fields(args[0], "showPrompt", "title", "message"); next = Copy(showInput: Bool(args[0].GetProperty("showPrompt")), inputTitle: Text(args[0].GetProperty("title")), inputMessage: Text(args[0].GetProperty("message"))); }
        else if (member == "setErrorAlert")
        { Fields(args[0], "showAlert", "title", "message", "style"); var alert = Int(args[0].GetProperty("style")); if (alert is < 0 or > 2) throw Invalid("验证警告类型无效。"); next = Copy(showError: Bool(args[0].GetProperty("showAlert")), alert: (ValidationAlert)alert, errorTitle: Text(args[0].GetProperty("title")), errorMessage: Text(args[0].GetProperty("message"))); }
        else if (member == "setRule")
        {
            Fields(args[0], "wholeNumber", "decimal", "date", "textLength", "list", "custom");
            var properties = args[0].EnumerateObject().ToArray();
            if (properties.Length == 0) { Write(sheet.Id, range, new SetScriptValidationCommand(sheet.Id, range, null)); return null; }
            if (properties.Length != 1) throw Invalid("验证规则必须恰有一个类型。");
            var property = properties[0]; var v = property.Value;
            var kind = property.Name switch { "wholeNumber" => ValidationKind.Whole, "decimal" => ValidationKind.Decimal, "date" => ValidationKind.Date, "textLength" => ValidationKind.TextLength, "list" => ValidationKind.List, _ => ValidationKind.Custom };
            string Formula(JsonElement value)
            {
                if (value.ValueKind == JsonValueKind.Number) return "=" + Number(value).ToString("R", CultureInfo.InvariantCulture);
                if (value.ValueKind == JsonValueKind.Object) return "=" + Reference(sheet, Address(value, sheet.Id));
                return Text(value);
            }
            if (kind == ValidationKind.List)
            {
                Fields(v, "source", "inCellDropDown"); var source = v.GetProperty("source"); var text = Formula(source);
                if (source.ValueKind == JsonValueKind.String && !text.StartsWith('=')) text = "=\"" + text.Replace("\"", "\"\"") + "\"";
                next = Copy(kind: kind, first: text, dropdown: Bool(v.GetProperty("inCellDropDown")));
            }
            else if (kind == ValidationKind.Custom) { Fields(v, "formula"); next = Copy(kind: kind, first: Text(v.GetProperty("formula"))); }
            else
            {
                Fields(v, "formula1", "formula2", "operator"); next = Copy(kind: kind, first: Formula(v.GetProperty("formula1")), second: v.TryGetProperty("formula2", out var second) ? Formula(second) : null,
                    comparison: (ValidationOperator)Comparison(EnumName("DataValidationOperator", v.GetProperty("operator"))));
            }
        }
        else throw Unsupported(member);
        Write(sheet.Id, range, new SetScriptValidationCommand(sheet.Id, range, next)); return null;
    }
}
