using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? Names(SheetId? scope, string member, JsonElement[] args)
    {
        switch (member)
        {
            case "getNames":
                Arity(args, 0); Spend(1, false);
                return Model.View.DefinedNames
                    .Where(definition => scope is null || definition.Scope == scope)
                    .Select(definition => NamedHandle(definition))
                    .ToArray();
            case "getNamedItem":
            {
                Arity(args, 1); Spend(1, false);
                var name = Text(args[0]);
                var definition = Model.View.DefinedNames.FirstOrDefault(item =>
                    item.Scope == scope && item.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                return definition is null ? null : NamedHandle(definition);
            }
            case "addNamedItem":
                return AddNamedItem(scope, args);
            default:
                throw Unsupported(member);
        }
    }

    private object? AddNamedItem(SheetId? scope, JsonElement[] args)
    {
        Arity(args, 2, 3);
        var name = Text(args[0]);
        ValidateDefinedName(name);
        if (args.Length == 3 && args[2].ValueKind != JsonValueKind.Null && Text(args[2]).Length > 0)
            throw Unsupported("命名区域的注释尚未建模；请省略 comment。");
        if (Model.View.DefinedNames.Any(item => item.Scope == scope && item.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            throw Invalid("名称已存在。");
        var defaultSheet = scope ?? Model.View.ActiveSheetId;
        var formula = NamedReference(args[1], defaultSheet);
        ExecuteObject(defaultSheet, new DefineNameCommand(new DefinedName(name, formula, scope)));
        return NamedHandle(new DefinedName(name, formula, scope));
    }

    private object NamedHandle(DefinedName definition) =>
        Handle("NamedItem", new { name = definition.Name, scopeSheetId = definition.Scope?.ToString() });

    private string NamedReference(JsonElement reference, SheetId defaultSheet)
    {
        if (reference.ValueKind == JsonValueKind.Object)
        {
            if (!reference.TryGetProperty("sheetId", out var id) || !Guid.TryParse(id.GetString(), out var sheetId))
                throw Invalid("需要 Range 引用。");
            var sheet = Model.View.Sheets.FirstOrDefault(item => item.Id.Value == sheetId)
                ?? throw new ModelFailure("InvalidReference", "区域句柄已失效。");
            return "=" + Reference(sheet, Address(reference, sheet.Id));
        }
        if (reference.ValueKind != JsonValueKind.String) throw Invalid("需要 Range 或区域引用字符串。");
        var text = Text(reference).Trim();
        if (text.Length == 0) throw Invalid("引用不能为空。");
        var expression = FormulaSyntax.Parse(text.StartsWith('=') ? text : "=" + text);
        if (expression is not ReferenceNode referenceNode) throw Invalid("addNamedItem 的引用必须是单个矩形区域。");
        var sheetName = referenceNode.Sheet ?? Model.View.GetSheet(defaultSheet).Name;
        var target = Model.View.Sheets.FirstOrDefault(item => item.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ModelFailure("ItemNotFound", "找不到引用的工作表。");
        var top = referenceNode.WholeColumns ? 0 : referenceNode.Row;
        var bottom = referenceNode.WholeColumns ? WorkbookLimits.MaxRows - 1
            : referenceNode.IsRange ? referenceNode.EndRow : referenceNode.Row;
        var left = referenceNode.Column;
        var right = referenceNode.IsRange ? referenceNode.EndColumn : referenceNode.Column;
        WorkbookLimits.ValidatePosition(top, left); WorkbookLimits.ValidatePosition(bottom, right);
        if (top > bottom || left > right) throw Invalid("引用的区域顺序无效。");
        return "=" + FormulaSyntax.Print(referenceNode with { Sheet = target.Name });
    }

    private object? NamedItemApi(JsonElement data, string member, JsonElement[] args)
    {
        var definition = FindNamedItem(data);
        switch (member)
        {
            case "getFormula":
                Arity(args, 0); Spend(1, false);
                return definition.Formula;
            case "setFormula":
            {
                Arity(args, 1);
                var formula = Text(args[0]);
                if (formula.Length is < 2 or > 8192 || formula[0] != '=')
                    throw Invalid("名称公式必须以 = 开头且不超过 8192 个字符。");
                _ = FormulaSyntax.Parse(formula);
                ExecuteObject(definition.Scope ?? Model.View.ActiveSheetId,
                    new EditDefinedNameCommand(definition.Name, definition.Scope, definition.Name, formula));
                return null;
            }
            case "getValue":
            {
                Arity(args, 0);
                if (TryResolveReference(definition, out var sheetId, out var range))
                {
                    var position = new CellPosition(range.Top, range.Left);
                    Model.PrepareValueRead(sheetId, new(range.Top, range.Left, range.Top, range.Left), cancellation);
                    Spend(1, false);
                    var cell = Model.ReadCell(sheetId, position, false, calculated: true);
                    EnsureReadable(cell);
                    return Scalar(cell?.Value ?? CellValue.Blank);
                }
                if (FormulaSyntax.Parse(definition.Formula) is ValueNode value) return Scalar(value.Value);
                throw Unsupported("该名称不是区域引用或简单常量，暂不能读取值。");
            }
            case "getRange":
                Arity(args, 0);
                if (!TryResolveReference(definition, out var resolvedSheet, out var resolvedRange))
                    throw Unsupported("该名称不是区域引用。");
                Spend(1, false);
                return RangeHandle(resolvedSheet, resolvedRange);
            case "getType":
                Arity(args, 0); Spend(1, false);
                if (TryResolveReference(definition, out _, out _)) return EnumNumber("NamedItemType", "range");
                if (FormulaSyntax.Parse(definition.Formula) is ValueNode constant)
                    return EnumNumber("NamedItemType", constant.Value.Kind switch
                    {
                        CellValueKind.Text => "string", CellValueKind.Boolean => "boolean", CellValueKind.Error => "error",
                        CellValueKind.Number => Math.Truncate(constant.Value.NumberValue) == constant.Value.NumberValue ? "integer" : "double",
                        _ => "string"
                    });
                throw Unsupported("该名称的类型尚未支持。");
            case "delete":
                Arity(args, 0);
                ExecuteObject(definition.Scope ?? Model.View.ActiveSheetId,
                    new DeleteDefinedNameCommand(definition.Name, definition.Scope));
                return null;
            default:
                throw Unsupported(member);
        }
    }

    private DefinedName FindNamedItem(JsonElement data)
    {
        if (!data.TryGetProperty("name", out var nameValue) || nameValue.ValueKind != JsonValueKind.String)
            throw new ModelFailure("InvalidReference", "命名区域句柄已失效。");
        SheetId? scope = null;
        if (data.TryGetProperty("scopeSheetId", out var scopeValue) && scopeValue.ValueKind == JsonValueKind.String &&
            Guid.TryParse(scopeValue.GetString(), out var scopeId))
            scope = new SheetId(scopeId);
        var name = nameValue.GetString()!;
        return Model.View.DefinedNames.FirstOrDefault(item =>
            item.Scope == scope && item.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ModelFailure("InvalidReference", "命名区域句柄已失效。");
    }

    private bool TryResolveReference(DefinedName definition, out SheetId sheetId, out CellRectangle range)
    {
        sheetId = default; range = default;
        try
        {
            if (FormulaSyntax.Parse(definition.Formula) is not ReferenceNode reference) return false;
            var sheet = reference.Sheet is null
                ? definition.Scope is { } scope ? Model.View.GetSheet(scope) : null
                : Model.View.Sheets.FirstOrDefault(item => item.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
            if (sheet is null) return false;
            var top = reference.WholeColumns ? 0 : reference.Row;
            var bottom = reference.WholeColumns ? WorkbookLimits.MaxRows - 1 : reference.IsRange ? reference.EndRow : reference.Row;
            var left = reference.Column;
            var right = reference.IsRange ? reference.EndColumn : reference.Column;
            WorkbookLimits.ValidatePosition(top, left); WorkbookLimits.ValidatePosition(bottom, right);
            if (top > bottom || left > right) return false;
            sheetId = sheet.Id; range = new(top, left, bottom, right);
            return true;
        }
        catch (Exception error) when (error is ArgumentException or FormatException or OverflowException)
        {
            return false;
        }
    }

    private static void ValidateDefinedName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 ||
            !(char.IsLetter(name[0]) || name[0] == '_') ||
            name.Any(c => !(char.IsLetterOrDigit(c) || c is '_' or '.')) ||
            System.Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z]{1,3}[1-9][0-9]*$") ||
            System.Text.RegularExpressions.Regex.IsMatch(name, "^[Rr][0-9]*[Cc][0-9]*$") ||
            name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("R", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("C", StringComparison.OrdinalIgnoreCase))
            throw Invalid("名称无效；请使用字母或下划线开头且不含空格的名称。");
    }
}
