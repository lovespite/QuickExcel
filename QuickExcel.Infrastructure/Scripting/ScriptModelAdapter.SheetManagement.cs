using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? WorkbookSheetApi(string member, JsonElement[] args)
    {
        switch (member)
        {
            case "addWorksheet":
            {
                Arity(args, 0, 1);
                if (Model.View.Sheets.Length >= WorkbookLimits.MaxSheets)
                    throw new ModelFailure("InvalidArgument", "工作表数量不能超过 64。");
                var name = args.Length == 1 && args[0].ValueKind != JsonValueKind.Null ? Text(args[0]) : null;
                if (name is not null)
                {
                    WorkbookLimits.ValidateSheetName(name);
                    if (Model.View.Sheets.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        throw Invalid("工作表名称已存在。");
                }
                else
                {
                    var number = Model.View.Sheets.Length + 1;
                    while (Model.View.Sheets.Any(s => s.Name.Equals("Sheet" + number, StringComparison.OrdinalIgnoreCase))) number++;
                    name = "Sheet" + number;
                }
                var command = new AddSheetCommand(name);
                ExecuteObject(Model.View.ActiveSheetId, command);
                return Handle("Worksheet", new Dictionary<string, object?> { ["sheetId"] = command.Id.ToString() });
            }
            case "getFirstWorksheet":
            case "getLastWorksheet":
            {
                Arity(args, 0, 1);
                var visibleOnly = args.Length == 1 && Bool(args[0]);
                Spend(1, false);
                var sheets = Model.View.Sheets.Where(s => !visibleOnly || s.Visibility == SheetVisibility.Visible).ToArray();
                if (sheets.Length == 0) return null;
                var sheet = member == "getFirstWorksheet" ? sheets[0] : sheets[^1];
                return Handle("Worksheet", new Dictionary<string, object?> { ["sheetId"] = sheet.Id.ToString() });
            }
            default:
                throw Unsupported(member);
        }
    }

    private object? WorksheetSheetApi(WorksheetSnapshot sheet, string member, JsonElement[] args)
    {
        switch (member)
        {
            case "setName":
                Arity(args, 1);
                ExecuteObject(sheet.Id, new RenameSheetCommand(sheet.Id, Text(args[0])));
                return null;
            case "getPosition":
                Arity(args, 0); Spend(1, false);
                return SheetPosition(sheet.Id);
            case "setPosition":
                Arity(args, 1);
                ExecuteObject(sheet.Id, new ReorderSheetCommand(sheet.Id, Int(args[0])));
                return null;
            case "getVisibility":
                Arity(args, 0); Spend(1, false);
                return EnumNumber("SheetVisibility", VisibilityName(sheet.Visibility));
            case "setVisibility":
            {
                Arity(args, 1);
                var visibility = Visibility(EnumName("SheetVisibility", args[0]));
                if (visibility != SheetVisibility.Visible && sheet.Visibility == SheetVisibility.Visible &&
                    Model.View.Sheets.Count(s => s.Visibility == SheetVisibility.Visible) == 1)
                    throw Invalid("至少需要保留一张可见工作表。");
                ExecuteObject(sheet.Id, new SetSheetVisibilityCommand(sheet.Id, visibility));
                return null;
            }
            case "getTabColor":
                Arity(args, 0); Spend(1, false);
                return Rgb(sheet.TabColor);
            case "setTabColor":
            {
                Arity(args, 1);
                var color = args[0].ValueKind == JsonValueKind.Null ? null : Text(args[0]);
                ExecuteObject(sheet.Id, new SetSheetTabColorCommand(sheet.Id, string.IsNullOrEmpty(color) ? null : color));
                return null;
            }
            case "getShowGridlines":
                Arity(args, 0); Spend(1, false);
                return sheet.ShowGridLines;
            case "setShowGridlines":
                Arity(args, 1);
                ExecuteObject(sheet.Id, new SetSheetGridLinesCommand(sheet.Id, Bool(args[0])));
                return null;
            case "getNext":
            case "getPrevious":
            {
                Arity(args, 0, 1);
                var visibleOnly = args.Length == 1 && Bool(args[0]);
                Spend(1, false);
                return AdjacentSheet(sheet.Id, member == "getNext" ? 1 : -1, visibleOnly);
            }
            case "delete":
                Arity(args, 0);
                ExecuteObject(sheet.Id, new DeleteSheetCommand(sheet.Id));
                return null;
            case "activate":
                Arity(args, 0);
                ExecuteObject(sheet.Id, new SetActiveSheetCommand(sheet.Id));
                return null;
            default:
                throw Unsupported(member);
        }
    }

    private void ExecuteObject(SheetId sheet, DocumentCommand command, int cost = 1) =>
        Model.ExecuteObjectCommand(sheet, cost, command);

    private int SheetPosition(SheetId id)
    {
        for (var index = 0; index < Model.View.Sheets.Length; index++)
            if (Model.View.Sheets[index].Id == id) return index;
        throw new ModelFailure("InvalidReference", "工作表句柄已失效。");
    }

    private object? AdjacentSheet(SheetId id, int direction, bool visibleOnly)
    {
        var sheets = Model.View.Sheets;
        for (var index = SheetPosition(id) + direction; index >= 0 && index < sheets.Length; index += direction)
            if (!visibleOnly || sheets[index].Visibility == SheetVisibility.Visible)
                return Handle("Worksheet", new Dictionary<string, object?> { ["sheetId"] = sheets[index].Id.ToString() });
        return null;
    }

    private static string VisibilityName(SheetVisibility visibility) => visibility switch
    {
        SheetVisibility.Visible => "visible",
        SheetVisibility.Hidden => "hidden",
        SheetVisibility.VeryHidden => "veryHidden",
        _ => throw Invalid("工作表可见性无效。")
    };

    private static SheetVisibility Visibility(string name) => name switch
    {
        "visible" => SheetVisibility.Visible,
        "hidden" => SheetVisibility.Hidden,
        "veryHidden" => SheetVisibility.VeryHidden,
        _ => throw Invalid("工作表可见性无效。")
    };

    private object? FreezePanesApi(WorksheetSnapshot sheet, string member, JsonElement[] args)
    {
        var view = sheet.View;
        switch (member)
        {
            case "freezeAt":
            {
                Arity(args, 1);
                var range = Address(args[0], sheet.Id);
                ApplyFreeze(sheet.Id, Freeze(view, range.Top, range.Left));
                return null;
            }
            case "freezeRows":
                Arity(args, 0, 1);
                ApplyFreeze(sheet.Id, Freeze(view, args.Length == 0 ? 0 : Int(args[0]), view.FrozenColumns));
                return null;
            case "freezeColumns":
                Arity(args, 0, 1);
                ApplyFreeze(sheet.Id, Freeze(view, view.FrozenRows, args.Length == 0 ? 0 : Int(args[0])));
                return null;
            case "unfreeze":
                Arity(args, 0);
                ApplyFreeze(sheet.Id, Freeze(view, 0, 0));
                return null;
            case "getLocation":
                Arity(args, 0); Spend(1, false);
                return view.HasFreeze
                    ? RangeHandle(sheet.Id, new(0, 0, Math.Max(1, view.FrozenRows) - 1, Math.Max(1, view.FrozenColumns) - 1))
                    : null;
            default:
                throw Unsupported(member);
        }
    }

    private void ApplyFreeze(SheetId sheetId, WorksheetView view) =>
        ExecuteObject(sheetId, new SetSheetViewCommand(sheetId, view));

    private static WorksheetView Freeze(WorksheetView view, int rows, int columns)
    {
        if (rows < 0 || rows >= WorkbookLimits.MaxRows || columns < 0 || columns >= WorkbookLimits.MaxColumns)
            throw Invalid("冻结行列数超出范围。");
        if (rows == 0 && columns == 0)
            return view with
            {
                FrozenRows = 0, FrozenColumns = 0, FrozenSplit = false, PaneTopLeftCell = default,
                ActivePane = SheetPane.TopLeft, TopRight = null, BottomLeft = null, BottomRight = null
            };
        return view with { FrozenRows = rows, FrozenColumns = columns, PaneTopLeftCell = new(rows, columns), FrozenSplit = false };
    }
}
