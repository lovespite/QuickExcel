using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? ProtectionApi(WorksheetSnapshot sheet, string member, JsonElement[] args)
    {
        var p = sheet.Data.Protection;
        if (member.StartsWith("get", StringComparison.Ordinal))
        {
            Arity(args, 0); Spend(1, false);
            return member switch
            {
                "getProtected" => p?.Sheet == true, "getIsPasswordProtected" => p?.HasPassword == true,
                "getCanPauseProtection" or "getIsPaused" => false,
                "getOptions" or "getSavedOptions" => ProtectionOptions(p ?? new()), _ => throw Unsupported(member)
            };
        }
        if (member == "unprotect")
        {
            Arity(args, 0, 1);
            if (args.Length != 0 || p?.HasPassword == true) throw Unsupported("密码保护必须在 Excel 中解除。");
            Write(sheet.Id, ScalarRange, new SetProtectionCommand(sheet.Id, null)); return null;
        }
        if (member == "protect")
        {
            Arity(args, 0, 2);
            if (args.Length == 2 && args[1].ValueKind != JsonValueKind.Null && Text(args[1]).Length != 0) throw Unsupported("密码保护尚未实现。");
            if (p?.Sheet == true) throw Invalid("工作表已经保护。");
            var next = args.Length == 0 || args[0].ValueKind == JsonValueKind.Null ? new WorksheetProtection() : ReadProtectionOptions(args[0]);
            Write(sheet.Id, ScalarRange, new SetProtectionCommand(sheet.Id, next)); return null;
        }
        throw Unsupported(member);
    }
    private static object ProtectionOptions(WorksheetProtection p) => new Dictionary<string, object?> {
        ["allowAutoFilter"] = !p.AutoFilter,
        ["allowDeleteColumns"] = !p.DeleteColumns,
        ["allowDeleteRows"] = !p.DeleteRows,
        ["allowEditObjects"] = !p.Objects,
        ["allowEditScenarios"] = !p.Scenarios,
        ["allowFormatCells"] = !p.FormatCells,
        ["allowFormatColumns"] = !p.FormatColumns,
        ["allowFormatRows"] = !p.FormatRows,
        ["allowInsertColumns"] = !p.InsertColumns,
        ["allowInsertHyperlinks"] = !p.InsertHyperlinks,
        ["allowInsertRows"] = !p.InsertRows,
        ["allowPivotTables"] = !p.PivotTables,
        ["allowSort"] = !p.Sort,
        ["selectionMode"] = p.SelectUnlockedCells ? 2 : p.SelectLockedCells ? 1 : 0
    };
    private static WorksheetProtection ReadProtectionOptions(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object) throw Invalid("保护选项必须是对象。");
        var p = new WorksheetProtection();
        foreach (var item in value.EnumerateObject())
        {
            if (item.Name == "selectionMode")
            {
                var mode = Int(item.Value); if (mode is < 0 or > 2) throw Invalid("选择模式枚举无效。");
                p = p with { SelectLockedCells = mode != 0, SelectUnlockedCells = mode == 2 }; continue;
            }
            var deny = !Bool(item.Value);
            p = item.Name switch
            {
                "allowAutoFilter" => p with { AutoFilter = deny }, "allowDeleteColumns" => p with { DeleteColumns = deny },
                "allowDeleteRows" => p with { DeleteRows = deny }, "allowEditObjects" => p with { Objects = deny },
                "allowEditScenarios" => p with { Scenarios = deny }, "allowFormatCells" => p with { FormatCells = deny },
                "allowFormatColumns" => p with { FormatColumns = deny }, "allowFormatRows" => p with { FormatRows = deny },
                "allowInsertColumns" => p with { InsertColumns = deny }, "allowInsertHyperlinks" => p with { InsertHyperlinks = deny },
                "allowInsertRows" => p with { InsertRows = deny }, "allowPivotTables" => p with { PivotTables = deny },
                "allowSort" => p with { Sort = deny }, _ => throw Invalid("未知保护选项：" + item.Name)
            };
        }
        return p;
    }
}
