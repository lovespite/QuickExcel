using System.Security.Cryptography;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Storage;

public sealed record WorkbookSavePart(string Part, string Action, string Reason);

public sealed partial class TypedWorkbookStorage
{
    private static bool HasWorkbookViews(WorkbookSnapshot snapshot) => snapshot.ActiveSheetId != snapshot.Sheets[0].Id || snapshot.Window != new WorkbookWindow() ||
        snapshot.Sheets.Any(s => s.Visibility != SheetVisibility.Visible || s.View != new WorksheetView());

    // Only the canonical empty v3 module is representable by a script-free XLSX.
    // Comments, disabled scripts, custom bindings and signatures are still user content.
    public static bool RequiresScriptContainer(WorkbookSnapshot snapshot)
    {
        var scripts = snapshot.ScriptModules;
        return snapshot.LocalStorage.Entries.Count != 0 || snapshot.PackageFiles.Entries.Count != 0 || scripts.ScriptApiVersion != 3 || scripts.Signature is not null || scripts.Modules.Count != 1 ||
            !scripts.Modules.TryGetValue("main", out var main) || main != new ScriptModule("scripts/main.js", "") ||
            !scripts.EntryModules.SequenceEqual(["main"]) || !scripts.Bindings.SequenceEqual([new ScriptBinding("main", "formula-functions")]);
    }

    public static IReadOnlyList<WorkbookSavePart> InspectSavePlan(WorkbookSnapshot snapshot)
    {
        if (snapshot.FileState is { } original && ContentEquals(snapshot, original.OriginalModel))
            return snapshot.Sheets.Any(s => s.Data.Charts.Any(c => c.IsExtended)) ?
                [new("*", "preserve-package", "工作簿内容未变，保留原包部件；脚本另由 XLSQ 保存。"),
                 new("ChartEx", "repair-excel-compatibility", "写出 XLSX 时补齐缺失的图表样式依赖，并移除会导致 Excel 拒绝文件的系列轴引用；保留筛选和完整源数据。")] :
                [new("*", "preserve-package", "工作簿内容未变，原 XLSX 整包保留；脚本另由 XLSQ 保存。")];
        var result = new List<WorkbookSavePart>
        {
            new("xl/workbook.xml; xl/worksheets/*; relationships", "rebuild", "由当前模型重建工作表、公式、名称、行列与关系；保留已验证的 Excel 版本/标识、行排版、拼音字体和错误提示范围。"),
            new("xl/styles.xml; xl/theme/*", "validated-preserve-or-rebuild", "仅保留已验证的命名样式、XF 与主题；单格修改追加组件。")
        };
        result.AddRange(XlsxMetadata.Plan(snapshot));
        result.AddRange(SaveIssues(snapshot).Select(issue => new WorkbookSavePart("unmodeled", "block-or-explicit-copy", issue)));
        return result;
    }

    private static WorkbookSaveResult SaveXlsx(string path, WorkbookSnapshot snapshot, bool discardUnsupported,
        CancellationToken cancellationToken, FileVersion? expected)
    {
        var bytes = BuildXlsx(snapshot, discardUnsupported, cancellationToken, out var preserve);
        WorkbookFileState? state = null;
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var saved = PackageIO.ReadFile(temp, new(cancellationToken));
            state = SavedFileState(snapshot, VerifyXlsx(path, snapshot, saved, preserve, cancellationToken));
        }, cancellationToken, expected);
        return new(new(Convert.ToHexString(SHA256.HashData(bytes))), state!);
    }

    private static byte[] BuildXlsx(WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken, out bool preserve, bool workbookPartOnly = false)
    {
        if (!workbookPartOnly && snapshot.LocalStorage.Entries.Count > 0) throw new InvalidOperationException(
            MessageText.Get("此工作簿含 localStorage 数据，主文件必须保存为 XLSQ。导出的 XLSX 副本不包含脚本存储。"));
        if (!workbookPartOnly && RequiresScriptContainer(snapshot)) throw new InvalidOperationException(
            MessageText.Get("此工作簿含 JavaScript、模块或签名，主文件必须保存为 XLSQ。交付 Excel 请使用明确的 XLSX 导出副本；导出不会标记主文件已保存。"));
        var issues = SaveIssues(snapshot);
        if (!discardUnsupported && issues.Count > 0) throw new WorkbookCompatibilityException(issues);
        preserve = !discardUnsupported && snapshot.FileState is { } source && ContentEquals(snapshot, source.OriginalModel);
        if (!preserve)
        {
            void ValidateArity(string formula, string location)
            {
                FormulaNode expression;
                try { expression = FormulaSyntax.Parse(formula); }
                catch (Exception ex) when (ex is FormatException or OverflowException) { return; } // Opaque Excel syntax follows the existing preservation rules.
                if (FormulaSyntax.Walk(expression).OfType<FunctionNode>().Any(f => StandardFunctions.Signatures.ContainsKey(f.Name) && !StandardFunctions.AcceptsArguments(f.Name, f.Arguments.Count)))
                    throw new InvalidOperationException(MessageText.Format($"{location} 的函数参数个数不在受支持范围内。请修正公式，或使用明确的值/可移植公式导出。"));
            }
            foreach (var name in snapshot.DefinedNames) ValidateArity(name.Formula, "名称 " + name.Name);
            foreach (var sheet in snapshot.Sheets)
            foreach (var (position, cell) in sheet.Cells)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (cell.Content.Kind != CellContentKind.Formula) continue;
                ValidateArity(cell.Content.FormulaText, sheet.Name + "!" + TypedWorkbookXml.Reference(position));
            }
        }
        return preserve ? XlsxCharts.RepairPreservedExtendedCharts(snapshot.FileState!.OriginalWorkbook.ToArray(), snapshot, cancellationToken) : TypedWorkbookXml.Write(snapshot, false, cancellationToken);
    }

    private static WorkbookFileState VerifyXlsx(string path, WorkbookSnapshot snapshot, byte[] saved, bool preserve, CancellationToken cancellationToken)
    {
        var reopened = OpenBytes(saved, path, false, new(cancellationToken));
        if (!ContentEquals(snapshot, reopened, allowRowHeightRounding: true))
            throw new InvalidDataException(MessageText.Get("XLSX 保存后重开校验未能还原单元格、公式、样式与工作表配置。"));
        if (!preserve) XlsxMetadata.VerifyWritten(snapshot, saved, cancellationToken);
        return reopened.FileState!;
    }
}
