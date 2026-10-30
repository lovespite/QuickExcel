using System.Security.Cryptography;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed record WorkbookSavePart(string Part, string Action, string Reason);

public sealed partial class TypedWorkbookStorage
{
    // Only the canonical empty v1 module is representable by a script-free XLSX.
    // Comments, disabled scripts, custom bindings and signatures are still user content.
    public static bool RequiresScriptContainer(WorkbookSnapshot snapshot)
    {
        var scripts = snapshot.ScriptModules;
        return scripts.ScriptApiVersion != 1 || scripts.Signature is not null || scripts.Modules.Count != 1 ||
            !scripts.Modules.TryGetValue("main", out var main) || main != new ScriptModule("scripts/main.js", "") ||
            !scripts.EntryModules.SequenceEqual(["main"]) || !scripts.Bindings.SequenceEqual([new ScriptBinding("main", "formula-functions")]);
    }

    public static IReadOnlyList<WorkbookSavePart> InspectSavePlan(WorkbookSnapshot snapshot)
    {
        if (snapshot.FileState is { } original && ContentEquals(snapshot, original.OriginalModel))
            return [new("*", "preserve-package", "工作簿内容未变，原 XLSX 整包保留；脚本另由 XLSQ 保存。")];
        var result = new List<WorkbookSavePart>
        {
            new("xl/workbook.xml; xl/worksheets/*; relationships", "rebuild", "由当前模型重建工作表、公式、名称、行列与关系。"),
            new("xl/styles.xml; xl/theme/*", "validated-preserve-or-rebuild", "仅保留已验证的命名样式、XF 与主题；单格修改追加组件。")
        };
        result.AddRange(XlsxMetadata.Plan(snapshot));
        result.AddRange(SaveIssues(snapshot).Select(issue => new WorkbookSavePart("unmodeled", "block-or-explicit-copy", issue)));
        return result;
    }

    private static FileVersion SaveXlsx(string path, WorkbookSnapshot snapshot, bool discardUnsupported,
        CancellationToken cancellationToken, FileVersion? expected)
    {
        if (RequiresScriptContainer(snapshot)) throw new InvalidOperationException(
            "此工作簿含 JavaScript、模块或签名，主文件必须保存为 XLSQ。交付 Excel 请使用明确的 XLSX 导出副本；导出不会标记主文件已保存。");
        var issues = SaveIssues(snapshot);
        if (!discardUnsupported && issues.Count > 0) throw new WorkbookCompatibilityException(issues);
        var preserve = !discardUnsupported && snapshot.FileState is { } source && ContentEquals(snapshot, source.OriginalModel);
        var bytes = preserve ? snapshot.FileState!.OriginalWorkbook.ToArray() : TypedWorkbookXml.Write(snapshot, false, cancellationToken);
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var saved = PackageIO.ReadFile(temp, new(cancellationToken));
            var reopened = OpenBytes(saved, path, false, new(cancellationToken));
            if (!ContentEquals(snapshot, reopened, allowRowHeightRounding: true))
                throw new InvalidDataException("XLSX 保存后重开校验未能还原单元格、公式、样式与工作表配置。");
            if (!preserve) XlsxMetadata.VerifyWritten(snapshot, saved, cancellationToken);
        }, cancellationToken, expected);
        return new(Convert.ToHexString(SHA256.HashData(bytes)));
    }
}
