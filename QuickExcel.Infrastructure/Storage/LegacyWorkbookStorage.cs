using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

/// <summary>The 0.2 storage boundary. XLSQ formatVersion 1 remains the on-disk format.</summary>
public sealed class LegacyWorkbookStorage : IWorkbookStorage
{
    public const string LegacyTypeWarning = "0.2 试用版沿用 XLSQ v1：保存及导出会将字面量写为文本，空白与空文本可能合并；重新打开时不猜测恢复数字、布尔或错误类型。公式结果仅作为缓存载入。";

    public WorkbookSnapshot Open(string path)
    {
        var (workbook, script) = DocumentPackage.Open(path);
        return LegacyWorkbookAdapter.FromWorkbook(workbook, script);
    }

    public void Save(string path, WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        DocumentPackage.Save(path, LegacyWorkbookAdapter.ToWorkbook(snapshot), snapshot.ScriptSource);
    }

    public void Export(string path, WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        DocumentPackage.Export(path, LegacyWorkbookAdapter.ToWorkbook(snapshot, valuesOnly: true));
    }
}
