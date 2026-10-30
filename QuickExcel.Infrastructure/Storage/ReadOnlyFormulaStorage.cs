using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    // Read-only files keep the complete source package. Navigation is intentionally not saved.
    // Direct callers cannot bypass the session guard with a modified snapshot or discardUnsupported.
    private static WorkbookSnapshot PrepareReadOnlyFormulaSave(WorkbookSnapshot snapshot, bool discardUnsupported)
    {
        if (snapshot.FileState is not { FormulaReadOnly: true } state) return snapshot;
        if (discardUnsupported) throw new InvalidOperationException("只读公式工作簿不能另存删减特性的公式副本；可另存完整原包或显式导出已有缓存的值快照。");
        var source = state.OriginalModel;
        if (snapshot.Sheets.Length != source.Sheets.Length || snapshot.ScriptModules != source.ScriptModules)
            throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var sheets = snapshot.Sheets.Select((sheet, i) =>
        {
            var old = source.Sheets[i]; var a = sheet.View; var b = old.View;
            if (sheet.Id != old.Id || a.FrozenRows != b.FrozenRows || a.FrozenColumns != b.FrozenColumns || a.FrozenSplit != b.FrozenSplit || a.Zoom != b.Zoom || a.NormalZoom != b.NormalZoom)
                throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
            return sheet.WithView(b);
        }).ToArray();
        var normalized = snapshot.WithSheets(sheets).WithActiveSheet(source.ActiveSheetId);
        if (!ContentEquals(normalized, source)) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        return normalized;
    }
}
