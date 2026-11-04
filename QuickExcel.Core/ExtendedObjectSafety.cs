using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Threaded comments stay attached to their cells; content edits never authorize moving or flattening discussions.</summary>
public static class ExtendedObjectSafety
{
    public const string ThreadedComments = "worksheet/threadedComments";
    private const string Message = "线程评论及回复尚不支持结构移动或复制；含旧版显示副本的讨论只支持原位保留，请在 Excel 中执行该操作。";
    private static bool Present(WorkbookSnapshot snapshot) => snapshot.Sheets.Any(s => !s.Data.ThreadComments.IsEmpty) || snapshot.FileState?.PreservedAttachments.Any(a => a.Feature == ThreadedComments) == true;
    public static void ValidateRetention(WorkbookSnapshot snapshot)
    {
        if (snapshot.FileState?.PreservedAttachments.Any(a => a.Feature == ThreadedComments) != true) return;
        var original = snapshot.FileState!.OriginalModel;
        if (!snapshot.Sheets.Select(s => (s.Id, s.Name)).SequenceEqual(original.Sheets.Select(s => (s.Id, s.Name))))
            throw new NotSupportedException(Message);
        foreach (var name in snapshot.FileState.PreservedAttachments.Where(a => a.Feature == ThreadedComments).Select(a => a.SheetName).Distinct())
        {
            var before = original.Sheets.Single(s => s.Name == name).Annotations;
            var after = snapshot.Sheets.Single(s => s.Name == name).Annotations;
            if (before.Notes.Count != after.Notes.Count || before.Notes.Any(p => after.Notes.GetValueOrDefault(p.Key) != p.Value) || !before.NoteDrawing.SequenceEqual(after.NoteDrawing))
                throw new NotSupportedException(Message);
        }
    }
    public static void ValidateEdit(WorkbookSnapshot before, WorkbookSnapshot after, DocumentCommand command, bool structureChanged)
    {
        if (!Present(before) || ReferenceEquals(before, after)) return;
        if (structureChanged && command is not (SetSheetViewCommand or SetSheetVisibilityCommand or SetAxisMetadataCommand or SetPrintingCommand
            or SetAutoFilterCommand or SetTableFilterCommand or ClearSortStateCommand or SetOutlineCollapsedCommand or SetChartCommand))
            throw new NotSupportedException(Message);
        ValidateRetention(after);
    }
    public static void ValidateCopy(WorkbookSnapshot snapshot, SheetId sheet, CellRectangle range)
    {
        var name = snapshot.GetSheet(sheet).Name;
        if (snapshot.GetSheet(sheet).Data.ThreadComments.Any(c => range.Contains(c.Position.Row, c.Position.Column)) || snapshot.FileState?.PreservedAttachments.Any(a => a.Feature == ThreadedComments && a.SheetName == name && range.Contains(a.Position.Row, a.Position.Column)) == true ||
            snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Code == "xlsx.unmodeled" && d.Feature.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException(Message);
    }
}
