using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class SetSheetViewCommand(SheetId sheetId, WorksheetView view) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old = snapshot.GetSheet(sheetId);
        if (old.View == view) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithView(view) : s)), [], true, CalculationUnchanged: true);
    }
}

public sealed class SetActiveSheetCommand(SheetId sheetId) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Visibility != SheetVisibility.Visible) throw new InvalidOperationException(MessageText.Get("不能激活隐藏工作表。"));
        if (snapshot.ActiveSheetId == sheetId) return new(snapshot, [], false);
        return new(snapshot.WithActiveSheet(sheetId), [], false, CalculationUnchanged: true);
    }
}

public sealed class SetSheetVisibilityCommand(SheetId sheetId, SheetVisibility visibility) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old = snapshot.GetSheet(sheetId);
        if (old.Visibility == visibility) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithVisibility(visibility) : s)), [], true, CalculationUnchanged: true);
    }
}

public sealed class SetSheetGridLinesCommand(SheetId sheetId, bool showGridLines) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old = snapshot.GetSheet(sheetId);
        if (old.ShowGridLines == showGridLines) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithShowGridLines(showGridLines) : s)), [], false, CalculationUnchanged: true);
    }
}
