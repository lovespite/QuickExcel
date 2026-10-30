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

public sealed class SetSheetVisibilityCommand(SheetId sheetId, SheetVisibility visibility) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old = snapshot.GetSheet(sheetId);
        if (old.Visibility == visibility) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithVisibility(visibility) : s)), [], true, CalculationUnchanged: true);
    }
}
