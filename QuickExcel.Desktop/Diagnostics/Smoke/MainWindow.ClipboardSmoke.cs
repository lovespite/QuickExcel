using System.IO;
using System.Windows;
using System.Windows.Input;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunClipboardSmokeAsync(string folder, Action<bool, string> check)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Clipboard", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1))),
            new(new(0, 1), new(CellContent.FromFormula("=A2+$A$2"))), new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(4)))),
            new(new(2, 0), new(CellContent.FromLiteral(CellValue.FromNumber(7))))]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true, NumberFormatId = 49 }]);
        InstallDocument(new WorkbookDocument(snapshot, _services), null, "Clipboard smoke");
        GridSurface.SelectRange(new(0, 0, 0, 1), new(0, 0));
        await HandleGridKeyAsync(Key.C, ModifierKeys.Control, false);
        check(Clipboard.GetData(InternalClipboardFormat) as string == _clipboardMarker && _workbookClipboard is not null && !_document.IsDirty,
            "Ctrl+C captures typed content and a private OS marker without editing the workbook");
        GridSurface.Select(new(4, 4)); await HandleGridKeyAsync(Key.V, ModifierKeys.Control, false);
        var pasted = _document.Snapshot;
        check(pasted.Sheets[0].GetCell(4, 4)!.Value == CellValue.FromText("00123") && pasted.Sheets[0].GetCell(4, 4)!.Content.StyleId == 1 &&
            pasted.Sheets[0].GetCell(4, 5)!.Content.FormulaText == "=(E6+$A$2)", "Ctrl+V keeps text IDs and styles while translating relative and absolute references");
        GridSurface.Select(new(6, 4)); await PasteClipboardAsync("all");
        check(_document.Snapshot.Sheets[0].GetCell(6, 4)!.Value.TextValue == "00123", "Repeated internal paste uses the same captured inputs");
        await UndoAsync(false);
        var before = _document.Snapshot;
        try { await PasteClipboardAsync("all"); throw new Exception("Expired clipboard accepted."); } catch (InvalidOperationException) { }
        check(ReferenceEquals(before, _document.Snapshot), "Undo expires the clipboard revision lease without any partial paste");
        GridSurface.Select(new(2, 0)); before = _document.Snapshot; await HandleGridKeyAsync(Key.X, ModifierKeys.Control, false);
        check(ReferenceEquals(before, _document.Snapshot), "Ctrl+X does not clear the source until a successful paste");
        GridSurface.Select(new(2, 1)); await HandleGridKeyAsync(Key.V, ModifierKeys.Control, false);
        check(_document.Snapshot.Sheets[0].GetCell(2, 0) is null && _document.Snapshot.Sheets[0].GetCell(2, 1)!.Value.NumberValue == 7 && _workbookClipboard is null,
            "Cut paste moves once and consumes the pending cut");
        await UndoAsync(false);
        check(_document.Snapshot.Sheets[0].GetCell(2, 0)!.Value.NumberValue == 7 && _document.Snapshot.Sheets[0].GetCell(2, 1) is null,
            "One undo restores both sides of a cut");
        GridSurface.Select(new(2, 0)); CopySelection(cut: true); await HandleGridKeyAsync(Key.Escape, ModifierKeys.None, false);
        check(_workbookClipboard is null && _document.Snapshot.Sheets[0].GetCell(2, 0)!.Value.NumberValue == 7, "Escape cancels a pending cut without clearing the source");
        Clipboard.SetText("=1+2"); GridSurface.Select(new(9, 0)); await PasteClipboardAsync("all");
        check(_document.Snapshot.Sheets[0].GetCell(9, 0)!.Value.TextValue == "=1+2", "External TSV still keeps formula text literal by default");
        GridSurface.Select(new(9, 1)); await PasteClipboardAsync("formulas");
        check(_document.Snapshot.Sheets[0].GetCell(9, 1)!.Content.Kind == CellContentKind.Formula && !_document.ScriptsEnabled,
            "Explicit external formula paste does not grant script authorization");
        GridSurface.Select(new(0, 1)); CopySelection(); GridSurface.Select(new(0, 7));
        RibbonPaste_Click(new System.Windows.Controls.MenuItem { Tag = "values" }, new RoutedEventArgs()); await _gridInput;
        check(_document.Snapshot.Sheets[0].GetCell(0, 7)!.Content.Kind == CellContentKind.Literal && _document.Snapshot.Sheets[0].GetCell(0, 7)!.Value.NumberValue == 8,
            "The values menu pastes a formula result as a typed number");
        GridSurface.Select(new(0, 0)); CopySelection(); GridSurface.Select(new(2, 0));
        RibbonPaste_Click(new System.Windows.Controls.MenuItem { Tag = "formats" }, new RoutedEventArgs()); await _gridInput;
        check(_document.Snapshot.Sheets[0].GetCell(2, 0)!.Value.NumberValue == 7 && _document.Snapshot.Sheets[0].GetCell(2, 0)!.Content.StyleId == 1,
            "The formats menu preserves the destination value");
        GridSurface.SelectRange(new(0, 0, 0, 1), new(0, 0)); CopySelection(); GridSurface.Select(new(0, 9));
        RibbonPaste_Click(new System.Windows.Controls.MenuItem { Tag = "transpose" }, new RoutedEventArgs()); await _gridInput;
        check(_document.Snapshot.Sheets[0].GetCell(0, 9)!.Value.TextValue == "00123" && _document.Snapshot.Sheets[0].GetCell(1, 9)!.Content.FormulaText == "=(K1+$A$2)",
            "The transpose menu maps cells vertically and rotates fully relative formula references");
        GridSurface.Select(new(0, 0)); CopySelection();
        _editor.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 12, 1, hidden: true));
        CopySelection(); GridSurface.Select(new(12, 0)); before = _document.Snapshot;
        try { await PasteClipboardAsync("all"); throw new Exception("Hidden target accepted."); } catch (InvalidOperationException) { }
        check(ReferenceEquals(before, _document.Snapshot), "Internal paste rejects hidden target rows without silently remapping");
        GridSurface.Select(new(0, 0)); CopySelection();
        GridSurface.Select(new(14, 0)); before = _document.Snapshot; _editState.StartComposition(); await PasteClipboardAsync("all"); _editState.EndComposition();
        check(ReferenceEquals(before, _document.Snapshot), "IME composition gates internal paste");
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(folder, "internal-clipboard.xlsx"); storage.Save(path, _document.Snapshot);
        var opened = storage.Open(path);
        check(opened.Sheets[0].GetCell(4, 4)!.Value.TextValue == "00123" && opened.Sheets[0].GetCell(4, 5)!.Content.FormulaText == "=(E6+$A$2)",
            "Desktop clipboard results survive normal XLSX save and reopen");
        InstallDocument(new WorkbookDocument(snapshot, _services), null, "New clipboard session"); before = _document.Snapshot;
        try { await PasteClipboardAsync("all"); throw new Exception("Old-session marker accepted."); } catch (InvalidOperationException) { }
        check(ReferenceEquals(before, _document.Snapshot), "A new session rejects the old marker even when the document identity and revision match");
    }
}
