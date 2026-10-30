using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private void SetSmokeFreeze(int rows, int columns)
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var view = sheet.View with { FrozenRows = rows, FrozenColumns = columns, FrozenSplit = false,
            PaneTopLeftCell = new(rows, columns), ActivePane = SheetPane.TopLeft,
            TopLeft = null, TopRight = null, BottomLeft = null, BottomRight = null };
        _document.Session.Execute(new SetSheetViewCommand(sheet.Id, view.WithActiveCell(sheet.View.ActiveCell)));
    }

    private async Task RunWorkbookViewSmokeAsync(string folder, Action<bool, string> check)
    {
        var first = new WorksheetSnapshot(SheetId.New(), "首表", view: new WorksheetView { Zoom = 90 }.WithActiveCell(new(3, 2)));
        var active = new WorksheetSnapshot(SheetId.New(), "活动表", view: new WorksheetView
        {
            Zoom = 150, FrozenRows = 2, FrozenColumns = 1, PaneTopLeftCell = new(2, 1)
        }.WithActiveCell(new(5, 3)));
        var hidden = new WorksheetSnapshot(SheetId.New(), "隐藏", visibility: SheetVisibility.Hidden);
        var secret = new WorksheetSnapshot(SheetId.New(), "非常隐藏", visibility: SheetVisibility.VeryHidden);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [first, hidden, active, secret], activeSheetId: active.Id);
        var path = Path.Combine(folder, "ui-view.xlsx"); new TypedWorkbookStorage().Save(path, snapshot);
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        try
        {
            await OpenPathAsync(path); await Idle();
            check(_sheetIndex == 2 && SheetPicker.Items.Count == 2 && SheetPicker.SelectedIndex == 1, "XLSX active sheet restores through visible-tab mapping");
            check(_selectedRow == 5 && _selectedColumn == 3 && CellAddress.Text == "D6", "Imported active cell restores without resetting to A1");
            check(_grid.Control.FrozenRowCount == 2 && _grid.Control.FrozenColumnCount == 1, "Imported frozen panes project to the actual grid");
            check(((ScaleTransform)SheetPageContent.LayoutTransform).ScaleX == 1.5, "Imported sheet zoom scales the actual grid");
            check(!_document.IsDirty && !_document.CanUndo, "Restoring a view creates no edit or undo entry");
            RibbonTabs.SelectedIndex = RibbonTabs.Items.Count - 1; await Idle();
            CaptureGridEvidence(Path.Combine(folder, "workbook-view.png"));
            NavigateToCell(2, 6, 4); await Idle();
            check(_sheetIndex == 2 && CellAddress.Text == "E7", "Address navigation maps worksheet identity through hidden tabs");
            var hiddenNavigationRejected = false;
            try { NavigateToCell(3, 0, 0); } catch (InvalidOperationException) { hiddenNavigationRejected = true; }
            check(hiddenNavigationRejected && _sheetIndex == 2, "Address navigation cannot activate a very-hidden worksheet");
            SheetPicker.SelectedIndex = 0; await Idle();
            check(_sheetIndex == 0 && CellAddress.Text == "C4" && ((ScaleTransform)SheetPageContent.LayoutTransform).ScaleX == .9, "Sheet switching restores each sheet's cell and zoom");
            GridSurface.Select(new(8, 4)); SheetPicker.SelectedIndex = 1; SheetPicker.SelectedIndex = 0; await Idle();
            check(CellAddress.Text == "E9" && !_document.IsDirty, "Navigation is remembered without marking workbook dirty");
            SetWorksheetZoom(125); await Idle();
            check(_document.IsDirty && ((ScaleTransform)SheetPageContent.LayoutTransform).ScaleX == 1.25, "Changing worksheet zoom is a persistent document edit");
            await UndoAsync(false); await Idle();
            check(((ScaleTransform)SheetPageContent.LayoutTransform).ScaleX == .9, "Undo restores the stored sheet zoom");
            SheetPicker.SelectedIndex = 1; await Idle();
            Freeze_Click(new Button { Tag = "clear" }, new()); await Idle();
            check(_grid.Control.FrozenRowCount == 0 && _grid.Control.FrozenColumnCount == 0, "Unfreeze command updates the grid and document");
            await UndoAsync(false); await Idle();
            check(_grid.Control.FrozenRowCount == 2 && _grid.Control.FrozenColumnCount == 1, "Undo restores both frozen axes");
            HideSheet_Click(this, new()); await Idle();
            check(SheetPicker.Items.Count == 1 && _sheetIndex == 0, "Hiding the active sheet moves to a visible sheet");
            var revision = _document.Snapshot.Revision; HideSheet_Click(this, new());
            check(_document.Snapshot.Revision == revision && SheetPicker.Items.Count == 1, "The last visible sheet cannot be hidden");
            _smokeTextDecision = "非常隐藏"; UnhideSheet_Click(this, new());
            check(_document.Snapshot.Sheets[3].Visibility == SheetVisibility.VeryHidden && SheetPicker.Items.Count == 1, "Ordinary unhide cannot expose a very-hidden sheet");
            _smokeTextDecision = "隐藏"; UnhideSheet_Click(this, new()); await Idle();
            check(_document.Snapshot.Sheets[1].Visibility == SheetVisibility.Visible && SheetPicker.Items.Count == 2, "Ordinary hidden sheet can be explicitly unhidden");
            check(await SaveAsync(false), "View and visibility changes save in place as XLSX");
            await OpenPathAsync(path); await Idle();
            check(_document.Snapshot.Sheets[2].Visibility == SheetVisibility.Hidden && _document.Snapshot.Sheets[3].Visibility == SheetVisibility.VeryHidden && SheetPicker.Items.Count == 2,
                "Reopening preserves both hidden states and tab visibility");
            GridSurface.Select(new(7, 6)); check(await SaveAsync(false), "Explicit save captures navigation even without a new content edit");
            await OpenPathAsync(path); await Idle();
            check(CellAddress.Text == "G8", "Saved navigation restores after closing and reopening");
            NavigateToCell(_sheetIndex, 500, 6); await Idle();
            check(!_document.IsDirty && _document.Snapshot.Sheets[_sheetIndex].View.TopLeftCell.Row > 400, "Scrolling to a distant cell is captured without creating an input edit");
            check(await SaveAsync(false), "Distant cell and viewport save together");
            await OpenPathAsync(path); await Idle();
            check(CellAddress.Text == "G501" && _grid.Control.GetModelCellBounds(500, 6) is not null, "Distant saved active cell remains visible on reopen");
        }
        finally { _smokeTextDecision = null; }
    }
}
