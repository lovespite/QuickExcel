using System.Windows;
using System.Windows.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private SheetId[] _visibleSheetIds = [];
    private bool _restoringWorkbookView;
    private void RememberWorksheetViewport()
    {
        if (_refreshing || _busy || _restoringWorkbookView || !_hasWorkbook) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        if (sheet.View.FrozenRows > 10 || sheet.View.FrozenColumns >= 256 || sheet.View.ActiveCell.Column >= 256) return;
        if (_grid.Control.GetScrollingModelCell() is { Row: { } row, Column: { } column } && row >= 0 && column >= 0)
            _document.Session.UpdateNavigation(sheet.Id, sheet.View.ActiveCell, new(row, column));
    }
    private void PopulateVisibleSheets()
    {
        var sheets = _document.Snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Visible).ToArray();
        _visibleSheetIds = sheets.Select(s => s.Id).ToArray();
        SheetPicker.ItemsSource = sheets.Select(s => s.Name).ToList();
        SheetPicker.SelectedIndex = Array.IndexOf(_visibleSheetIds, _document.Snapshot.Sheets[_sheetIndex].Id);
    }
    private void EnsureVisibleSheet()
    {
        var snapshot = _document.Snapshot;
        if (_sheetIndex >= snapshot.Sheets.Length || snapshot.Sheets[_sheetIndex].Visibility != SheetVisibility.Visible)
            _sheetIndex = snapshot.Sheets.IndexOf(snapshot.GetSheet(snapshot.ActiveSheetId));
    }
    private void ApplyWorksheetZoom() => SheetPageContent.LayoutTransform = new ScaleTransform(
        _document.Snapshot.Sheets[_sheetIndex].View.Zoom / 100d, _document.Snapshot.Sheets[_sheetIndex].View.Zoom / 100d);

    private void RestoreWorksheetNavigation()
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var view = sheet.View;
        _restoringWorkbookView = true;
        try
        {
            var cell = new CellPosition(Math.Clamp(view.ActiveCell.Row, 0, _rowCount - 1), Math.Clamp(view.ActiveCell.Column, 0, _columnCount - 1));
            if (view.Selection is { } selection && selection.Range.Bottom < _rowCount && selection.Range.Right < _columnCount)
                GridSurface.SelectRange(new(selection.Range.Top, selection.Range.Left, selection.Range.Bottom, selection.Range.Right), cell);
            else GridSurface.Select(cell);
            var topLeft = view.HasFreeze ? view.PaneTopLeftCell : view.TopLeftCell;
            _grid.Control.RestoreModelScroll(topLeft.Row, topLeft.Column);
            _selectedRow = _grid.Current.Row; _selectedColumn = _grid.Current.Column;
            UpdateFormulaBar();
        }
        finally { _restoringWorkbookView = false; }
        if (!sheet.Data.Tables.IsEmpty) RefreshWorkspaceView();
        if (view.ActiveCell.Column >= 256 || view.FrozenRows > 10 || view.FrozenColumns >= 256)
            StatusText.Text = "视图超出当前网格显示范围；原视图仍保留，冻结显示上限为 10 行及可见列范围。";
    }

    private bool SetWorksheetZoom(int zoom, bool commitEditors = true)
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var command = new SetSheetViewCommand(sheet.Id, sheet.View with { Zoom = zoom, NormalZoom = zoom });
        if (commitEditors) return ExecuteWorkbookCommand(command);
        // Ctrl+wheel changes only the view; keep an in-progress cell/formula draft uncommitted.
        if (_busy) return false;
        try
        {
            var result = _editor.Execute(command);
            if (!result.Changed) return true;
            RefreshSheet(); UpdateChrome();
            return true;
        }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private void HideSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        ExecuteWorkbookCommand(new SetSheetVisibilityCommand(_document.Snapshot.Sheets[_sheetIndex].Id, SheetVisibility.Hidden));
    }
    private void UnhideSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var hidden = _document.Snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Hidden).ToArray();
        if (hidden.Length == 0) { StatusText.Text = "没有可取消隐藏的普通工作表。"; return; }
        var name = PromptText("取消隐藏工作表", "输入以下工作表名称：" + string.Join("、", hidden.Select(s => s.Name)), hidden[0].Name);
        if (name is null) return;
        var selected = hidden.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (selected is null) { ReportEditFailure(new ArgumentException("只能取消隐藏列表中的普通工作表。")); return; }
        ExecuteWorkbookCommand(new SetSheetVisibilityCommand(selected.Id, SheetVisibility.Visible));
    }
}
