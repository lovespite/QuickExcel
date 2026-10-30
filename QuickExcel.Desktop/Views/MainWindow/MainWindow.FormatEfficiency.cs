using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private WorkbookClipboard? _formatBrush;
    private void CancelFormatBrush()
    {
        _formatBrush = null;
        if (RibbonFormatBrush is not null) RibbonFormatBrush.Content = "格式刷";
    }
    private void FormatBrush_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var snapshot = _document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
            CurrentView.RequireCurrentFilter(snapshot);
            if (_formatBrush is null)
            {
                if (snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
                _formatBrush = WorkbookClipboard.Capture(snapshot, sheet.Id, new(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn));
                RibbonFormatBrush.Content = "应用格式";
                StatusText.Text = "已采样当前单元格格式。选择目标区域后再次点击“应用格式”；Esc 取消。";
                return;
            }
            var range = SelectedRange();
            if (range.Count > WorkbookClipboard.MaxCells) throw new InvalidOperationException("格式刷单次上限为 20,000 格。");
            CurrentView.RequireVisibleTargets(snapshot, _sheetIndex, Enumerable.Range(range.Top, range.Rows));
            var brush = _formatBrush; CancelFormatBrush();
            if (ExecuteWorkbookCommand(new PasteCellsCommand(brush, sheet.Id, new(range.Top, range.Left, range.Bottom, range.Right), brush.Revision, CellPasteMode.Formats)))
                StatusText.Text = "已应用格式，内容和公式保持原值，可一次撤销。";
        }
        catch (Exception ex) { CancelFormatBrush(); ReportEditFailure(ex); }
    }

    private void DecimalPlaces_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var snapshot = _document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
            CurrentView.RequireCurrentFilter(snapshot);
            var selection = GridSurface.Selection.Count == 0 ? [new CellPosition(_selectedRow, _selectedColumn)] : GridSurface.Selection;
            var addresses = selection.Where(p => !CurrentView.IsFilteredOut(snapshot, _sheetIndex, p.Row))
                .Select(p => new CellAddress(sheet.Id, p.Row, p.Column));
            var delta = (sender as FrameworkElement)?.Tag as string == "increase" ? 1 : -1;
            if (ExecuteWorkbookCommand(new AdjustDecimalPlacesCommand(addresses, delta))) StatusText.Text = "已调整显示小数位，数值和公式保持原值。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { UpdateRibbonState(); }
    }
}
