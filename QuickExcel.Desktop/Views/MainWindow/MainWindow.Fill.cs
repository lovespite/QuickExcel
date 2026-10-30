using System.Windows;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private (Guid Session, SheetId Sheet, long Revision)? _fillLease;
    private Action<SeriesDialog>? _smokeSeriesDriver;
    private void InstallFillHandlers()
    {
        _grid.BeforeFill = () =>
        {
            if (_busy || !CanCommitComposition() || !CommitEditors() || _document.Snapshot.FileState?.FormulaReadOnly == true) return false;
            _fillLease = (_document.Session.SessionId, _document.Snapshot.Sheets[_sheetIndex].Id, _document.Snapshot.Revision); return true;
        };
        _grid.FillRequested += async (source, extent) =>
        {
            try
            {
                if (_fillLease is not { } lease || lease.Session != _document.Session.SessionId || lease.Sheet != _document.Snapshot.Sheets[_sheetIndex].Id)
                    throw new InvalidOperationException("填充期间已切换文档或工作表，请重新拖动。");
                await (_gridInput = ApplyFillAsync(source, extent, lease.Revision));
            }
            catch (Exception ex) { ReportEditFailure(ex); }
            finally { _fillLease = null; }
        };
        Deactivated += (_, _) => _grid.CancelFillGesture();
    }
    private async Task ApplyFillAsync(GridRange source, GridRange extent, long? expectedRevision = null)
    {
        if (_busy || !CanCommitComposition() || !CommitEditors()) return;
        CurrentView.RequireVisibleTargets(_document.Snapshot, _sheetIndex, Enumerable.Range(extent.Top, extent.Rows));
        var result = _editor.Execute(new FillCellsCommand(_document.Snapshot.Sheets[_sheetIndex].Id,
            new(source.Top, source.Left, source.Bottom, source.Right), new(extent.Top, extent.Left, extent.Bottom, extent.Right)), expectedRevision);
        await FinishFillAsync(result, extent, "已复制填充，公式按每格位置调整，可一次撤销。");
    }
    private async Task FillSelectionAsync(bool right)
    {
        if (_busy || !CanCommitComposition() || !CommitEditors()) return;
        var range = GridRange.Rectangle(GridSurface.Selection);
        var source = right ? range with { Right = range.Left } : range with { Bottom = range.Top };
        var extent = range;
        if (right && range.Columns == 1)
        {
            if (range.Left == 0) { StatusText.Text = "首列左侧没有可复制的单元格。"; return; }
            source = range with { Left = range.Left - 1, Right = range.Left - 1 };
            extent = range with { Left = source.Left };
        }
        else if (!right && range.Rows == 1)
        {
            if (range.Top == 0) { StatusText.Text = "首行上方没有可复制的单元格。"; return; }
            source = range with { Top = range.Top - 1, Bottom = range.Top - 1 };
            extent = range with { Top = source.Top };
        }
        await ApplyFillAsync(source, extent);
        GridSurface.SelectRange(range, new(range.Top, range.Left));
    }
    private async Task FinishFillAsync(CommandResult result, GridRange extent, string message)
    {
        if (!result.Changed) { StatusText.Text = "选区没有需要填充或清除格式的内容。"; return; }
        _calculationPending |= !result.CalculationUnchanged; RefreshSheet(true); UpdateChrome();
        GridSurface.SelectRange(extent, new(extent.Top, extent.Left));
        if (_calculationPending) await RecalculateAsync();
        StatusText.Text = message;
    }
    private async void Fill_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition() || !CommitEditors()) return;
        try
        {
            var action = (sender as FrameworkElement)?.Tag as string;
            if (action is "down" or "right") { await (_gridInput = FillSelectionAsync(action == "right")); return; }
            var range = GridRange.Rectangle(GridSurface.Selection);
            CurrentView.RequireVisibleTargets(_document.Snapshot, _sheetIndex, Enumerable.Range(range.Top, range.Rows));
            var sheet = _document.Snapshot.Sheets[_sheetIndex]; var revision = _document.Snapshot.Revision;
            DocumentCommand command;
            if (action == "clear-formats") command = new ClearCellFormatsCommand(GridSurface.Selection.Select(p => new CellAddress(sheet.Id, p.Row, p.Column)));
            else
            {
                if (range.Rows > 1 && range.Columns > 1) throw new ArgumentException("序列需要单行或单列选区。");
                var dialog = new SeriesDialog(this);
                if (_smokeSeriesDriver is { } driver) dialog.Loaded += (_, _) => Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => driver(dialog)));
                if (dialog.ShowDialog() != true) return;
                command = new FillSeriesCommand(sheet.Id, new(range.Top, range.Left, range.Bottom, range.Right), dialog.Kind, dialog.Step);
            }
            var result = _editor.Execute(command, revision);
            await (_gridInput = FinishFillAsync(result, range, action == "clear-formats" ? "已清除单元格格式，内容与分段字体保持，可一次撤销。" : "已填充序列，可一次撤销。"));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
