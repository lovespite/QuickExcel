using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private CancellationTokenSource? _selectionStatisticsCancellation;
    private Task _selectionStatisticsTask = Task.CompletedTask;
    private void QueueSelectionStatistics()
    {
        _selectionStatisticsCancellation?.Cancel();
        if (_document is null || _windowClosed) return;
        var cancellation = new CancellationTokenSource(); _selectionStatisticsCancellation = cancellation;
        var document = _document; var snapshot = document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
        var ranges = GridSurface.SelectionRanges.Take(SelectionStatistics.MaxRanges + 1).Select(r => new CellRectangle(r.Top, r.Left, r.Bottom, r.Right)).ToArray();
        if (ranges.Length == 0) ranges = [new(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn)];
        try
        {
            var filter = CurrentView.CaptureRowFilter(snapshot, _sheetIndex);
            SelectionText.Text = "统计中…";
            _selectionStatisticsTask = RunAsync();
            async Task RunAsync()
            {
                try
                {
                    await Task.Delay(80, cancellation.Token);
                    var result = await Task.Run(() => SelectionStatistics.Compute(snapshot, sheet.Id, ranges, filter, cancellation.Token), cancellation.Token);
                    if (cancellation.IsCancellationRequested || _windowClosed || !ReferenceEquals(document, _document) || snapshot.Revision != _document.Snapshot.Revision) return;
                    string Number(double? value) => value?.ToString("G12", CultureInfo.CurrentCulture) ?? "—";
                    SelectionText.Text = !result.Complete ? result.Limitation : result.Numbers == 0 ? $"计数：{result.Count:N0}" :
                        $"平均值：{Number(result.Average)}   计数：{result.Count:N0}   求和：{Number(result.Sum)}";
                    if (result.Complete && (result.OldCaches > 0 || result.MissingResults > 0)) SelectionText.Text += " · 含待重算结果";
                    SelectionText.ToolTip = "仅统计选区内可见行列，排除手动隐藏和筛选隐藏；重叠选区不重复计数。\n" +
                        "计数包含文本、空文本、布尔及错误；求和/平均只计数值，不转换文本或布尔。\n" +
                        $"数值 {result.Numbers:N0}；错误 {result.Errors:N0}（不参与求和/平均）；旧缓存 {result.OldCaches:N0}；缺失结果 {result.MissingResults:N0}。" +
                        (result.Sum is null && result.Numbers > 0 ? "\n求和超出有限数值范围。" : "") +
                        (result.Limitation is null ? "" : "\n" + result.Limitation);
                }
                catch (OperationCanceledException) { }
                finally { if (ReferenceEquals(_selectionStatisticsCancellation, cancellation)) _selectionStatisticsCancellation = null; cancellation.Dispose(); }
            }
        }
        catch (InvalidOperationException ex)
        {
            _selectionStatisticsCancellation = null; cancellation.Dispose();
            SelectionText.Text = "筛选待更新，暂不统计"; SelectionText.ToolTip = ex.Message;
        }
    }
}
