using Avalonia.Controls;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly DispatcherTimer _calculationTimer = new() { Interval = TimeSpan.FromMilliseconds(180) };
    private DocumentSession? _observedSession;
    private bool _calculationPending, _windowClosed;
    private string? _calculationFailure;
    private CancellationTokenSource? _calculationCancellation, _statisticsCancellation;
    private readonly TextBlock _selectionStatistics = new() { Margin = new(12, 5) };
    private readonly Button _cancelCalculation = new() { Content = "停止计算  Esc", IsVisible = false, Margin = new(8, 0) };
    private void InstallCalculation()
    {
        _cancelCalculation.Click += (_, _) => _calculationCancellation?.Cancel();
        _calculationTimer.Tick += async (_, _) =>
        {
            if (_working || _workspace.IsBusy || _grid.IsEditing || _formulaDirty || _grid.IsComposing || _formula.IsComposing || _grid.IsSelectionDragging) return;
            _calculationTimer.Stop();
            if (_calculationPending && !_windowClosed) await PerformAsync(CalculatePendingAsync);
        };
        Closed += (_, _) =>
        {
            _windowClosed = true; _calculationTimer.Stop(); _calculationCancellation?.Cancel(); _statisticsCancellation?.Cancel();
            if (_observedSession is not null) _observedSession.Changed -= OnDocumentChanged;
        };
    }
    private void ObserveDocument()
    {
        if (ReferenceEquals(_observedSession, Document.Session)) return;
        if (_observedSession is not null) _observedSession.Changed -= OnDocumentChanged;
        _observedSession = Document.Session; _observedSession.Changed += OnDocumentChanged;
        _calculationTimer.Stop(); _calculationPending = false; _calculationFailure = null;
    }
    private void OnDocumentChanged(object? sender, DocumentChange change)
    {
        void Apply()
        {
            if (_windowClosed || !ReferenceEquals(sender, Document.Session)) return;
            RefreshScriptActions();
            if (!change.CalculationOnly && !change.CalculationUnchanged) { _calculationPending = true; _calculationTimer.Stop(); _calculationTimer.Start(); }
        }
        if (Dispatcher.UIThread.CheckAccess()) Apply(); else Dispatcher.UIThread.Post(Apply);
    }
    private Task CalculatePendingAsync() => CalculatePendingCoreAsync(false);
    private async Task CalculatePendingCoreAsync(bool forceFull)
    {
        if (!_calculationPending) return;
        _calculationTimer.Stop();
        using var cancellation = new CancellationTokenSource(); _calculationCancellation = cancellation; _cancelCalculation.IsVisible = true;
        try
        {
            if (await _workspace.CalculateAsync(cancellation.Token, forceFull)) { _calculationPending = false; _calculationFailure = null; _status.Text = "重算完成"; }
        }
        catch (Exception ex) { _calculationFailure = ex.Message; _status.Text = "计算未完成，保留输入及上次结果：" + ex.Message; }
        finally { _calculationCancellation = null; _cancelCalculation.IsVisible = false; QueueStatistics(); }
    }
    private async Task<bool> FlushEditsAsync()
    {
        if (!await CommitEditorsAsync()) return false;
        await CalculatePendingAsync(); return true;
    }
    private void QueueStatistics()
    {
        _statisticsCancellation?.Cancel(); if (_windowClosed) return;
        var cancellation = new CancellationTokenSource(); _statisticsCancellation = cancellation;
        var document = Document; var snapshot = document.Snapshot; var sheet = snapshot.ActiveSheetId;
        var ranges = _grid.SelectionRanges.Take(SelectionStatistics.MaxRanges + 1).Select(Rectangle).ToArray();
        _ = ComputeAsync();
        async Task ComputeAsync()
        {
            try
            {
                await Task.Delay(80, cancellation.Token);
                var result = await Task.Run(() => SelectionStatistics.Compute(snapshot, sheet, ranges, cancellationToken: cancellation.Token), cancellation.Token);
                if (cancellation.IsCancellationRequested || _windowClosed || !ReferenceEquals(document, Document) || Document.Snapshot.Revision != snapshot.Revision) return;
                string Number(double? value) => value?.ToString("G12") ?? "—";
                _selectionStatistics.Text = !result.Complete ? result.Limitation : result.Numbers == 0 ? $"计数：{result.Count:N0}" :
                    $"平均值：{Number(result.Average)}   计数：{result.Count:N0}   求和：{Number(result.Sum)}";
                if (result.OldCaches > 0 || result.MissingResults > 0) _selectionStatistics.Text += " · 含待重算结果";
                ToolTip.SetTip(_selectionStatistics, $"仅统计可见行列，重叠选区不重复计数。\n数值 {result.Numbers:N0}；错误 {result.Errors:N0}；旧缓存 {result.OldCaches:N0}；缺失结果 {result.MissingResults:N0}。");
            }
            catch (OperationCanceledException) { }
            catch (InvalidOperationException ex) { if (!cancellation.IsCancellationRequested) { _selectionStatistics.Text = "筛选待更新"; ToolTip.SetTip(_selectionStatistics, ex.Message); } }
            finally { if (ReferenceEquals(_statisticsCancellation, cancellation)) _statisticsCancellation = null; cancellation.Dispose(); }
        }
    }
}
