using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly DispatcherTimer _eventTimer = new() { Interval = TimeSpan.FromMilliseconds(100) };
    private void InstallScriptEvents()
    {
        _grid.CellClicked += () => QueueCellEvent("cell.click");
        _grid.DraftStarted += () => QueueCellEvent("cell.edit");
        Activated += (_, _) => { if (_book.IsVisible) Document.Session.RaiseScriptEvent("workbook.activate"); };
        Deactivated += (_, _) => { if (!_windowClosed && _book.IsVisible) Document.Session.RaiseScriptEvent("workbook.deactivate"); };
        _eventTimer.Tick += async (_, _) =>
        {
            if (_working || _execution.IsRunning || _workspace.IsBusy || _scriptCancellation is not null || _windowClosed || _closePending ||
                _grid.IsEditing || _formulaDirty || _grid.IsComposing || _formula.IsComposing || _grid.IsSelectionDragging || !Document.Session.HasScriptEvents) return;
            await PerformAsync(DispatchEventsAsync);
        };
        Closed += (_, _) => _eventTimer.Stop();
        _eventTimer.Start();
    }
    private void QueueCellEvent(string name, SheetId? sheet = null, CellPosition? cell = null)
    {
        if (_synchronizing || _windowClosed) return;
        var ranges = cell is { } position ? [new CellRectangle(position.Row, position.Column, position.Row, position.Column).ToString()] :
            _grid.SelectionRanges.Take(64).Select(r => Rectangle(r).ToString()).ToArray();
        Document.Session.RaiseScriptEvent(name, sheet ?? Document.Snapshot.ActiveSheetId, ranges);
    }
    private async Task DispatchEventsAsync()
    {
        if (!Document.Session.HasScriptEvents) return;
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitScriptsAsync() || !Document.Session.HasScriptEvents) return;
        var document = Document;
        _calculationTimer.Stop();
        using var cancellation = new CancellationTokenSource();
        using var feedback = BeginExecution("正在执行事件代码…", cancellation);
        _scriptCancellation = cancellation;
        try { await Task.Run(() => document.DispatchScriptEvents(cancellation.Token), cancellation.Token); }
        finally
        {
            _scriptCancellation = null;
            await RefreshScriptPageAfterMutationAsync();
            Synchronize(); RefreshScriptDiagnostics();
            if (_calculationPending) _calculationTimer.Start();
        }
    }
}
