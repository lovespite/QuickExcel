using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Primitives.PopupPositioning;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Popup? _chartFilterPopup;
    private Func<bool>? _chartFilterContextCurrent;
#if QUICKEXCEL_DIAGNOSTICS
    private Func<ChartFilterPanel, Task>? _chartFilterDriver;
#endif
    private void CloseChartFilter() { if (_chartFilterPopup is { } popup) popup.IsOpen = false; }
    private async Task ChartFilterAsync()
    {
        if (_grid.SelectedDrawingObject is not { Kind: DrawingObjectKind.Chart } selected) { _status.Text = "请先选中一个图表。"; return; }
        if (!await FlushEditsAsync()) return;
        CloseFilterPanel(); CloseChartFilter();
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var sheet = snapshot.GetSheet(snapshot.ActiveSheetId); var chart = sheet.Data.Charts[selected.Index];
        bool Current() => session == Document.Session.SessionId && snapshot.Revision == Document.Snapshot.Revision && sheet.Id == Document.Snapshot.ActiveSheetId && _grid.SelectedDrawingObject == selected;
        var panel = new ChartFilterPanel(snapshot, chart, _grid.CanEditDrawingObjects);
        var completion = new TaskCompletionSource<WorksheetChart?>(TaskCreationOptions.RunContinuationsAsynchronously); var selectData = false;
        var popup = new Popup { Child = panel, PlacementTarget = _grid, Placement = PlacementMode.AnchorAndGravity,
            PlacementRect = _grid.ChartFilterButtonBounds ?? new Rect(100, 40, 30, 30), PlacementAnchor = PopupAnchor.TopRight, PlacementGravity = PopupGravity.BottomRight,
            IsLightDismissEnabled = true };
        panel.Height = Math.Clamp(ClientSize.Height - 60, 340, 520);
        panel.Completed += value => { completion.TrySetResult(value); popup.IsOpen = false; };
        panel.SelectDataRequested += () => { selectData = true; completion.TrySetResult(null); popup.IsOpen = false; };
        popup.Closed += (_, _) => completion.TrySetResult(null);
        _chartFilterPopup = popup; _chartFilterContextCurrent = Current; _book.Children.Insert(0, popup);
        WorksheetChart? result;
        try
        {
            popup.IsOpen = true; panel.FocusSearch();
#if QUICKEXCEL_DIAGNOSTICS
            if (Program.SmokeDirectory is not null && _chartFilterDriver is { } driver) { _chartFilterDriver = null; await driver(panel); }
#endif
            result = await completion.Task;
        }
        finally { popup.IsOpen = false; popup.Child = null; _book.Children.Remove(popup); _chartFilterPopup = null; _chartFilterContextCurrent = null; }
        if (!Current()) return;
        if (selectData) { await ChartsAsync(); return; }
        if (result is not null && !chart.ContentEquals(result) && await ExecuteAsync(new SetChartCommand(sheet.Id, selected.Index, result), snapshot.Revision))
            _status.Text = "已应用图表筛选，可撤销；保存文件后保留。";
        _grid.Focus();
    }
}
