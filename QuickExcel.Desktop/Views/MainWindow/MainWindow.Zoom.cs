using System.Windows;
using System.Windows.Input;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private static readonly int[] ZoomStops = [10, 25, 50, 75, 100, 125, 150, 175, 200, 300, 400];

    private void InstallZoomHandlers()
    {
        // FastWpfGrid's own Ctrl+wheel handler changes CellFontSize, which desynchronizes
        // cell geometry from QuickExcel's snapshot projection. Handle it before the grid
        // sees it and use the same worksheet zoom path as the settings dialog.
        PreviewMouseWheel += (_, e) =>
        {
            if (e.Handled || _busy || _document is null || !_hasWorkbook ||
                (Keyboard.Modifiers & ModifierKeys.Control) == 0) return;
            if (e.OriginalSource is not DependencyObject target || !(ReferenceEquals(target, SheetGrid) || SheetGrid.IsAncestorOf(target))) return;
            e.Handled = true;
            try { ZoomByWheel(e.Delta > 0); }
            catch (Exception ex) { ReportEditFailure(ex); }
        };
    }

    private bool ZoomByWheel(bool increase)
    {
        if (_busy || _document is null) return false;
        var current = _document.Snapshot.Sheets[_sheetIndex].View.Zoom;
        var next = NextZoom(current, increase);
        if (next == current) { StatusText.Text = $"工作表缩放 {current}%（已达到缩放上限或下限）。"; return false; }
        if (!SetWorksheetZoom(next, commitEditors: false)) return false;
        StatusText.Text = $"工作表缩放 {next}%（Ctrl+滚轮调整，可撤销）。";
        return true;
    }

    private static int NextZoom(int current, bool increase) => increase
        ? ZoomStops.FirstOrDefault(stop => stop > current, ZoomStops[^1])
        : ZoomStops.LastOrDefault(stop => stop < current, ZoomStops[0]);
}
