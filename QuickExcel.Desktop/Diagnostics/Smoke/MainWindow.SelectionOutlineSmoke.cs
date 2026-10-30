using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FastWpfGrid;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunSelectionOutlineSmokeAsync(string folder)
    {
        var cases = new List<string>();
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "选框重叠回归.xlsq");
        ApplyValues([(0, 0, "7"), (0, 1, "8"), (1, 0, "11"), (1, 1, "12")]);
        RefreshSheet();
        var grid = _grid.Control;
        var overlay = (FrameworkElement)grid.FindName("selectionOverlay");
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(grid);
        }
        IEnumerable<GeometryDrawing> Shapes(Drawing drawing) => drawing is DrawingGroup group
            ? group.Children.SelectMany(Shapes) : drawing is GeometryDrawing geometry ? [geometry] : [];
        void CheckOutlines(string scenario, params GridCellRange[] ranges)
        {
            var expected = ranges.SelectMany(grid.GetModelRangeFragments).Select(f => Rect.Intersect(f.Bounds, f.Clip)).ToArray();
            var actual = Shapes(VisualTreeHelper.GetDrawing(overlay)).Where(d => d.Pen is not null)
                .Select(d => d.Geometry.Bounds).ToArray();
            if (actual.Length != expected.Length || expected.Any(e => actual.Count(r =>
                Math.Abs(e.Left - r.Left) < 1 && Math.Abs(e.Top - r.Top) < 1 &&
                Math.Abs(e.Width - r.Width) < 1 && Math.Abs(e.Height - r.Height) < 1) != 1))
                throw new Exception("Selection outline: " + scenario + ": expected one outline per visible range, without nested or duplicate borders.");
            cases.Add(scenario);
        }
        byte[] GridPixels()
        {
            var dpi = VisualTreeHelper.GetDpi(_grid.Element);
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(_grid.Element.ActualWidth * dpi.DpiScaleX),
                (int)Math.Ceiling(_grid.Element.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
            bitmap.Render(_grid.Element);
            var pixels = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4];
            bitmap.CopyPixels(pixels, bitmap.PixelWidth * 4, 0); return pixels;
        }
        async Task CheckEditorPixels(bool editing)
        {
            if (!_grid.Editor.IsKeyboardFocused) throw new Exception("Editor must be focused to exercise native focus chrome.");
            var actual = GridPixels(); var opacity = _grid.Editor.Opacity;
            byte[] hidden;
            try { _grid.Editor.Opacity = 0; await Idle(); hidden = GridPixels(); }
            finally { _grid.Editor.Opacity = opacity; await Idle(); }
            if (actual.SequenceEqual(hidden) == editing)
                throw new Exception("Selection outline: native editor must draw pixels only while editing.");
            cases.Add(editing ? "Active native editor remains visible" : "Focused prepared editor contributes no pixels or theme border");
        }
        var rectangle = new GridCellRange(0, 0, 1, 1);
        GridSurface.SelectRange(new(0, 0, 1, 1), new(1, 1)); await Idle();
        CaptureGridEvidence(Path.Combine(folder, "selection-outline.png"));
        CheckOutlines("2x2 selection with bottom-right current cell", rectangle);
        await CheckEditorPixels(false);
        if (_grid.Current != new QuickExcel.Domain.CellPosition(1, 1) || FormulaBar.Text != "12")
            throw new Exception("Removing the nested outline must preserve the current editing address and formula bar.");
        cases.Add("Current cell and formula bar remain B2 / 12");
        GridSurface.BeginEdit(); await Idle(); await CheckEditorPixels(true); GridSurface.Cancel(); await Idle();
        await CheckEditorPixels(false);
        var drag = typeof(FastGridControl).GetMethod("SetSelectedRectangle", BindingFlags.Instance | BindingFlags.NonPublic)!
            .CreateDelegate<Action<FastGridCellAddress, FastGridCellAddress>>(grid);
        drag(new(0, 0), new(1, 1)); await Idle(); CheckOutlines("Forward drag", rectangle);
        drag(new(1, 1), new(0, 0)); await Idle(); CheckOutlines("Reverse drag", rectangle);
        GridSurface.Select(new(0, 0));
        await NavigateGridAsync(Key.Right, ModifierKeys.Shift); await NavigateGridAsync(Key.Down, ModifierKeys.Shift); await Idle();
        CheckOutlines("Shift+arrows selection", rectangle);
        GridSurface.Select(new(1, 1)); await Idle();
        CheckOutlines("Single cell draws its border once", new GridCellRange(1, 1, 1, 1));
        var remove = typeof(FastGridControl).GetMethod("RemoveSelectedCell", BindingFlags.Instance | BindingFlags.NonPublic)!
            .CreateDelegate<Action<FastGridCellAddress>>(grid);
        var changed = typeof(FastGridControl).GetMethod("OnChangeSelectedCells", BindingFlags.Instance | BindingFlags.NonPublic)!
            .CreateDelegate<Action<bool>>(grid);
        remove(grid.CurrentCell); changed(true); await Idle();
        CheckOutlines("Ctrl-deselected current cell retains a focus outline", new GridCellRange(1, 1, 1, 1));
        SetSmokeFreeze(1, 1); RefreshSheet();
        GridSurface.SelectRange(new(0, 0, 1, 1), new(1, 1)); await Idle();
        CheckOutlines("Frozen panes do not duplicate the current-cell outline", rectangle);
        SetSmokeFreeze(0, 0); RefreshSheet();
        File.WriteAllLines(Path.Combine(folder, "selection-outline-cases.txt"), cases);
    }
}
