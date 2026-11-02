using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCellLayoutSmokeAsync(string folder, Action<bool, string> check, Func<int, int, Task> click, Func<int, int, Task> doubleClick)
    {
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        Rect VisibleBounds(int row, int column) => _grid.Control.GetModelCellBounds(row, column)
            ?? throw new InvalidOperationException($"Cell layout fixture requires visible cell ({row}, {column}); " +
                $"scroll=({_grid.Control.FirstVisibleRowScrollIndex}, {_grid.Control.FirstVisibleColumnScrollIndex}), " +
                $"grid={_grid.Control.ActualWidth:0.##}x{_grid.Control.ActualHeight:0.##}, selection={_grid.Current}.");
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "网格布局回归.xlsq");
        ApplyValues([(0, 0, "中文 Aa 123"), (1, 0, "中文 Aa 123"), (2, 0, "多行\nAa 123")]);
        var id = _document.Snapshot.Sheets[0].Id;
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Row, 0, 4, 64));
        for (var row = 0; row < 3; row++)
            ExecuteWorkbookCommand(new SetCellStyleCommand([new(id, row, 0)],
                new(FontSize: row == 1 ? 18 : 11, Bold: row == 1, Vertical: row, Horizontal: row + 1, WrapText: row == 2)));
        var transform = SheetPageContent.LayoutTransform;
        try
        {
            foreach (var zoom in new[] { 1d, 1.25, 1.5 })
            {
                SheetPageContent.LayoutTransform = new ScaleTransform(zoom, zoom);
                for (var row = 0; row < 3; row++)
                {
                    GridSurface.Select(new(3, 2)); await Idle();
                    _grid.Control.ScrollModelIntoView(new(row, 0)); await Idle();
                    var before = VisibleBounds(row, 0);
                    await click(row, 0); await Idle();
                    var editor = _grid.Editor;
                    check(_gridEditorArmed && !editor.IsHitTestVisible && editor.Foreground == Brushes.Transparent &&
                        before == _grid.Control.GetModelCellBounds(row, 0),
                        $"Selection preserves content geometry and uses a transparent native editor: alignment {row}, zoom {zoom:P0}");
                    check(editor.IsKeyboardFocused, $"Prepared input remains focused at zoom {zoom:P0}, alignment {row}");
                    await doubleClick(row, 0); await Idle();
                    check(_grid.IsEditing && editor.IsKeyboardFocused && editor.IsHitTestVisible &&
                        editor.FontSize == (row == 1 ? 18 : 11) * 96d / 72d &&
                        editor.TextWrapping == (row == 2 ? TextWrapping.Wrap : TextWrapping.NoWrap),
                        $"Explicit editor uses workbook size and wrapping at zoom {zoom:P0}, alignment {row}");
                    check(Math.Abs(editor.Width - before.Width) < 1 && Math.Abs(editor.Height - before.Height) < 1,
                        "Native editor fits the same grid cell bounds");
                    await HandleGridKeyAsync(System.Windows.Input.Key.Escape, System.Windows.Input.ModifierKeys.None, true);
                }
            }
        }
        finally { SheetPageContent.LayoutTransform = transform; }
        GridSurface.Select(new(3, 2)); await Idle();
        ExecuteWorkbookCommand(new SetCellStyleCommand(
            (from row in Enumerable.Range(0, 3) from column in Enumerable.Range(0, 2) select new CellAddress(id, row, column)),
            new(FillColor: "FFCCEE99", Border: true)));
        await Idle();
        // Zoom and selection can leave A1 above the viewport. Exercise that state
        // explicitly, then restore the sampled region without selecting its fill cells.
        _grid.Control.ScrollModelIntoView(new(80, 8)); await Idle();
        check(_grid.Control.GetModelCellBounds(0, 0) is null, "Layout fixture starts with the sampled region outside the viewport");
        _grid.Control.ScrollModelIntoView(new(0, 0)); await Idle();
        check(_grid.Current == new CellPosition(3, 2), "Restoring the sampled viewport preserves the selection outside the fill cells");
        var first = VisibleBounds(0, 0);
        var right = VisibleBounds(0, 1);
        var below = VisibleBounds(1, 0);
        check(Math.Abs(first.Right - right.Left) < 1 && Math.Abs(first.Bottom - below.Top) < 1,
            "Adjacent rendered cells tile without horizontal or vertical gaps");
        // The artwork draws from the control origin; a centred (explicitly sized)
        // raster viewport would shift every styled cell by ~1 DIP (~2px at 200%).
        var origin = _grid.Control.RasterOrigin;
        check(Math.Abs(origin.X) < 0.01 && Math.Abs(origin.Y) < 0.01,
            $"Raster viewport shares the control origin with the artwork layer ({origin.X:0.##},{origin.Y:0.##})");

        var artwork = _grid.Artwork;
        var dpi = VisualTreeHelper.GetDpi(artwork);
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(artwork.ActualWidth * dpi.DpiScaleX),
            (int)Math.Ceiling(artwork.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        bitmap.Render(artwork);
        var blank = VisibleBounds(1, 1);
        var crop = new CroppedBitmap(bitmap, new Int32Rect((int)(blank.Left * dpi.DpiScaleX), (int)(blank.Top * dpi.DpiScaleY),
            (int)(blank.Width * dpi.DpiScaleX), (int)(blank.Height * dpi.DpiScaleY)));
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(crop));
        using (var stream = File.Create(Path.Combine(folder, "cell-fill.png"))) encoder.Save(stream);
        var pixels = new byte[crop.PixelWidth * crop.PixelHeight * 4];
        crop.CopyPixels(pixels, crop.PixelWidth * 4, 0);
        bool Filled(int x, int y)
        {
            var offset = (y * crop.PixelWidth + x) * 4;
            return pixels[offset] == 0x99 && pixels[offset + 1] == 0xEE && pixels[offset + 2] == 0xCC && pixels[offset + 3] == 255;
        }
        check(Enumerable.Range(3, crop.PixelWidth - 6).All(x => Filled(x, 3) && Filled(x, crop.PixelHeight - 4)) &&
            Enumerable.Range(3, crop.PixelHeight - 6).All(y => Filled(3, y) && Filled(crop.PixelWidth - 4, y)),
            "Cell fill reaches all four edges inside the border without padding seams");
        await click(0, 0);
        CaptureGridEvidence(Path.Combine(folder, "QuickExcel-cell-layout.png"));
    }
}
