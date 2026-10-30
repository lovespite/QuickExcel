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
                    var before = _grid.Control.GetModelCellBounds(row, 0);
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
                    check(Math.Abs(editor.Width - before!.Value.Width) < 1 && Math.Abs(editor.Height - before.Value.Height) < 1,
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
        var first = _grid.Control.GetModelCellBounds(0, 0)!.Value;
        var right = _grid.Control.GetModelCellBounds(0, 1)!.Value;
        var below = _grid.Control.GetModelCellBounds(1, 0)!.Value;
        check(Math.Abs(first.Right - right.Left) < 1 && Math.Abs(first.Bottom - below.Top) < 1,
            "Adjacent rendered cells tile without horizontal or vertical gaps");

        var artwork = _grid.Artwork;
        var dpi = VisualTreeHelper.GetDpi(artwork);
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(artwork.ActualWidth * dpi.DpiScaleX),
            (int)Math.Ceiling(artwork.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        bitmap.Render(artwork);
        var blank = _grid.Control.GetModelCellBounds(1, 1)!.Value;
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
