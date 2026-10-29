using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCellLayoutSmokeAsync(string folder, Action<bool, string> check, Func<int, int, Task> click)
    {
        async Task Idle() => await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
        static Rect Bounds(FrameworkElement element, Visual relative) => element.TransformToAncestor(relative).TransformBounds(new Rect(element.RenderSize));
        static bool Near(double x, double y) => Math.Abs(x - y) < 0.01;
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
                    var display = (TextBlock)SheetGrid.Columns[0].GetCellContent(_rows[row]);
                    var cell = (DataGridCell)display.Parent;
                    var before = Bounds(display, cell);
                    var baseline = display.BaselineOffset;
                    var textOrigin = display.TranslatePoint(display.ContentStart.GetCharacterRect(System.Windows.Documents.LogicalDirection.Forward).TopLeft, cell);
                    await click(row, 0);
                    var editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[row]);
                    var prepared = (TextBlock)editor.Template.FindName("PreparedText", editor);
                    var after = Bounds(prepared, cell);
                    check(_gridEditorArmed && prepared.IsVisible && before == after && Near(baseline, prepared.BaselineOffset),
                        $"Selection preserves text bounds and baseline: alignment {row}, zoom {zoom:P0}");
                    check(editor.IsKeyboardFocused && ((ScrollViewer)editor.Template.FindName("PART_ContentHost", editor)).Opacity == 0,
                        $"Prepared input remains focused at zoom {zoom:P0}, alignment {row}");
                    await click(row, 0); await Idle();
                    check(!prepared.IsVisible && editor.IsKeyboardFocused && editor.IsHitTestVisible,
                        $"Explicit editing exposes native text input at zoom {zoom:P0}, alignment {row}");
                    var editOrigin = editor.TranslatePoint(editor.GetRectFromCharacterIndex(0).TopLeft, cell);
                    check((textOrigin - editOrigin).Length < 0.01,
                        $"Explicit editing preserves text origin at zoom {zoom:P0}, alignment {row}: {textOrigin} -> {editOrigin}");
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
        var first = (DataGridCell)SheetGrid.Columns[0].GetCellContent(_rows[0]).Parent;
        var right = (DataGridCell)SheetGrid.Columns[1].GetCellContent(_rows[0]).Parent;
        var below = (DataGridCell)SheetGrid.Columns[0].GetCellContent(_rows[1]).Parent;
        check(Near(Bounds(first, SheetGrid).Right, Bounds(right, SheetGrid).Left) &&
            Near(Bounds(first, SheetGrid).Bottom, Bounds(below, SheetGrid).Top), "Adjacent cell containers tile with no horizontal or vertical gaps");

        // Inspect actual pixels, including the outer area beyond the text's padding.
        var blank = (DataGridCell)SheetGrid.Columns[1].GetCellContent(_rows[1]).Parent;
        var dpi = VisualTreeHelper.GetDpi(blank);
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(blank.ActualWidth * dpi.DpiScaleX),
            (int)Math.Ceiling(blank.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        var visual = new DrawingVisual();
        using (var drawing = visual.RenderOpen())
            drawing.DrawRectangle(new VisualBrush(blank), null, new Rect(blank.RenderSize));
        bitmap.Render(visual);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var stream = File.Create(Path.Combine(folder, "cell-fill.png"))) encoder.Save(stream);
        var pixels = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4];
        bitmap.CopyPixels(pixels, bitmap.PixelWidth * 4, 0);
        bool Filled(int x, int y)
        {
            var offset = (y * bitmap.PixelWidth + x) * 4;
            return pixels[offset] == 0x99 && pixels[offset + 1] == 0xEE && pixels[offset + 2] == 0xCC && pixels[offset + 3] == 255;
        }
        check(Enumerable.Range(3, bitmap.PixelWidth - 6).All(x => Filled(x, 3) && Filled(x, bitmap.PixelHeight - 4)) &&
            Enumerable.Range(3, bitmap.PixelHeight - 6).All(y => Filled(3, y) && Filled(bitmap.PixelWidth - 4, y)),
            "Cell fill reaches all four edges inside the border without padding seams");
        await click(0, 0);
        CaptureGridEvidence(Path.Combine(folder, "QuickExcel-cell-layout.png"));
    }
}
