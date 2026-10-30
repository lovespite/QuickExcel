using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCellTextSmokeAsync(string folder, Action<bool, string> check)
    {
        var rich = new RichCellText([new("红色 Red", new("Microsoft YaHei", 18, true, false, "FFFF0000")),
            new(" / "), new("蓝色 Blue", new("Arial", 22, false, true, "FF0000FF"))]);
        var annotations = new WorksheetAnnotations([new(new(1, 0), new("https://example.invalid/no-request", Tooltip: "网页说明"))],
            [new(new(0, 0), new("审核", new([new("第一行批注\n第二行完整内容")])))]);
        var sheet = new WorksheetSnapshot(SheetId.New(), "内容显示", [new(new(0, 0), new(CellContent.FromRichText(rich))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("示例链接"), 1)))], rows: [new(0, new(55))], columns: [new(0, new(65))], annotations: annotations);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { FontColor = "FF0000FF", Underline = "single" }]);
        var path = Path.Combine(folder, "cell-text.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, book);
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        await OpenPathAsync(path); RibbonTabs.SelectedIndex = 0; GridSurface.Select(new(0, 0)); await Idle();
        check(FormulaBar.IsReadOnly && FormulaBar.Text == rich.Text && !GridSurface.BeginEdit(), "Imported rich text cannot be silently flattened by grid or formula-bar editing");
        var rect = _grid.Control.GetModelCellBounds(0, 0)!.Value; var dpi = VisualTreeHelper.GetDpi(_grid.Artwork);
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(_grid.Artwork.ActualWidth * dpi.DpiScaleX), (int)Math.Ceiling(_grid.Artwork.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        bitmap.Render(_grid.Artwork);
        var pixels = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4]; bitmap.CopyPixels(pixels, bitmap.PixelWidth * 4, 0); var red = 0; var blue = 0;
        for (var y = (int)(rect.Top * dpi.DpiScaleY) + 3; y < Math.Min(bitmap.PixelHeight, rect.Bottom * dpi.DpiScaleY) - 3; y++)
        for (var x = (int)(rect.Left * dpi.DpiScaleX) + 3; x < Math.Min(bitmap.PixelWidth, rect.Right * dpi.DpiScaleX) - 12; x++)
        { var p = (y * bitmap.PixelWidth + x) * 4; if (pixels[p + 2] > 180 && pixels[p] < 90 && pixels[p + 1] < 90) red++; if (pixels[p] > 180 && pixels[p + 1] < 90 && pixels[p + 2] < 90) blue++; }
        check(red > 50 && blue > 50, $"Actual grid paints both rich-text run colors (red={red}, blue={blue})");
        check(_grid.MeasureAutoFit(SheetAxis.Row, 0) > 28, "Automatic row sizing measures the larger rich-text font");
        _grid.UpdateCellInformation(new(rect.Left + 5, rect.Top + 5));
        check(_grid.Control.ToolTip is TextBlock { Text: var tip } && tip.Contains("批注 · 审核") && tip.Contains("第二行完整内容"), "Cell hover information includes the complete short traditional note");
        var linkRect = _grid.Control.GetModelCellBounds(1, 0)!.Value; _grid.UpdateCellInformation(new(linkRect.Left + 5, linkRect.Top + 5));
        check(_grid.Control.ToolTip is TextBlock { Text: var link } && link.Contains("https://example.invalid/no-request") && link.Contains("不自动访问"), "Hyperlink target and screen tip are shown as inert text");
        CaptureGridEvidence(Path.Combine(folder, "cell-text.png"));
        ApplyRibbonStyle(new(Bold: true)); await Idle();
        check(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText!.Runs.All(r => r.Font?.Bold == true), "Whole-cell font formatting applies to every rich-text run");
        await UndoAsync(false); await Idle();
        check(rich.Equals(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText), "Undoing whole-cell font formatting restores mixed run fonts");
        var menu = _grid.Control.ContextMenu!; menu.IsOpen = true; await Idle(); var flatten = (MenuItem)menu.Items[0];
        check(flatten.IsEnabled, "Rich-text conversion is available through an explicit context-menu action");
        flatten.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); menu.IsOpen = false; await Idle();
        check(_document.IsDirty && !FormulaBar.IsReadOnly && _document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText is null, "Explicit conversion keeps the text and enables ordinary editing");
        await UndoAsync(false); await Idle();
        check(FormulaBar.IsReadOnly && rich.Equals(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText), "One undo restores the complete run formatting");
        _editor.SetCell(0, 3, 3, "edited elsewhere"); RefreshSheet();
        check(await SaveAsync(false), "Ordinary XLSX save accepts modeled rich text, links and notes");
        await OpenPathAsync(path); await Idle();
        check(rich.Equals(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText) && _document.Snapshot.Sheets[0].Annotations.ContentEquals(annotations), "Desktop save/reopen preserves rich text and attachments");
        GridSurface.Select(new(3, 3)); await Idle(); check(!FormulaBar.IsReadOnly && GridSurface.BeginEdit(), "Moving to an ordinary cell restores editing"); GridSurface.Cancel();
    }
}
