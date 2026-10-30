using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunNumberFormatSmokeAsync(string folder, Action<bool, string> check)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Formats", cells: [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(-1234.567), 1))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(1234.5), 2)))],
            columns: [new(0, new(Width: 45))], defaultRowHeight: 36);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(),
            new() { NumberFormatCode = "#,##0.000;[Red](#,##0.000)", FontSize = 16, FontColor = "FF0000FF" },
            new() { NumberFormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)", FontSize = 16 }]);
        InstallDocument(new WorkbookDocument(snapshot, _services), null, "数字格式验收");
        GridSurface.Select(new(8, 5));
        await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
        await WaitForFastGridPaintAsync(_grid.Control);
        check(_grid.DisplayText(0, 0) == "(1,234.567)" && _document.GetCellInput(0, 0, 0) == "-1234.567",
            "Grid uses formatted amounts while the formula editor retains the original numeric input");
        check(GridClipboard.Write(new(0, 0, 0, 0), (r, c) => _document.GetCellValue(0, r, c)) == "(1,234.567)",
            "Copy display values uses the same amount formatter as the grid");
        var width = _grid.MeasureAutoFit(SheetAxis.Column, 0);
        check(width > 100 && width < 300, "Auto-fit measures formatted content without expanding accounting fill to the previous column width");
        var image = new RenderTargetBitmap((int)Math.Ceiling(_grid.Artwork.ActualWidth), (int)Math.Ceiling(_grid.Artwork.ActualHeight), 96, 96, PixelFormats.Pbgra32);
        image.Render(_grid.Artwork);
        var pixels = new byte[image.PixelWidth * image.PixelHeight * 4]; image.CopyPixels(pixels, image.PixelWidth * 4, 0);
        var bounds = _grid.Control.GetModelCellBounds(0, 0)!.Value;
        var red = 0; var blue = 0;
        for (var y = (int)bounds.Top + 2; y < bounds.Bottom - 2; y++)
        for (var x = (int)bounds.Left + 2; x < bounds.Right - 2; x++)
        {
            var i = (y * image.PixelWidth + x) * 4;
            if (pixels[i + 2] > 160 && pixels[i + 1] < 100 && pixels[i] < 100) red++;
            if (pixels[i] > 160 && pixels[i + 1] < 100 && pixels[i + 2] < 100) blue++;
        }
        check(red > 20 && blue == 0, "Number-format red overrides the stored blue font in actual WPF glyph pixels");
        var accounting = _grid.Control.GetModelCellBounds(1, 0)!.Value;
        var inkLeft = false; var inkRight = false;
        for (var y = (int)accounting.Top + 2; y < accounting.Bottom - 2; y++)
        for (var x = (int)accounting.Left + 2; x < accounting.Right - 2; x++)
        {
            var i = (y * image.PixelWidth + x) * 4;
            if (pixels[i] > 100 || pixels[i + 1] > 100 || pixels[i + 2] > 100) continue;
            if (x < accounting.Left + 35) inkLeft = true;
            if (x > accounting.Right - 70) inkRight = true;
        }
        check(inkLeft && inkRight, "Accounting fill places the currency at the left and digits at the right of the cell");
        CaptureGridEvidence(Path.Combine(folder, "number-formats.png"));
    }
}
