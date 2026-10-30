using System.Globalization;
using System.IO;
using System.IO.Compression;
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
    private async Task RunDateDisplaySmokeAsync(string folder)
    {
        var serial = ExcelSerialDate.FromDateTime(new(2026, 10, 27, 13, 5, 9), ExcelDateSystem.Windows1900).Serial;
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var labels = new[] { "日期", "时间", "日期与时间", "累计时长", "公式日期", "普通数值" };
        for (var i = 0; i < labels.Length; i++)
            cells[new(i, 0)] = new(CellContent.FromLiteral(CellValue.FromText(labels[i])));
        cells[new(0, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(serial), 1));
        cells[new(1, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(serial % 1), 2));
        cells[new(2, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(serial), 3));
        cells[new(3, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(1.5), 4));
        cells[new(4, 1)] = new(CellContent.FromFormula("=B1+1", 1), new(CellValue.FromNumber(serial + 1), CalculationState.Stale, 0, 0, 0));
        cells[new(5, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(12345)));
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "日期时间", cells, columns: [new KeyValuePair<int, ColumnMetadata>(1, new(32))])],
            styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }, new() { NumberFormatId = 21 },
                new() { NumberFormatCode = "yyyy-mm-dd hh:mm:ss" }, new() { NumberFormatId = 46 }]);
        var seed = Path.Combine(folder, "dates-source.xlsq");
        new TypedWorkbookStorage().Save(seed, snapshot);
        var xlsx = Path.Combine(folder, "dates-source.xlsx");
        using (var zip = ZipFile.OpenRead(seed)) zip.GetEntry("workbook.xlsx")!.ExtractToFile(xlsx);
        await OpenPathAsync(xlsx);
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        string[] expected = ["2026-10-27", "13:05:09", "2026-10-27 13:05:09", "36:00:00", "2026-10-28", "12345"];
        for (var row = 0; row < expected.Length; row++)
        {
            SheetGrid.UpdateLayout();
            if (_grid.DisplayText(row, 1) != expected[row])
                throw new Exception($"Date display mismatch in actual QuickExcel.Grid.Wpf row {row}.");
        }
        GridSurface.Select(new(0, 1));
        SheetGrid.Focus();
        if (!GridSurface.BeginEdit() || ((System.Windows.FrameworkElement)_grid.Editor) is not TextBox editor ||
            editor.Text != serial.ToString("R", CultureInfo.InvariantCulture))
            throw new Exception("Date editor must preserve the full numeric input.");
        if (!CommitEdit()) throw new Exception("Unchanged date edit did not commit.");
        if (_document.Snapshot.Sheets[0].GetCell(0, 1)!.Value != CellValue.FromNumber(serial))
            throw new Exception("Unchanged date edit changed the numeric value.");
        _filePath = Path.Combine(folder, "dates-saved.xlsq");
        if (!await SaveAsync(false)) throw new Exception("Date smoke save failed.");
        using (var reopened = WorkbookDocument.Open(_filePath, _services))
            for (var row = 0; row < expected.Length; row++)
                if (reopened.GetCellValue(0, row, 1) != expected[row]) throw new Exception("Date display changed after save/reopen.");

        RefreshSheet();
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        UpdateLayout();
        var dpi = VisualTreeHelper.GetDpi(this);
        var bitmap = new RenderTargetBitmap((int)(ActualWidth * dpi.DpiScaleX), (int)(ActualHeight * dpi.DpiScaleY),
            96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(this);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var file = File.Create(Path.Combine(folder, "QuickExcel-dates.png"))) encoder.Save(file);
        File.WriteAllText(Path.Combine(folder, "date-display-smoke-result.txt"),
            "PASS: XLSX opened through desktop OpenPath; actual QuickExcel.Grid.Wpf shows date/time/datetime/elapsed/formula/general values; edit commit preserves numeric precision; XLSQ save/reopen preserves display and numbers.");
    }
}
