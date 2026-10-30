using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunXlsxRenderingSmokeAsync(string folder, string? source)
    {
        Directory.CreateDirectory(folder);
        _suppressAutoCalculation = true;
        var cases = new List<string>();
        void Check(bool condition, string message)
        { if (!condition) throw new Exception(message); cases.Add(message); }
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        try
        {
            var sheet = new WorksheetSnapshot(SheetId.New(), "Rendering", new Dictionary<CellPosition, CellSnapshot>
            {
                [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromText("中文 Aa 123 · 18 pt"), 1)),
                [new(1, 0)] = new(CellContent.FromLiteral(CellValue.FromText("宽列内容仍可横向滚动查看"), 1))
            }, rows: [new(2, new(64))], columns: [new(0, new(149.109375))], defaultRowHeight: 30.4);
            var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles:
                [new(), new() { FontName = "Times New Roman", FontSize = 18, FontColor = "FFED7D31", FillColor = "FFF2F2F2", Vertical = 1 }]);
            await RunInitialSheetLayoutSmokeAsync(snapshot, Check);
            InstallDocument(new WorkbookDocument(snapshot, _services), null, "XLSX rendering regression");
            GridSurface.Select(new(3, 1)); await Idle();
            var rect = _grid.Control.GetModelCellBounds(0, 0)!.Value;
            Check(Math.Abs(rect.Width - (149.109375 * 7 + 5)) < 1, "Wide columns keep their stored width beyond 400 DIP");
            Check(Math.Abs(rect.Height - 30.4) < 1, "Row geometry uses model DIP without a second point conversion");
            var image = CaptureArtwork();
            var pixels = new byte[image.PixelWidth * image.PixelHeight * 4]; image.CopyPixels(pixels, image.PixelWidth * 4, 0);
            var dpi = VisualTreeHelper.GetDpi(_grid.Artwork);
            var coloredRows = new HashSet<int>();
            for (var y = (int)(rect.Top * dpi.DpiScaleY) + 2; y < Math.Min(image.PixelHeight, (int)(rect.Bottom * dpi.DpiScaleY) - 2); y++)
            for (var x = (int)(rect.Left * dpi.DpiScaleX) + 5; x < Math.Min(image.PixelWidth, (int)(rect.Right * dpi.DpiScaleX) - 5); x++)
            {
                var i = (y * image.PixelWidth + x) * 4;
                if (pixels[i + 2] > 180 && pixels[i + 1] < 190 && pixels[i] < 150) coloredRows.Add(y);
            }
            Save(image, "synthetic-artwork.png");
            Check(coloredRows.Count >= 14 * dpi.DpiScaleY, $"18 pt glyphs render at their full readable height in a 22.8 pt Excel row (colored rows: {coloredRows.Count})");
            Check(BeginGridResize(SheetAxis.Column, 0), "Wide column resize begins");
            PreviewGridResize(20); CompleteGridResize(false); await Idle();
            Check(Math.Abs(_grid.Control.GetModelCellBounds(0, 0)!.Value.Width - rect.Width - 20) < 1, "Resize starts from the actual wide column width");
            _document.Undo(); RefreshSheet(); await Idle();
            Check(Math.Abs(_grid.Control.GetModelCellBounds(0, 0)!.Value.Width - rect.Width) < 1, "Wide column resize remains undoable");

            _editor.SetCell(0, 3, 1, "=_xlfn.XLOOKUP(1,A1:A2,B1:B2)"); RefreshSheet();
            var inspectedSnapshot = _document.Snapshot;
            _smokeCompatibilityDriver = dialog =>
            {
                try
                {
                    var report = (System.Windows.Controls.TextBox)dialog.Content;
                    Check(report.IsReadOnly && report.Text.Contains("【内置不可计算／待验证】") && report.Text.Contains("Rendering · B4"),
                        "Actual compatibility dialog groups unsupported formulas with their current cell location");
                    Check(report.Text.Contains("【可编辑（已支持子集）】") && report.Text.Contains("当前无缓存结果"),
                        "Actual compatibility dialog separates editable content from unavailable calculation results");
                    File.WriteAllText(Path.Combine(folder, "compatibility-dialog.txt"), report.Text);
                }
                finally { dialog.Close(); }
            };
            await ShowCompatibilityAsync(); _smokeCompatibilityDriver = null;
            Check(ReferenceEquals(inspectedSnapshot, _document.Snapshot) && !_document.ScriptsEnabled,
                "Compatibility inspection leaves the snapshot and script authorization unchanged");
            _document.Undo(); RefreshSheet(); await Idle();

            await RunRowStyleInheritanceSmokeAsync(folder, Check);
            await RunNumberFormatSmokeAsync(folder, Check);
            await RunExtendedStyleSmokeAsync(folder, Check);
            await RunXlsxSaveSmokeAsync(folder, Check);
            await RunWorkbookViewSmokeAsync(folder, Check);
            await RunPreservedFormulaSmokeAsync(folder, Check);
            await RunCellTextSmokeAsync(folder, Check);
            await RunCompatibilityClosureSmokeAsync(folder, Check);
            await RunClipboardSmokeAsync(folder, Check);
            await RunFillSmokeAsync(folder, Check);

            if (source is not null)
            {
                var original = File.ReadAllBytes(source);
                InstallDocument(WorkbookDocument.Open(source, _services), null, Path.GetFileName(source));
                await RunXlsxViewportConsistencySmokeAsync(folder, Check);
                Check(!LogText.Text.Contains("未建模元素"), "Opening uses a concise compatibility explanation");
                File.WriteAllLines(Path.Combine(folder, "compatibility.txt"), _document.Snapshot.FileState!.CompatibilityIssues);
                var storage = new TypedWorkbookStorage();
                var preserved = Path.Combine(folder, "source-preserved.xlsq");
                storage.Save(preserved, _document.Snapshot);
                Check(storage.Open(preserved).FileState!.OriginalWorkbook.SequenceEqual(original),
                    "Saving an unchanged source preserves its entire XLSX payload");
                var subset = Path.Combine(folder, "supported-copy.xlsq");
                storage.Save(subset, _document.Snapshot, discardUnsupported: true);
                var reopened = storage.Open(subset);
                Check(reopened.Styles.SequenceEqual(_document.Snapshot.Styles) && reopened.Sheets.Zip(_document.Snapshot.Sheets)
                    .All(pair => pair.First.Cells.Count == pair.Second.Cells.Count &&
                        Math.Abs((pair.First.DefaultRowHeight ?? 0) - (pair.Second.DefaultRowHeight ?? 0)) < 1e-10),
                    "Explicit supported copy reopens with imported styles, cells and row geometry");
                for (var index = 0; index < _document.Snapshot.Sheets.Length; index++)
                {
                    SheetPicker.SelectedIndex = index; await Idle();
                    GridSurface.Select(new(0, 0)); await Idle();
                    CaptureGridEvidence(Path.Combine(folder, $"source-sheet-{index + 1}.png"));
                    var imported = _document.Snapshot.Sheets[index];
                    foreach (var column in imported.Columns.Where(c => c.Value.Width is > 57 && !c.Value.Hidden))
                    {
                        GridSurface.Select(new(1, column.Key)); await Idle();
                        var bounds = _grid.Control.GetModelCellBounds(1, column.Key)!.Value;
                        Check(Math.Abs(bounds.Width - Math.Clamp(column.Value.Width!.Value * 7 + 5, 1, 10000)) < 1,
                            $"Source sheet {index + 1}, column {column.Key + 1}: imported wide geometry retained");
                        CaptureGridEvidence(Path.Combine(folder, $"source-sheet-{index + 1}-column-{column.Key + 1}.png"));
                    }
                }
                Check(original.SequenceEqual(File.ReadAllBytes(source)), "Reproduction leaves the source XLSX unchanged");
            }
            File.WriteAllText(Path.Combine(folder, "xlsx-rendering-smoke-result.txt"), "PASS: " + cases.Count + " checks\n" + string.Join("\n", cases));
        }
        finally { _suppressAutoCalculation = false; _smokeCompatibilityDriver = null; }

        RenderTargetBitmap CaptureArtwork()
        {
            var artwork = _grid.Artwork; var dpi = VisualTreeHelper.GetDpi(artwork);
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(artwork.ActualWidth * dpi.DpiScaleX),
                (int)Math.Ceiling(artwork.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
            bitmap.Render(artwork); return bitmap;
        }
        void Save(BitmapSource bitmap, string name)
        {
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using var stream = File.Create(Path.Combine(folder, name)); encoder.Save(stream);
        }
    }
}
