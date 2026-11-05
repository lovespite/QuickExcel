using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Domain;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunMergedBorderSmokeAsync(string directory, Action<bool, string> check)
    {
        async Task<SKBitmap> Capture(string name)
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var path = Path.Combine(directory, name + ".png");
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)_grid.Bounds.Width, (int)_grid.Bounds.Height), new Vector(96, 96));
            bitmap.Render(_grid); bitmap.Save(path, PngBitmapEncoderOptions.Default);
            return SKBitmap.Decode(path);
        }
        static CellPosition Neighbor(int side, int index) => side switch
        {
            0 => new(index, 0), 1 => new(index, 4), 2 => new(0, index), _ => new(4, index)
        };
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var styles = new List<StyleDefinition> { new() };
        for (var side = 0; side < 4; side++)
        foreach (var (index, kind, color) in new[] { (1, 1, "FFFF0000"), (3, 5, "FF0000FF") })
        {
            styles.Add(side switch
            {
                0 => new() { RightBorder = kind, RightBorderColor = color },
                1 => new() { LeftBorder = kind, LeftBorderColor = color },
                2 => new() { BottomBorder = kind, BottomBorderColor = color },
                _ => new() { TopBorder = kind, TopBorderColor = color }
            });
            cells[Neighbor(side, index)] = new(CellContent.FromLiteral(CellValue.Blank, styles.Count - 1));
        }
        var storage = new TypedWorkbookStorage();
        foreach (var frozen in new[] { false, true })
        {
            var sheet = new WorksheetSnapshot(SheetId.New(), "Merged borders", cells, mergedCells: ["B2:D4"], showGridLines: false,
                view: new() { FrozenRows = frozen ? 2 : 0, FrozenColumns = frozen ? 2 : 0, PaneTopLeftCell = frozen ? new(2, 2) : new(0, 0) });
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles);
            var file = Path.Combine(directory, $"merged-borders-{frozen}.xlsq"); storage.Save(file, book); await OpenPathAsync(file);
            _grid.Select(new(5, 5)); _grid.ScrollTo(frozen ? 15 : 0, frozen ? 12 : 0);
            using var pixels = await Capture($"merged-borders-{frozen}");
            for (var side = 0; side < 4; side++)
            for (var index = 1; index <= 3; index++)
            {
                var bounds = _grid.GetCellBounds(Neighbor(side, index))!.Value;
                var x = side switch { 0 => bounds.Right, 1 => bounds.Left, _ => bounds.Center.X };
                var y = side switch { 2 => bounds.Bottom, 3 => bounds.Top, _ => bounds.Center.Y };
                var line = Enumerable.Range(-2, 5).Select(offset => pixels.GetPixel((int)x + (side < 2 ? offset : 0), (int)y + (side >= 2 ? offset : 0))).ToArray();
                check(index switch
                {
                    1 => line.Any(p => p.Red > 200 && p.Green < 80 && p.Blue < 80),
                    2 => line.All(p => p == SKColors.White),
                    _ => line.Any(p => p.Blue > 200 && p.Red < 80 && p.Green < 80)
                }, $"Merged side {side}, segment {index} retains its own stroke/gap (frozen/scrolling={frozen})");
            }
        }
        // Optional local reproduction is copied to the isolated evidence directory;
        // regular smoke runs depend only on the generated fixture above.
        if (Environment.GetEnvironmentVariable("QUICKEXCEL_BORDER_SAMPLE") is not { Length: > 0 } sample) return;
        var copy = Path.Combine(directory, "bug-sample1.xlsq"); File.Copy(sample, copy);
        var source = storage.Open(copy); var valueFile = Path.Combine(directory, "bug-sample1-values.xlsx"); storage.Export(valueFile, source);
        foreach (var file in new[] { copy, valueFile })
        {
            await OpenPathAsync(file); _grid.Select(new(8, 2));
            foreach (var offset in new[] { 200, 240 })
            {
                _grid.ScrollTo(0, offset);
                using var pixels = await Capture($"{Path.GetFileNameWithoutExtension(file)}-scroll-{offset}");
                foreach (var row in new[] { 7, 8 })
                {
                    var bounds = _grid.GetCellBounds(new(row, 1))!.Value;
                    var line = Enumerable.Range(-1, 4).Select(dx => pixels.GetPixel((int)bounds.Left + dx, (int)bounds.Center.Y)).ToArray();
                    var hasBorder = line.Any(p => p.Red < 100 && p.Green < 100 && p.Blue < 100);
                    check(hasBorder == (row == 7), $"{Path.GetFileName(file)} B{row + 1} left border matches stored style at scroll {offset}");
                }
            }
        }
    }
}
