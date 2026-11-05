using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunCellEditorAlignmentChecksAsync(string directory, Action<bool, string> check)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        var styles = new List<StyleDefinition> { new() };
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var vertical = 0; vertical < 3; vertical++)
        for (var horizontal = 1; horizontal <= 3; horizontal++)
        {
            styles.Add(new() { Horizontal = horizontal, Vertical = vertical });
            cells[new(vertical, horizontal - 1)] = new(CellContent.FromLiteral(CellValue.FromText("Align"), styles.Count - 1));
        }
        cells[new(3, 0)] = new(CellInput.Parse("Text"));
        cells[new(3, 1)] = new(CellInput.Parse("123"));
        cells[new(3, 2)] = new(CellInput.Parse("=1+2"));
        cells[new(4, 0)] = new(CellContent.FromLiteral(CellValue.FromText("Merged"), 5));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Editor alignment", cells,
            rows: Enumerable.Range(0, 7).ToDictionary(r => r, r => new RowMetadata(Height: 80, StyleId: r == 6 ? 5 : 0)),
            columns: new Dictionary<int, ColumnMetadata> { [3] = new(StyleId: 9) }, mergedCells: ["A5:C6"]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles);
        var path = Path.Combine(directory, "editor-alignment.xlsq");
        new TypedWorkbookStorage().Save(path, snapshot); await OpenPathAsync(path); await CalculateAsync(); await Layout();

        async Task CheckPixels(string name, int horizontal, int vertical)
        {
            var editor = _grid.CellEditor;
            // Remove the caret/selection from the bitmap so only the actual text determines its bounds.
            editor.SelectionStart = editor.SelectionEnd = editor.CaretIndex = 0; _grid.Focus(); await Layout();
            var file = Path.Combine(directory, "editor-" + name + ".png");
            using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(editor.Bounds.Width), (int)Math.Ceiling(editor.Bounds.Height)), new Vector(96, 96)))
            { image.Render(editor); image.Save(file, PngBitmapEncoderOptions.Default); }
            using var pixels = SKBitmap.Decode(file);
            var left = pixels.Width; var right = -1; var top = pixels.Height; var bottom = -1;
            for (var y = 2; y < pixels.Height - 2; y++)
            for (var x = 2; x < pixels.Width - 2; x++)
            {
                var color = pixels.GetPixel(x, y);
                if (color.Alpha < 200 || color.Red > 100 || color.Green > 100 || color.Blue > 100) continue;
                left = Math.Min(left, x); right = Math.Max(right, x); top = Math.Min(top, y); bottom = Math.Max(bottom, y);
            }
            check(right > left && bottom > top, $"{name}: native editor renders text");
            var xError = horizontal switch { 2 => Math.Abs((left + right + 1 - pixels.Width) / 2d), 3 => pixels.Width - right - 1, _ => left };
            var yError = vertical switch { 1 => Math.Abs((top + bottom + 1 - pixels.Height) / 2d), 2 => pixels.Height - bottom - 1, _ => top };
            check(xError <= 8 && yError <= 8, $"{name}: text pixels respect horizontal/vertical alignment (offsets {xError}, {yError})");
        }

        var revision = Document.Snapshot.Revision;
        for (var vertical = 0; vertical < 3; vertical++)
        for (var horizontal = 1; horizontal <= 3; horizontal++)
        {
            _grid.Select(new(vertical, horizontal - 1)); check(_grid.BeginEdit(), "Aligned cell enters native editing");
            await CheckPixels($"h{horizontal}-v{vertical}", horizontal, vertical); _grid.Cancel();
        }
        for (var column = 0; column < 3; column++)
        {
            _grid.Select(new(3, column)); _grid.BeginEdit();
            await CheckPixels("general-" + column, column == 0 ? 1 : 3, 2); _grid.Cancel();
        }
        _grid.Select(new(6, 1)); _grid.BeginEdit(); _grid.CellEditor.Text = "Row";
        await CheckPixels("inherited-row", 2, 1); _grid.Cancel();
        _grid.Select(new(0, 3)); _grid.BeginEdit(); _grid.CellEditor.Text = "Column";
        await CheckPixels("inherited-column", 3, 2); _grid.Cancel();
        _grid.Select(new(5, 2)); _grid.BeginEdit();
        check(_grid.Current == new CellPosition(4, 0) && _grid.CellEditor.Width > 300 && _grid.CellEditor.Height > 100,
            "Merged editor uses its anchor style across the merged rectangle");
        _grid.CellEditor.Text = "Merged\nTwo lines"; await CheckPixels("merged-multiline", 2, 1); _grid.Cancel();
        check(Document.Snapshot.Revision == revision, "Entering and cancelling aligned editors does not change workbook content or styles");

        _grid.Select(new(0, 0)); await ApplyStyleAsync(new(Horizontal: 3, Vertical: 2));
        _grid.BeginEdit(); await CheckPixels("style-refresh", 3, 2); _grid.Cancel();
        await History(false); _grid.BeginEdit(); await CheckPixels("style-undo", 1, 0); _grid.Cancel();
        await ZoomAsync(150); _grid.Select(new(1, 1)); _grid.BeginEdit();
        await CheckPixels("zoom-center", 2, 1);
        var presenter = _grid.CellEditor.GetVisualDescendants().OfType<TextPresenter>().Single();
        presenter.PreeditText = "zhong";
        check(_grid.IsComposing && !CommitEditors() && _grid.CellEditor.TextAlignment == TextAlignment.Center,
            "Centered native editor retains IME composition and commit gating");
        presenter.PreeditText = null; _grid.Cancel();
    }
}
