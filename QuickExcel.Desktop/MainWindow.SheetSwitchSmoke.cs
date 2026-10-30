using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunInitialSheetLayoutSmokeAsync(WorkbookSnapshot snapshot, Action<bool, string> check)
    {
        var surface = new FastGridSurface((_, _) => false);
        // Opening a file from the start page binds the snapshot before the grid has a viewport.
        surface.Bind(snapshot, 0, 100, 12);
        var host = new Window { Content = surface.Element, Width = 1100, Height = 640, ShowInTaskbar = false };
        try
        {
            host.Show();
            await Dispatcher.InvokeAsync(host.UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(surface.Control);
            var before = surface.Control.GetModelCellBounds(0, 0)!.Value;
            var beforeRow = surface.Control.GetModelCellBounds(2, 0)!.Value;
            var content = (Image)surface.Control.FindName("image");
            var beforeContent = CaptureStylePixels(content); var beforeArtwork = CaptureStylePixels(surface.Artwork);
            surface.Bind(snapshot, 0, 100, 12);
            await Dispatcher.InvokeAsync(host.UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(surface.Control);
            check(before == surface.Control.GetModelCellBounds(0, 0) && beforeRow == surface.Control.GetModelCellBounds(2, 0),
                "First viewport preserves imported dimensions before and after a sheet refresh");
            check(beforeContent.SequenceEqual(CaptureStylePixels(content)) && beforeArtwork.SequenceEqual(CaptureStylePixels(surface.Artwork)),
                "First viewport and subsequent sheet binding paint identical content and formatting");
        }
        finally { host.Close(); }
    }

    private async Task RunSheetSwitchSmokeAsync(string folder)
    {
        var cases = new List<string>();
        var plain = new WorksheetSnapshot(SheetId.New(), "Plain", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(1, 1)] = new(CellContent.FromLiteral(CellValue.FromText("Plain text")))
        }, rows: [new(2, new(Height: 62))], columns: [new(0, new(Width: 24))]);
        var styled = new WorksheetSnapshot(SheetId.New(), "Styled", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(1, 1)] = new(CellContent.FromLiteral(CellValue.FromText("Styled text"), 1)),
            [new(3, 2)] = new(CellContent.FromLiteral(CellValue.FromText("Merged text"), 1))
        }, rows: [new(5, new(StyleId: 1))], columns: [new(4, new(StyleId: 1))], mergedCells: ["C4:D5"]);
        var alternate = new WorksheetSnapshot(SheetId.New(), "Alternate", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(1, 1)] = new(CellContent.FromLiteral(CellValue.FromText("Other style"), 2))
        }, showGridLines: false, defaultRowHeight: 40);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [plain, styled, alternate], styles:
        [
            new(),
            new() { FillColor = "FFFFCC55", FontColor = "FFB00020", FontSize = 18, Bold = true, Italic = true,
                BottomBorder = 1, BottomBorderColor = "FFFF0000", Horizontal = 2, Vertical = 1 },
            new() { FillColor = "FF55CCFF", FontColor = "FF004499", FontSize = 12, Horizontal = 3 }
        ]);
        var path = Path.Combine(folder, "sheet-switch.xlsx");
        new TypedWorkbookStorage().Export(path, snapshot);
        InstallDocument(WorkbookDocument.Open(path, _services), null, "Sheet switch regression");
        var grid = _grid.Control;
        var content = (Image)grid.FindName("image");
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(grid);
        }
        byte[] Pixels(FrameworkElement element) => CaptureStylePixels(element);
        void Check(bool condition, string message)
        {
            if (!condition) throw new Exception("Sheet switch: " + message);
            cases.Add(message);
        }
        await RunInitialSheetLayoutSmokeAsync(_document.Snapshot, Check);
        await Idle();
        var plainContent = Pixels(content); var plainArtwork = Pixels(_grid.Artwork);
        SheetPicker.SelectedIndex = 1; await Idle();
        var styledContent = Pixels(content); var styledArtwork = Pixels(_grid.Artwork);
        CaptureGridEvidence(Path.Combine(folder, "sheet-switch-styled-initial.png"));
        Check(!plainContent.SequenceEqual(styledContent) && !plainArtwork.SequenceEqual(styledArtwork),
            "Styled sheet paints distinct cell content and formatting");
        void Scroll(string axis, double value)
        {
            var bar = (ScrollBar)grid.FindName(axis);
            bar.Value = axis == "hscroll" ? value * 140 : value;
            bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.ThumbPosition, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
        }
        foreach (var (row, column) in new[] { (2, 1), (20, 5), (0, 0) })
        {
            Scroll("vscroll", row); Scroll("hscroll", column); await Idle();
            var cachedContent = Pixels(content); var cachedArtwork = Pixels(_grid.Artwork);
            grid.InvalidateAll(); _grid.Artwork.InvalidateVisual(); await Idle();
            Check(cachedContent.SequenceEqual(Pixels(content)) && cachedArtwork.SequenceEqual(Pixels(_grid.Artwork)),
                $"Scrolled XLSX viewport {row}/{column} matches a fresh repaint");
        }
        Check(styledContent.SequenceEqual(Pixels(content)) && styledArtwork.SequenceEqual(Pixels(_grid.Artwork)),
            "Scrolling back restores the original XLSX formatting pixels");
        for (var pass = 0; pass < 2; pass++)
        {
            SheetPicker.SelectedIndex = 0; await Idle();
            CaptureGridEvidence(Path.Combine(folder, $"sheet-switch-plain-{pass}.png"));
            Check(plainContent.SequenceEqual(Pixels(content)), "Plain sheet restores its cell pixels on pass " + pass);
            Check(plainArtwork.SequenceEqual(Pixels(_grid.Artwork)), "Plain sheet clears previous fills, fonts, borders and merges on pass " + pass);
            SheetPicker.SelectedIndex = 2; await Idle();
            CaptureGridEvidence(Path.Combine(folder, $"sheet-switch-alternate-{pass}.png"));
            Check(Math.Abs(grid.GetModelCellBounds(1, 1)!.Value.Height - 40) < 1,
                "Alternate sheet applies its own dimensions on pass " + pass);
            SheetPicker.SelectedIndex = 1; await Idle();
            CaptureGridEvidence(Path.Combine(folder, $"sheet-switch-styled-return-{pass}.png"));
            Check(styledContent.SequenceEqual(Pixels(content)), "Styled sheet restores its cell pixels on pass " + pass);
            Check(styledArtwork.SequenceEqual(Pixels(_grid.Artwork)), "Styled sheet restores its own formatting on pass " + pass);
        }
        File.WriteAllLines(Path.Combine(folder, "sheet-switch-cases.txt"), cases);
    }

    private async Task RunXlsxViewportConsistencySmokeAsync(string folder, Action<bool, string> check)
    {
        var grid = _grid.Control;
        var content = (Image)grid.FindName("image");
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(grid);
        }
        byte[] Pixels(FrameworkElement element) => CaptureStylePixels(element);
        void Scroll(string axis, double value)
        {
            var bar = (ScrollBar)grid.FindName(axis); bar.Value = axis == "hscroll" ? value * 140 : value;
            bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.ThumbPosition, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
        }
        var baselines = new List<(byte[] Content, byte[] Artwork)>();
        for (var index = 0; index < _document.Snapshot.Sheets.Length; index++)
        {
            SheetPicker.SelectedIndex = index; GridSurface.Select(new(0, 0));
            Scroll("vscroll", 0); Scroll("hscroll", 0); await Idle();
            var baseline = (Content: Pixels(content), Artwork: Pixels(_grid.Artwork));
            baselines.Add(baseline);
            foreach (var (row, column) in new[] { (3, 1), (8, 4), (0, 8), (0, 0) })
            {
                Scroll("vscroll", row); Scroll("hscroll", column); await Idle();
                CaptureGridEvidence(Path.Combine(folder, $"viewport-{index}-{row}-{column}.png"));
                var cachedContent = Pixels(content); var cachedArtwork = Pixels(_grid.Artwork);
                grid.InvalidateAll(); _grid.Artwork.InvalidateVisual(); await Idle();
                check(cachedContent.SequenceEqual(Pixels(content)) && cachedArtwork.SequenceEqual(Pixels(_grid.Artwork)),
                    $"XLSX sheet {index}, viewport {row}/{column}: cached content and formatting match a full repaint");
            }
            check(baseline.Content.SequenceEqual(Pixels(content)) && baseline.Artwork.SequenceEqual(Pixels(_grid.Artwork)),
                $"XLSX sheet {index}: scrolling back preserves the original formatting pixels");
        }
        for (var index = 0; index < baselines.Count; index++)
        {
            SheetPicker.SelectedIndex = index; GridSurface.Select(new(0, 0)); await Idle();
            check(baselines[index].Content.SequenceEqual(Pixels(content)) && baselines[index].Artwork.SequenceEqual(Pixels(_grid.Artwork)),
                $"XLSX sheet {index}: switching back preserves the original formatting pixels");
        }
    }

    private static byte[] CaptureStylePixels(FrameworkElement element)
    {
        var dpi = VisualTreeHelper.GetDpi(element);
        var bitmap = new RenderTargetBitmap((int)Math.Ceiling(element.ActualWidth * dpi.DpiScaleX),
            (int)Math.Ceiling(element.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        bitmap.Render(element);
        var pixels = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4];
        bitmap.CopyPixels(pixels, bitmap.PixelWidth * 4, 0);
        return pixels;
    }
}
