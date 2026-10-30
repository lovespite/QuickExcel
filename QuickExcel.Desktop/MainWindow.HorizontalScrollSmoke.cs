using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunHorizontalScrollSmokeAsync(string folder)
    {
        var cases = new List<string>();
        void Check(bool condition, string message)
        {
            if (!condition) throw new Exception("Horizontal scroll: " + message);
            cases.Add(message);
        }
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < 30; row++)
        for (var column = 0; column < 30; column++)
            cells[new(row, column)] = new(CellContent.FromLiteral(CellValue.FromText($"内容 {row}:{column}"), 1));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Scroll", cells,
            columns: [new(3, new(Hidden: true)), new(4, new(Width: 45))], mergedCells: ["B2:C3"]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles:
            [new(), new() { FontSize = 13, Bold = true, WrapText = true, FillColor = "FFFFF2CC", FontColor = "FF234567", BottomBorder = 1 }]);
        var surface = new FastGridSurface((_, _) => false);
        var grid = surface.Control;
        var host = new Window { Owner = this, Width = 960, Height = 580, ShowInTaskbar = false, Content = surface.Element };
        var horizontal = (ScrollBar)grid.FindName("hscroll");
        var vertical = (ScrollBar)grid.FindName("vscroll");
        async Task Idle() => await WaitForFastGridPaintAsync(grid);
        static void Scroll(ScrollBar bar, double position, ScrollEventType type = ScrollEventType.ThumbTrack)
        {
            bar.Value = position;
            bar.RaiseEvent(new ScrollEventArgs(type, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
        }
        async Task FreshPixelsMatch(string message)
        {
            var content = (Image)grid.FindName("image");
            var before = CaptureStylePixels(content); var artwork = CaptureStylePixels(surface.Artwork);
            grid.InvalidateAll(); surface.Artwork.Refresh(); await Idle();
            Check(before.SequenceEqual(CaptureStylePixels(content)) && artwork.SequenceEqual(CaptureStylePixels(surface.Artwork)), message);
        }
        try
        {
            host.Show(); surface.Bind(snapshot, 0, 100, 30); await Idle();
            var dpi = VisualTreeHelper.GetDpi(grid);
            Check(horizontal.ViewportSize == grid.GridScrollAreaWidth, "Horizontal scrollbar extent uses device pixels");
            var initial = grid.GetModelCellBounds(5, 1)!.Value;
            Scroll(horizontal, 1); await Idle();
            Check(grid.FirstVisibleColumnScrollIndex == 0 && Math.Abs(initial.Left - grid.GetModelCellBounds(5, 1)!.Value.Left - 1 / dpi.DpiScaleX) < .001,
                "One device pixel of thumb movement translates content without snapping to the next column");
            var draws = surface.Artwork.CellDrawingCount; var reads = surface.ReadCount;
            var notifications = 0;
            grid.ViewportChanged += (_, _) => notifications++;
            Scroll(horizontal, 2); await Idle();
            Check(notifications == 1, "A changed thumb position emits one viewport notification");
            Check(surface.Artwork.CellDrawingCount == draws && surface.ReadCount == reads,
                "Scrolling within the same visible cells reuses content and styled text without model reads or layout");
            var contentSource = ((Image)grid.FindName("image")).Source;
            Scroll(horizontal, 2, ScrollEventType.EndScroll); await Idle();
            Check(notifications == 1 && ReferenceEquals(contentSource, ((Image)grid.FindName("image")).Source),
                "EndScroll at the same position does not repaint or notify");
            notifications = 0;
            for (var i = 3; i < 103; i++) Scroll(horizontal, i);
            Check(notifications == 0 && grid.IsRenderPending, "A burst of thumb input is queued without synchronous viewport rebuilds");
            await Idle();
            Check(notifications == 1 && horizontal.Value == 102,
                "A burst of one hundred thumb positions paints only the latest viewport");
            Scroll(horizontal, 160); Scroll(horizontal, 180);
            Scroll(horizontal, 180, ScrollEventType.EndScroll);
            Check(notifications == 2 && horizontal.Value == 180,
                "Releasing a queued thumb drag commits the final position immediately");
            await Idle();
            Check(notifications == 2, "A stopped scroll timer cannot replay an obsolete position after release");
            var render = grid.GetType().GetMethod("HorizontalScrollRendering", BindingFlags.Instance | BindingFlags.NonPublic)!;
            static RenderingEventArgs Frame(TimeSpan time) => (RenderingEventArgs)Activator.CreateInstance(typeof(RenderingEventArgs),
                BindingFlags.Instance | BindingFlags.NonPublic, null, [time], null)!;
            var frameTime = TimeSpan.FromDays(1);
            Scroll(horizontal, 200);
            var beforeFrame = ((Image)grid.FindName("image")).Source;
            render.Invoke(grid, [null, Frame(frameTime)]);
            Check(notifications == 3 && !ReferenceEquals(beforeFrame, ((Image)grid.FindName("image")).Source),
                "A composition frame submits the latest thumb position and content together");
            Scroll(horizontal, 210); Scroll(horizontal, 220);
            render.Invoke(grid, [null, Frame(frameTime)]);
            Check(notifications == 3, "Repeated callbacks for one composition frame do not replay thumb input");
            render.Invoke(grid, [null, Frame(frameTime + TimeSpan.FromMilliseconds(16))]);
            Check(notifications == 4 && Math.Abs(initial.Left - grid.GetModelCellBounds(5, 1)!.Value.Left - 220 / dpi.DpiScaleX) < .001,
                "The following frame uses the newest input and discards superseded positions");
            await Idle();
            Check(notifications == 4, "Completed frame work leaves no timer replay");
            Scroll(horizontal, 300);
            surface.Select(new(0, 0)); await Idle();
            Check(horizontal.Value == 0, "Keyboard/programmatic navigation cancels an older pending thumb request");
            Scroll(horizontal, 2); await Idle();
            var x = grid.GetModelCellBounds(5, 1)!.Value.Left;
            Scroll(vertical, 2); await Idle();
            Check(grid.GetModelCellBounds(5, 1)!.Value.Left == x && horizontal.Value == 2,
                "Vertical scrolling preserves the partial horizontal offset");
            var bounds = grid.GetModelCellBounds(5, 1)!.Value;
            Check(grid.HitTestModel(new(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2)) is { Row: 5, Column: 1 },
                "Hit testing follows fractional horizontal scrolling");
            await FreshPixelsMatch("Cached fractional text, fills, borders and merged cells match a fresh paint");

            surface.View.FrozenRows = surface.View.FrozenColumns = 1; surface.Refresh(snapshot); await Idle();
            Scroll(vertical, 0); Scroll(horizontal, 0); await Idle();
            var frozen = grid.GetModelCellBounds(0, 0);
            Scroll(horizontal, 37); await Idle();
            Check(grid.GetModelCellBounds(0, 0) == frozen && !grid.GetVisibleModelColumns().Contains(3),
                "Frozen intersection remains stationary and hidden columns stay excluded");
            await FreshPixelsMatch("Fractionally clipped frozen/merged content matches a fresh paint");
            surface.Select(new(4, 1)); surface.BeginEdit(); await Idle();
            Check(surface.Editor.Visibility == Visibility.Visible && surface.Editor.Clip is not null,
                "Native editor uses the visible pane clip after fractional scrolling");
            surface.Cancel();
            var image = new RenderTargetBitmap((int)Math.Ceiling(surface.Element.ActualWidth * dpi.DpiScaleX),
                (int)Math.Ceiling(surface.Element.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
            image.Render(surface.Element);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image));
            using (var output = File.Create(Path.Combine(folder, "horizontal-scroll.png"))) encoder.Save(output);
            for (var i = 0; i < 15; i++) { Scroll(horizontal, i * 130); await Idle(); }
            Check(surface.Artwork.CachedCellDrawingCount <= grid.GetVisibleModelRows().Count * grid.GetVisibleModelColumns().Count + sheet.MergedCells.Length,
                "Styled drawing cache stays bounded to the viewport during a long drag");
            Scroll(horizontal, horizontal.Maximum); await Idle();
            var last = grid.GetModelCellBounds(0, 29)!.Value;
            var right = ((Image)grid.FindName("image")).ActualWidth;
            Check(Math.Abs(last.Right - right) < 1, "Rightmost thumb position reveals the final column without blank overscroll");
            host.Width += 160; host.UpdateLayout(); await Idle();
            Check(Math.Abs(grid.GetModelCellBounds(0, 29)!.Value.Right - ((Image)grid.FindName("image")).ActualWidth) < 1,
                "Widening the viewport clamps the partial rightmost scroll position");

            Scroll(horizontal, 0); Scroll(vertical, 0); await Idle();
            var oldArtwork = CaptureStylePixels(surface.Artwork);
            var changed = snapshot.WithSheets([sheet.WithCells(sheet.Cells.SetItem(new(0, 0),
                new(CellContent.FromLiteral(CellValue.FromText("已更新"), 1))))]);
            surface.Refresh(changed); await Idle();
            Check(!oldArtwork.SequenceEqual(CaptureStylePixels(surface.Artwork)), "Snapshot updates invalidate cached text");
            await FreshPixelsMatch("Updated cached artwork matches a fresh paint");
            grid.UseSmoothScrolling = false; await Idle();
            Scroll(horizontal, 2); await Idle();
            Check(grid.FirstVisibleColumnScrollIndex == 2 && horizontal.Value == 2, "Legacy bitmap scrollbar retains item-based behavior");
            grid.UseSmoothScrolling = true; await Idle();
            Check(horizontal.ViewportSize == grid.GridScrollAreaWidth && horizontal.Value > 2,
                "Switching back restores pixel scrollbar units at the current column");
            var hugeSheet = new WorksheetSnapshot(SheetId.New(), "Oversized", cells,
                columns: [new(0, new(Width: 1400))], rows: [new(0, new(Height: 1000))]);
            surface.View = new();
            surface.Bind(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [hugeSheet], styles: snapshot.Styles), 0, 30, 30); await Idle();
            Scroll(horizontal, 20); await Idle();
            draws = surface.Artwork.CellDrawingCount;
            Scroll(horizontal, 21); await Idle();
            Check(surface.Artwork.CellDrawingCount == draws, "Oversized styled cells reuse a bounded vector drawing without huge bitmap allocation");
            await FreshPixelsMatch("Oversized styled content remains correctly clipped while scrolling");
            // A small model must not retain B2:C3 from the larger fixture.
            var shortSnapshot = snapshot.WithSheets([new WorksheetSnapshot(SheetId.New(), "Short")]);
            surface.View = new(); surface.Bind(shortSnapshot, 0, 2, 2); await Idle();
            Check(horizontal.Maximum == 0 && horizontal.Value == 0 && grid.FirstVisibleColumnScrollIndex == 0,
                "Binding a short sheet resets the horizontal range and offset");
            surface.Bind(snapshot, 0, 100, 30); await Idle();
            Scroll(horizontal, 200);
            host.Content = null;
            await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            host.Content = surface.Element; await Idle();
            Check(grid.FirstVisibleColumnScrollIndex == 0 && grid.GetModelCellBounds(5, 1)!.Value.Left == initial.Left,
                "Unloading cancels pending thumb input before the surface is reused");
            File.WriteAllLines(Path.Combine(folder, "horizontal-scroll-cases.txt"), cases);
        }
        finally { host.Close(); }
    }
}
