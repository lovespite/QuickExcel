using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FastWpfGrid;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private static async Task WaitForFastGridPaintAsync(FastGridControl control, Action<string>? observe = null)
    {
        var painted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        void Completed(object? sender, EventArgs args) => painted.TrySetResult();
        control.RenderCompleted += Completed;
        try
        {
            await control.Dispatcher.InvokeAsync(() => observe?.Invoke("probe-observed"), DispatcherPriority.Render);
            observe?.Invoke(control.IsRenderPending ? "probe-pending" : "probe-completed");
            if (!control.IsRenderPending) return;
            // Observe the completed render pass directly. Polling every 16 ms can
            // add another timer interval after painting has already finished.
            await painted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally { control.RenderCompleted -= Completed; }
    }

    // Invoke the same selection/hover path as mouse movement without moving the user's OS pointer.
    private static Action<int, int> DragSelectionProbe(FastGridControl control)
    {
        T Bind<T>(string name) where T : Delegate => typeof(FastGridControl)
            .GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<T>(control);
        var select = Bind<Action<FastGridCellAddress, FastGridCellAddress>>("SetSelectedRectangle");
        var hoverRow = Bind<Action<int?>>("SetHoverRow");
        var hoverCell = Bind<Action<FastGridCellAddress>>("SetHoverCell");
        return (row, column) => { select(new(0, 0), new(row, column)); hoverRow(row); hoverCell(new(row, column)); };
    }

    private async Task VerifyFastGridRegressionAsync(Action<bool, string> check)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Small", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(1))),
            [new(1, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(5)))
        });
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var surface = new FastGridSurface((_, _) => true);
        // Preserve the original bitmap-path regressions; exercise the retained
        // pixel-scrolling path separately below against that visual baseline.
        surface.Control.UseSmoothScrolling = false;
        async Task Idle() => await WaitForFastGridPaintAsync(surface.Control);
        var host = new Window { Owner = this, Content = surface.Element, Width = 960, Height = 500, ShowInTaskbar = false };
        host.Show();
        try
        {
            surface.Bind(snapshot, 0, 2, 4); await Idle();
            surface.View.SetRegion(new(Guid.NewGuid(), sheet.Id, new(0, 0, 1, 3)));
            surface.Select(new(1, 3));
            surface.View.Apply(snapshot, 0, new(First: new(FilterOperator.GreaterOrEqual, "10")));
            surface.Refresh(snapshot); await Idle();
            var grid = surface.Control;
            check(surface.Current == new CellPosition(0, 3), "Filtering the last data row moves the editor to the header");
            check(grid.GetCellAddress(new(grid.HeaderWidth + 10, grid.HeaderHeight + 150)).IsEmpty,
                "Blank space below a filtered two-row sheet cannot become a negative model row");
            check(grid.GetCellAddress(new(grid.HeaderWidth + 800, grid.HeaderHeight + 5)).IsEmpty,
                "Blank space to the right cannot become an out-of-range model column");
            check(grid.VisibleRowCount == 1, "Filtered short grid renders only its remaining visible row");
            surface.View.FrozenRows = 1; surface.Refresh(snapshot); await Idle();
            check(grid.VisibleRowCount == 0 && grid.GetModelCellBounds(0, 0) is not null,
                "Frozen header remains usable when every data row is filtered out");
            surface.View.FrozenRows = 0; surface.View.ClearConditions(); surface.Refresh(snapshot); await Idle();
            surface.Select(new(1, 3));
            check(surface.Current == new CellPosition(1, 3), "Clearing filters restores the last original cell address");

            surface.View = new(); surface.Bind(snapshot, 0, 100_000, 256); await Idle();
            var drag = DragSelectionProbe(grid);
            drag(4, 2); await Idle();
            var notifications = 0; surface.SelectionChanged += () => notifications++;
            var reads = surface.ReadCount;
            for (var i = 0; i < 100; i++) drag(4, 2);
            await Idle();
            check(notifications == 0 && surface.ReadCount == reads,
                "One hundred pointer moves inside the same selected cell do no repeated selection/editor work");
            reads = surface.ReadCount; grid.InvalidateAll(); await Idle(); var fullReads = surface.ReadCount - reads;
            reads = surface.ReadCount; drag(5, 2); await Idle();
            check(surface.ReadCount - reads < fullReads / 2,
                "Extending a drag by one row redraws changed cells instead of the entire viewport");
            byte[] Pixels()
            {
                var bitmap = (WriteableBitmap)typeof(FastGridControl).GetField("_drawBuffer", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)!;
                return ReadPixels(bitmap);
            }
            static byte[] ReadPixels(BitmapSource bitmap)
            {
                var stride = bitmap.PixelWidth * 4; var bytes = new byte[stride * bitmap.PixelHeight];
                bitmap.CopyPixels(bytes, stride, 0); return bytes;
            }
            BitmapSource Frame() => (BitmapSource)((Image)grid.FindName("image")).Source;
            var overlay = (FrameworkElement)grid.FindName("selectionOverlay");
            BitmapSource OverlayFrame()
            {
                var frame = new RenderTargetBitmap((int)Math.Ceiling(overlay.ActualWidth), (int)Math.Ceiling(overlay.ActualHeight),
                    96, 96, PixelFormats.Pbgra32);
                frame.Render(overlay); return frame;
            }
            bool OverlayContains(int row, int column)
            {
                var bounds = grid.GetModelCellBounds(row, column);
                if (bounds is null) return false;
                var rect = bounds.Value; var frame = OverlayFrame(); var pixel = new byte[4];
                frame.CopyPixels(new Int32Rect((int)(rect.Left + rect.Width / 2), (int)(rect.Top + rect.Height / 2), 1, 1), pixel, 4, 0);
                return pixel[3] > 0;
            }
            var priorFrame = Frame(); var priorPixels = ReadPixels(priorFrame);
            foreach (var endpoint in new[] { (2, 1), (8, 4), (0, 0), (5, 2) })
            {
                drag(endpoint.Item1, endpoint.Item2); await Idle(); var partial = Pixels();
                grid.InvalidateAll(); await Idle();
                check(partial.SequenceEqual(Pixels()), $"Incremental drag rendering matches a full repaint at {endpoint}");
            }
            for (var i = 0; i < 100; i++) drag(i % 10, i % 5);
            await Idle(); var coalesced = Pixels(); grid.InvalidateAll(); await Idle();
            check(coalesced.SequenceEqual(Pixels()) && surface.Current == new CellPosition(9, 4),
                "Queued drag changes retain the latest selection and match a complete repaint");
            check(Frame() is not WriteableBitmap && Frame().IsFrozen && ReadPixels(Frame()).SequenceEqual(Pixels()),
                "The displayed immutable frame matches the private software canvas");
            check(ReadPixels(priorFrame).SequenceEqual(priorPixels),
                "Later selection drawing cannot mutate a frame still in use by the compositor");
            drag(99_000, 250); await Idle(); var unchangedFrame = Frame();
            drag(99_001, 251); await Idle();
            check(ReferenceEquals(unchangedFrame, Frame()),
                "Moving a large selection edge outside the viewport reuses the unchanged frame");

            void Field(string name, object value) => typeof(FastGridControl).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(grid, value);
            var tick = typeof(FastGridControl).GetMethod("_dragTimer_Tick", BindingFlags.Instance | BindingFlags.NonPublic)!
                .CreateDelegate<Action<object, EventArgs>>(grid);
            surface.Select(new(0, 0)); await Idle();
            Field("_dragStartCell", new FastGridCellAddress(0, 0));
            var contentFrame = Frame(); reads = surface.ReadCount;
            drag(9, 9); await Idle(); var tenByTen = ReadPixels(OverlayFrame());
            drag(19, 19); await Idle();
            check(ReferenceEquals(contentFrame, Frame()) && surface.ReadCount == reads,
                "Dragging from 10x10 to 20x20 reuses content pixels and does not read cell values");
            check(!tenByTen.SequenceEqual(ReadPixels(OverlayFrame())) && OverlayContains(9, 3) && !OverlayContains(21, 3),
                "The independent overlay expands and clips the selected rectangle correctly");
            var overlayBeforeFullPaint = ReadPixels(OverlayFrame()); grid.InvalidateAll(); await Idle();
            check(overlayBeforeFullPaint.SequenceEqual(ReadPixels(OverlayFrame())),
                "Retained overlay pixels agree with a full content and overlay refresh");
            drag(2, 2); await Idle();
            check(OverlayContains(1, 1) && !OverlayContains(4, 1), "Shrinking the overlay removes the old selection fill");
            Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
            Field("_mouseMoveRow", (int?)1); Field("_mouseMoveColumn", (int?)1);
            var selectionEvents = 0; grid.SelectedCellsChanged += CountSelection;
            void CountSelection(object? sender, FastWpfGrid.SelectionChangedEventArgs e) => selectionEvents++;
            try
            {
                for (var i = 0; i < 3; i++) tick(grid, EventArgs.Empty);
                await Idle();
                check(surface.Current == new CellPosition(grid.FirstVisibleRowScrollIndex + grid.VisibleRowCount - 1,
                    grid.FirstVisibleColumnScrollIndex + grid.VisibleColumnCount - 1) && selectionEvents == 3,
                    "Diagonal auto-scroll advances both edges with one selection update per tick");
                check(surface.Editor.Visibility != Visibility.Visible && ReadPixels(Frame()).SequenceEqual(Pixels()),
                    "Auto-scroll keeps the editor hidden and presents the latest selection");
                var scrolled = Pixels(); grid.InvalidateAll(); await Idle();
                check(scrolled.SequenceEqual(Pixels()), "Auto-scroll pixels match a full repaint");
            }
            finally
            {
                grid.SelectedCellsChanged -= CountSelection;
                Field("_dragStartCell", FastGridCellAddress.Empty);
                Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
            }
            host.Width = 1100; host.Height = 650;
            surface.Element.LayoutTransform = new ScaleTransform(1.5, 1.5); await Idle();
            check(ReadPixels(Frame()).SequenceEqual(Pixels()), "Resizing and 150% zoom present a correctly sized complete frame");
            var dpi = VisualTreeHelper.GetDpi(grid);
            var viewport = (FrameworkElement)grid.FindName("imageGrid");
            var buffer = (WriteableBitmap)typeof(FastGridControl).GetField("_drawBuffer", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)!;
            check(Math.Abs(grid.RenderScaleX - dpi.DpiScaleX * 1.5) < 0.01 && Math.Abs(grid.RenderScaleY - dpi.DpiScaleY * 1.5) < 0.01 &&
                Math.Abs(buffer.PixelWidth - (int)Math.Ceiling(((int)viewport.ActualWidth - 2) * grid.RenderScaleX)) <= 1,
                "Layout zoom drives the raster density instead of resampling a window-DPI bitmap");
            host.Content = null; await Idle(); host.Content = surface.Element; await Idle();
            check(ReadPixels(Frame()).SequenceEqual(Pixels()), "Reloading the grid refreshes the presented frame");
            surface.Element.LayoutTransform = Transform.Identity; surface.Select(new(0, 0)); await Idle();
            surface.View.SetRegion(new(Guid.NewGuid(), sheet.Id, new(0, 0, 8, 3)));
            surface.View.Apply(snapshot, 0, new(First: new(FilterOperator.GreaterOrEqual, "1")));
            surface.View.FrozenRows = surface.View.FrozenColumns = 1;
            surface.Refresh(snapshot); await Idle(); drag(4, 2); await Idle();
            var filtered = Pixels(); grid.InvalidateAll(); await Idle();
            check(surface.Current == new CellPosition(11, 2) && surface.Control.SelectedCellCount == 15,
                "Dragging across frozen and filtered rows preserves original addresses and excludes hidden cells");
            check(filtered.SequenceEqual(Pixels()), "Frozen filtered drag repaint matches a full repaint");
            Field("_dragStartCell", new FastGridCellAddress(0, 0)); drag(4, 2); await Idle();
            check(OverlayContains(0, 0) && OverlayContains(0, 2) && OverlayContains(11, 0) && OverlayContains(11, 2) && !OverlayContains(12, 2),
                "Selection overlay covers all four frozen/filtered quadrants without covering the next row");
            Field("_dragStartCell", FastGridCellAddress.Empty);
            grid.UseSelectionOverlay = false; await Idle();
            check(ReadPixels(OverlayFrame()).All(b => b == 0), "Turning off the opt-in layer restores bitmap selection without a stale overlay");
            grid.UseSelectionOverlay = true; await Idle();
            await VerifySmoothFastGridAsync(surface, snapshot, host, check);
        }
        finally { host.Close(); }
    }

    private async Task VerifySmoothFastGridAsync(FastGridSurface surface, WorkbookSnapshot snapshot, Window host, Action<bool, string> check)
    {
        var grid = surface.Control;
        void Field(string name, object value) => typeof(FastGridControl).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(grid, value);
        var advance = typeof(FastGridControl).GetMethod("AdvanceDragAutoScroll", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Action<double>>(grid);
        async Task Idle() => await WaitForFastGridPaintAsync(grid);
        ImageSource Content() => ((Image)grid.FindName("image")).Source;
        byte[] Pixels()
        {
            var canvas = (WriteableBitmap)typeof(FastGridControl).GetField("_drawBuffer", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)!;
            var visual = new DrawingVisual(); RenderOptions.SetBitmapScalingMode(visual, BitmapScalingMode.NearestNeighbor);
            using (var dc = visual.RenderOpen()) dc.DrawImage(Content(), new Rect(0, 0, canvas.PixelWidth, canvas.PixelHeight));
            var frame = new RenderTargetBitmap(canvas.PixelWidth, canvas.PixelHeight, 96, 96, PixelFormats.Pbgra32); frame.Render(visual);
            var bytes = new byte[canvas.PixelWidth * canvas.PixelHeight * 4]; frame.CopyPixels(bytes, canvas.PixelWidth * 4, 0); return bytes;
        }
        surface.View = new(); surface.Bind(snapshot, 0, 100_000, 256); surface.Select(new(0, 0)); await Idle();
        Field("_mouseOverCell", FastGridCellAddress.Empty); grid.InvalidateAll(); await Idle();
        var bitmapPixels = Pixels();
        grid.UseSmoothScrolling = true; await Idle();
        var retainedPixels = Pixels();
        var canvas = (WriteableBitmap)typeof(FastGridControl).GetField("_drawBuffer", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)!;
        // WPF compositing of individual premultiplied images can round a channel
        // by one. Exclude only the last pixel row/column: the old rasterizer clamps
        // off-screen borders there, while retained content correctly clips them.
        var differences = Enumerable.Range(0, bitmapPixels.Length).Where(i => i / 4 % canvas.PixelWidth < canvas.PixelWidth - 1 &&
            i / 4 / canvas.PixelWidth < canvas.PixelHeight - 1 && Math.Abs(bitmapPixels[i] - retainedPixels[i]) > 1).ToArray();
        check(Content() is DrawingImage && Content().IsFrozen && differences.Length == 0,
            "Retained content preserves interior raster pixels within one compositing level; mismatched channels=" + differences.Length);
        Field("_dragStartCell", new FastGridCellAddress(0, 0));
        var original = Content(); var reads = surface.ReadCount;
        var endpoint = typeof(FastGridControl).GetMethod("GetDragEndpoint", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Func<Point, FastGridCellAddress>>(grid);
        var fastExit = endpoint(new Point(canvas.PixelWidth + 30, canvas.PixelHeight + 30));
        check(fastExit.Row == grid.VisibleRowCount - 1 && fastExit.Column == grid.VisibleColumnCount - 1,
            "A fast pointer jump outside selects the visible corner without waiting for a scroll callback");
        DragSelectionProbe(grid)(19, 19); await Idle();
        check(ReferenceEquals(original, Content()) && reads == surface.ReadCount, "Selection-only dragging reuses retained content without reading cells");
        var before = grid.GetModelCellBounds(8, 8)!.Value;
        Field("_mouseCursorPoint", null!); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
        advance(0.016); await Idle();
        var after = grid.GetModelCellBounds(8, 8)!.Value;
        check(before.Top > after.Top && before.Top - after.Top < before.Height && before.Left > after.Left && before.Left - after.Left < before.Width,
            "A 16 ms edge step moves both axes by pixels smaller than one cell");
        check(grid.CurrentCell.Row == grid.FirstVisibleRowScrollIndex + grid.VisibleRowCount - 1 &&
            grid.CurrentCell.Column == grid.FirstVisibleColumnScrollIndex + grid.VisibleColumnCount - 1,
            "Pixel auto-scroll selection includes the partially visible trailing cell on both axes");
        var smoothPixels = Pixels(); grid.InvalidateAll(); await Idle();
        check(smoothPixels.SequenceEqual(Pixels()), "Cached pixel-scroll content matches a fresh full redraw");
        reads = surface.ReadCount;
        for (var i = 0; i < 20; i++) { advance(0.016); await Idle(); }
        var addedReads = surface.ReadCount - reads;
        reads = surface.ReadCount; grid.InvalidateAll(); await Idle(); var fullReads = surface.ReadCount - reads;
        check(addedReads < fullReads * 3, "Twenty diagonal pixel steps read only newly exposed content instead of twenty viewports");
        var center = grid.GetModelCellBounds(8, 8)!.Value;
        var hit = grid.GetCellAddress(new Point((center.Left + center.Width / 2) * grid.RenderScaleX,
            (center.Top + center.Height / 2) * grid.RenderScaleY));
        check(hit.Row == 8 && hit.Column == 8, "Hit testing agrees with translated cell geometry after fractional scroll");
        Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
        var stopped = grid.GetModelCellBounds(8, 8)!.Value; advance(0.05); await Idle();
        check(grid.GetModelCellBounds(8, 8) == stopped, "Re-entering the viewport stops edge movement immediately");
        Field("_mouseIsBehindTop", true); Field("_mouseIsBehindLeft", true);
        for (var i = 0; i < 20; i++) advance(0.05);
        await Idle();
        check(grid.FirstVisibleRowScrollIndex == 0 && grid.FirstVisibleColumnScrollIndex == 0 && grid.GetModelCellBounds(8, 8) == before,
            "Reverse pixel scrolling clamps exactly at the top and left sheet boundaries");
        Field("_mouseIsBehindTop", false); Field("_mouseIsBehindLeft", false); Field("_dragStartCell", FastGridCellAddress.Empty);

        surface.View.FrozenRows = surface.View.FrozenColumns = 1; surface.Refresh(snapshot); await Idle();
        var frozen = grid.GetModelCellBounds(0, 0)!.Value;
        Field("_dragStartCell", new FastGridCellAddress(0, 0)); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
        advance(0.016); await Idle();
        check(grid.GetModelCellBounds(0, 0) == frozen, "Fractional scrolling keeps the frozen intersection stationary");
        var frozenPixels = Pixels(); grid.InvalidateAll(); await Idle();
        check(frozenPixels.SequenceEqual(Pixels()), "Frozen pixel scrolling has identical cached and fresh pixels");
        surface.View.SetRegion(new(Guid.NewGuid(), snapshot.Sheets[0].Id, new(0, 0, 8, 3)));
        surface.View.Apply(snapshot, 0, new(First: new(FilterOperator.GreaterOrEqual, "1")));
        surface.Refresh(snapshot); await Idle();
        Field("_dragStartCell", new FastGridCellAddress(0, 0)); advance(0.016); await Idle();
        check(grid.CurrentModelCell.Row == grid.CurrentCell.Row + 7 && grid.GetSelectedModelRanges().All(r => r.Bottom < 2 || r.Top > 8),
            "Pixel edge selection across frozen filtered rows retains model addresses and excludes hidden rows");
        var filteredPixels = Pixels(); grid.InvalidateAll(); await Idle();
        check(filteredPixels.SequenceEqual(Pixels()), "Filtered fractional scrolling has identical cached and fresh pixels");
        Field("_dragStartCell", FastGridCellAddress.Empty); Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
        host.Content = null; await Idle(); host.Content = surface.Element; await Idle();
        check(!grid.IsRenderPending && Content() is DrawingImage, "Retained content reloads without leaving pending work");
        surface.View = new(); surface.Bind(snapshot, 0, 100_000, 256); await Idle();
        grid.SetModelDimensions(new Dictionary<int, double> { [3] = 72 }, new Dictionary<int, double> { [3] = 180 }); await Idle();
        Field("_dragStartCell", new FastGridCellAddress(0, 0)); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
        advance(0.016); await Idle();
        var sized = grid.GetModelCellBounds(3, 3)!.Value;
        hit = grid.GetCellAddress(new Point((sized.Left + sized.Width / 2) * grid.RenderScaleX, (sized.Top + sized.Height / 2) * grid.RenderScaleY));
        check(Math.Abs(sized.Width - 180) < 1 && Math.Abs(sized.Height - 72) < 1 && hit.Row == 3 && hit.Column == 3,
            "Custom row heights and column widths remain aligned with pixel hit testing");
        var beforeStall = grid.GetModelCellBounds(3, 3)!.Value; advance(5); await Idle();
        check(beforeStall.Top - grid.GetModelCellBounds(3, 3)!.Value.Top <= 19,
            "A stalled dispatcher does not catch up with a multi-second scroll jump");
        var lostCapture = typeof(FastGridControl).GetMethod("OnLostMouseCapture", BindingFlags.Instance | BindingFlags.NonPublic)!
            .CreateDelegate<Action<System.Windows.Input.MouseEventArgs>>(grid);
        var timer = (DispatcherTimer)typeof(FastGridControl).GetField("_dragTimer", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)!;
        var startScrolling = typeof(FastGridControl).GetMethod("StartDragScrolling", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Action>(grid);
        startScrolling(); lostCapture(new(System.Windows.Input.Mouse.PrimaryDevice, Environment.TickCount)); await Idle();
        var released = grid.GetModelCellBounds(3, 3); advance(0.05); await Idle();
        check(!timer.IsEnabled && !(bool)typeof(FastGridControl).GetField("_dragRenderingSubscribed", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid)! &&
            !grid.IsSelectionDragging && grid.GetModelCellBounds(3, 3) == released,
            "Lost capture detaches the rendering callback and preserves the final viewport position");
        surface.Select(new(99_999, 255)); await Idle();
        Field("_dragStartCell", new FastGridCellAddress(99_999, 255)); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
        for (var i = 0; i < 30; i++) advance(0.05);
        await Idle(); var endFrame = Content(); advance(0.05); await Idle();
        var endCell = grid.GetModelCellBounds(99_999, 255)!.Value;
        check(ReferenceEquals(endFrame, Content()) && endCell.Bottom <= canvas.PixelHeight / grid.RenderScaleY + 1 &&
            endCell.Right <= canvas.PixelWidth / grid.RenderScaleX + 1,
            "Pixel scrolling clamps at the bottom and right without blank overscroll or redundant frames");
        Field("_dragStartCell", FastGridCellAddress.Empty); Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
        surface.Bind(snapshot, 0, 2, 4); await Idle();
        Field("_dragStartCell", new FastGridCellAddress(0, 0)); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
        var shortFrame = Content(); advance(0.05); await Idle();
        check(ReferenceEquals(shortFrame, Content()) && grid.FirstVisibleRowScrollIndex == 0 && grid.FirstVisibleColumnScrollIndex == 0,
            "A sheet smaller than the viewport never starts edge scrolling into blank space");
        Field("_dragStartCell", FastGridCellAddress.Empty); Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
    }
}
