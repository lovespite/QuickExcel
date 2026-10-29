using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using FastWpfGrid;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFastGridDragBenchmarkAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        const double paintBudgetP95Ms = 50;
        const double scrollCadenceBudgetP95Ms = 33.34;
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < 300; row++)
            for (var column = 0; column < 40; column++)
                cells[new(row, column)] = new(CellContent.FromLiteral(CellValue.FromText($"数据 {row}:{column}")));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Drag", cells);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var surface = new FastGridSurface((_, _) => false);
        var grid = surface.Control;
        var status = new TextBlock();
        surface.SelectionChanged += () => status.Text = $"{surface.Current} · {grid.SelectedCellCount:N0}";
        var panel = new DockPanel(); DockPanel.SetDock(status, Dock.Bottom); panel.Children.Add(status); panel.Children.Add(surface.Element);
        var host = new Window { Owner = this, Width = 1400, Height = 850, ShowInTaskbar = false, Content = panel };
        var results = new Dictionary<string, object>();
        var raw = new Dictionary<string, List<double>>();
        using var process = Process.GetCurrentProcess();
        void Field(string name, object value) => typeof(FastGridControl).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(grid, value);
        var tick = typeof(FastGridControl).GetMethod("_dragTimer_Tick", BindingFlags.Instance | BindingFlags.NonPublic)!
            .CreateDelegate<Action<object, EventArgs>>(grid);
        var cellStep = typeof(FastGridControl).GetMethod("AdvanceLegacyDragAutoScroll", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Action>(grid);
        var startScrolling = typeof(FastGridControl).GetMethod("StartDragScrolling", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Action>(grid);
        var stopScrolling = typeof(FastGridControl).GetMethod("StopDragScrolling", BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<Action>(grid);
        var drag = DragSelectionProbe(grid);
        async Task Sample(string name, Action<int> step)
        {
            var input = new List<double>(); var paint = new List<double>();
            var memory = new List<double>();
            var reads = surface.ReadCount;
            var allocated = GC.GetTotalAllocatedBytes();
            var collections = new[] { GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2) };
            for (var i = -5; i < 60; i++)
            {
                // Give WPF's composition thread time to consume the preceding bitmap.
                // Back-to-back ApplicationIdle calls alone do not exercise a sustained drag.
                await Task.Delay(16);
                var timer = Stopwatch.StartNew(); step(i + 5); var inputMs = timer.Elapsed.TotalMilliseconds;
                await WaitForFastGridPaintAsync(grid);
                var paintMs = timer.Elapsed.TotalMilliseconds;
                process.Refresh();
                if (i >= 0) { input.Add(inputMs); paint.Add(paintMs); memory.Add(process.PrivateMemorySize64); }
            }
            raw[name + ".inputMs"] = input; raw[name + ".paintMs"] = paint;
            raw[name + ".privateBytes"] = memory;
            results[name] = new { inputP95Ms = input.Order().ElementAt(56), paintP50Ms = paint.Order().ElementAt(30),
                paintP95Ms = paint.Order().ElementAt(56), modelReads = surface.ReadCount - reads, selectedCells = grid.SelectedCellCount,
                sampledPeakPrivateBytes = memory.Max(), finalPrivateBytes = memory[^1], allocatedBytes = GC.GetTotalAllocatedBytes() - allocated,
                gc = Enumerable.Range(0, 3).Select(g => GC.CollectionCount(g) - collections[g]).ToArray() };
        }
        async Task TimedScroll(string name, bool smooth)
        {
            Field("_dragStartCell", FastGridCellAddress.Empty); Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
            grid.UseSmoothScrolling = smooth; surface.Select(new(0, 0)); await WaitForFastGridPaintAsync(grid);
            Field("_dragStartCell", new FastGridCellAddress(0, 0)); Field("_mouseIsBehindBottom", true); Field("_mouseIsBehindRight", true);
            Field("_mouseCursorPoint", null!); Field("_lastDragScrollTimestamp", 0L);
            var clock = Stopwatch.StartNew(); var times = new List<double>();
            void Changed(object? sender, EventArgs e) { if (clock.Elapsed.TotalMilliseconds >= 200) times.Add(clock.Elapsed.TotalMilliseconds); }
            grid.ViewportChanged += Changed;
            var reads = surface.ReadCount; var gen2 = GC.CollectionCount(2); var allocated = GC.GetTotalAllocatedBytes();
            try { startScrolling(); await Task.Delay(2200); }
            finally { stopScrolling(); grid.ViewportChanged -= Changed; }
            await WaitForFastGridPaintAsync(grid);
            var intervals = times.Zip(times.Skip(1), (a, b) => b - a).ToList();
            raw[name + ".intervalMs"] = intervals;
            results[name] = new { updates = times.Count, intervalP50Ms = intervals.Order().ElementAt(intervals.Count / 2),
                intervalP95Ms = intervals.Order().ElementAt((int)((intervals.Count - 1) * .95)),
                modelReads = surface.ReadCount - reads, allocatedBytes = GC.GetTotalAllocatedBytes() - allocated,
                gen2 = GC.CollectionCount(2) - gen2, firstRow = grid.FirstVisibleRowScrollIndex, firstColumn = grid.FirstVisibleColumnScrollIndex,
                offsetY = typeof(FastGridControl).GetField("_scrollOffsetY", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid),
                offsetX = typeof(FastGridControl).GetField("_scrollOffsetX", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(grid) };
        }
        try
        {
            host.Show(); surface.Bind(snapshot, 0, 100_000, 256); await WaitForFastGridPaintAsync(grid);
            Field("_dragStartCell", new FastGridCellAddress(0, 0));
            await Sample("viewportDrag", i => drag(2 + i % 18, 1 + i % 8));
            await Sample("drag10x10", i => drag(9 + i % 4, 9));
            await Sample("drag20x20", i => drag(19 + i % 4, 19));
            await Sample("drag40x40", i => drag(39 + i % 4, 39));
            await Sample("largeRangeDrag", i => drag(99_000 + i, 250 + i % 6));
            Field("_mouseIsBehindBottom", true); Field("_mouseMoveColumn", (int?)8);
            await Sample("autoScrollRows", _ => tick(grid, EventArgs.Empty));
            Field("_mouseIsBehindRight", true); Field("_mouseMoveRow", (int?)20);
            await Sample("autoScrollDiagonal", _ => tick(grid, EventArgs.Empty));
            // Same origin, addresses and whole-cell travel for both renderers.
            // This isolates cache cost from the new time-based scrolling speed.
            foreach (var smooth in new[] { false, true })
            {
                Field("_dragStartCell", FastGridCellAddress.Empty); Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
                grid.UseSmoothScrolling = smooth; surface.Select(new(0, 0)); await WaitForFastGridPaintAsync(grid);
                Field("_dragStartCell", new FastGridCellAddress(0, 0)); Field("_mouseIsBehindBottom", true);
                var prefix = smooth ? "retained" : "bitmap";
                await Sample(prefix + "CellStepRows", _ => cellStep());
                Field("_mouseIsBehindRight", true);
                await Sample(prefix + "CellStepDiagonal", _ => cellStep());
            }
            Field("_mouseIsBehindBottom", false); Field("_mouseIsBehindRight", false);
            Field("_dragStartCell", FastGridCellAddress.Empty);
            surface.Select(new(0, 0)); await WaitForFastGridPaintAsync(grid);
            Field("_dragStartCell", new FastGridCellAddress(0, 0));
            for (var repeat = 0; repeat < 3; repeat++)
                await Sample("sustainedDrag" + repeat, i => drag(2 + i % 18, 1 + i % 8));
            await TimedScroll("legacyAutoScrollTimer", false);
            await TimedScroll("pixelAutoScrollFrames", true);
        }
        finally { stopScrolling(); Field("_dragStartCell", FastGridCellAddress.Empty); host.Close(); }
        var passed = raw.Where(p => p.Key.EndsWith(".paintMs")).All(p => p.Value.Order().ElementAt(56) <= paintBudgetP95Ms);
        var cadence = raw["pixelAutoScrollFrames.intervalMs"].Order().ToArray();
        passed &= cadence.Length > 0 && cadence[(int)((cadence.Length - 1) * .95)] <= scrollCadenceBudgetP95Ms;
        File.WriteAllText(Path.Combine(folder, "fast-grid-drag-benchmark.json"), JsonSerializer.Serialize(new
        {
            paintBudgetP95Ms, scrollCadenceBudgetP95Ms, passed, results, raw, samples = 60, warmup = 5, viewport = "1400x850 DIP window",
            endpoint = "16 ms paced selection/hover and scroll handler to content/overlay repaint completion; cell-step controls have equal cell travel, pixel movement is time-based; two 2.2-second actual scheduler runs (legacy timer / coalesced WPF Rendering plus fallback timer) record update cadence, not physical presentation latency",
            os = Environment.OSVersion.ToString(), runtime = Environment.Version.ToString()
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "fast-grid-drag-benchmark-result.txt"), $"PASS: drag benchmark completed; performance budget {(passed ? "met" : "NOT met")}.");
    }
}
