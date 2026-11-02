using System.Diagnostics;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunGridBenchmarkAsync(string directory)
    {
        Directory.CreateDirectory(directory); Exception? error = null; var passed = true;
        var results = new Dictionary<string, object>(); var raw = new Dictionary<string, object>();
        using var grid = new WorkbookGrid();
        var host = new Window { Width = 1400, Height = 850, ShowInTaskbar = false, Content = grid, Title = "QuickExcel · Grid benchmark" };
        static double P95(IEnumerable<double> samples) { var ordered = samples.Order().ToArray(); return ordered.Length == 0 ? double.PositiveInfinity : ordered[(int)Math.Ceiling(ordered.Length * .95) - 1]; }
        async Task<Dictionary<string, double>> DrawAsync(Action action)
        {
            var done = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var start = Stopwatch.GetTimestamp(); var phases = new Dictionary<string, double>();
            void Trace(string stage, long timestamp)
            {
                phases[stage] = Stopwatch.GetElapsedTime(start, timestamp).TotalMilliseconds;
                if (phases.ContainsKey("content-submit") && phases.ContainsKey("selection-end")) done.TrySetResult();
            }
            grid.RenderingTrace += Trace;
            try { action(); phases["state-applied"] = Stopwatch.GetElapsedTime(start).TotalMilliseconds; await done.Task.WaitAsync(TimeSpan.FromSeconds(10)); return phases; }
            finally { grid.RenderingTrace -= Trace; }
        }
        try
        {
            host.Show(this); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            foreach (var styled in new[] { false, true })
            {
                var cells = new Dictionary<CellPosition, CellSnapshot>();
                for (var r = 0; r < 100; r++) for (var c = 0; c < 100; c++) cells[new(r, c)] = new(CellContent.FromLiteral(CellValue.FromText($"横向滚动 {r}:{c}"), styled ? 1 : 0));
                var sheet = new WorksheetSnapshot(SheetId.New(), "Scroll", cells);
                var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles:
                    [new(), new() { FontSize = 12, FillColor = "FFFFF2CC", FontColor = "FF234567", WrapText = true, Vertical = 1 }]);
                await DrawAsync(() => grid.Bind(snapshot, 0, 100_000, 256));
                var phases = new List<Dictionary<string, double>>(); var allocations = GC.GetTotalAllocatedBytes(); var drawings = grid.CellDrawingCount;
                for (var i = -5; i < 60; i++)
                {
                    await Task.Delay(16);
                    var sample = await DrawAsync(() => grid.SetBenchmarkHorizontalOffset((8 + ((i + 5) % 40) * .125) * 112));
                    if (i >= 0) phases.Add(sample);
                }
                var p95 = P95(phases.Select(p => Math.Max(p["content-submit"], p["selection-end"])));
                var name = styled ? "styled" : "plain"; passed &= p95 <= 16.67;
                results[name] = new { paintP95Ms = p95, withinBudget = p95 <= 16.67, newCellDrawings = grid.CellDrawingCount - drawings,
                    allocatedBytes = GC.GetTotalAllocatedBytes() - allocations, grid.CachedNativeBytes, privateBytes = Process.GetCurrentProcess().PrivateMemorySize64 };
                raw[name] = phases;
            }
            var frames = new List<long>(); void Frame(long stamp) => frames.Add(stamp);
            grid.DragFrameApplied += Frame;
            try { grid.BeginBenchmarkDrag(new(grid.Bounds.Width + 100, grid.Bounds.Height + 100)); await Task.Delay(2200); }
            finally { grid.CancelInputGesture(); grid.DragFrameApplied -= Frame; }
            var intervals = frames.Zip(frames.Skip(1), (a, b) => Stopwatch.GetElapsedTime(a, b).TotalMilliseconds).ToArray();
            var dragP95 = P95(intervals); passed &= frames.Count > 20 && dragP95 <= 33.34;
            results["drag"] = new { frames = frames.Count, intervalP95Ms = dragP95, withinBudget = dragP95 <= 33.34, offsetX = grid.ScrollOffset.X, offsetY = grid.ScrollOffset.Y };
            raw["dragIntervalsMs"] = intervals;
            foreach (var count in new[] { 25_000, 100_000, 250_000 })
            {
                var cells = new Dictionary<CellPosition, CellSnapshot>(count + 1);
                for (var i = 0; i < count; i++) cells[new(i / 100, i % 100)] = new(CellContent.FromLiteral(CellValue.FromNumber(i), i % 8192));
                cells[new(199_999, WorkbookLimits.MaxColumns - 1)] = new(CellContent.FromLiteral(CellValue.FromText("稀疏远端")));
                var styles = Enumerable.Range(0, 8192).Select(i => new StyleDefinition { FillColor = "FF" + (0xE00000 + i).ToString("X6") }).ToArray();
                var sheet = new WorksheetSnapshot(SheetId.New(), "Load", cells);
                var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles);
                var bind = await DrawAsync(() => grid.Bind(snapshot, 0, WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns));
                var drawings = grid.CellDrawingCount; var allocated = GC.GetTotalAllocatedBytes();
                var far = await DrawAsync(() => grid.Select(new(199_999, WorkbookLimits.MaxColumns - 1)));
                var bounded = grid.CellDrawingCount - drawings < 1000 && grid.CachedNativeBytes <= 96L * 1024 * 1024;
                passed &= bounded;
                results["load" + count] = new { storedCells = cells.Count, styles = styles.Length, bounded, farViewportDrawings = grid.CellDrawingCount - drawings,
                    allocatedFarViewportBytes = GC.GetTotalAllocatedBytes() - allocated, grid.CachedNativeBytes, privateBytes = Process.GetCurrentProcess().PrivateMemorySize64 };
                raw["load" + count] = new { bind, far };
            }
        }
        catch (Exception ex) { error = ex; passed = false; }
        finally { host.Close(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed, error = error?.ToString(), results, raw,
            paintBudgetP95Ms = 16.67, dragBudgetP95Ms = 33.34, samples = 60, warmup = 5, windowDip = "1400x850", scale = RenderScaling,
            endpoint = "16 ms paced native ScrollBar.ValueChanged to content bitmap copy + Avalonia DrawImage submission and selection CPU draw; not GPU completion or physical display latency",
            dragEndpoint = "synthetic drag position through production RequestAnimationFrame scheduler; not physical pointer input",
            runtime = Environment.Version.ToString(), os = Environment.OSVersion.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).Shutdown(passed ? 0 : 1);
    }
}
