using System.Collections.Immutable;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFastGridBenchmarkAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        // Budgets are declared before sampling and also documented in the implementation record.
        var budgets = new { navigationP95Ms = 100, hugeSelectionP95Ms = 50, dragStepP95Ms = 50, filterP95Ms = 500 };
        File.WriteAllText(Path.Combine(folder, "budgets.json"), JsonSerializer.Serialize(budgets));
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < 100_000; row += 10)
        {
            cells[new(row, 0)] = new(CellContent.FromLiteral(CellValue.FromText("部门" + row % 7)));
            cells[new(row, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(row)));
            cells[new(row, 2)] = new(CellContent.FromLiteral(CellValue.FromText("中文长文本用于同一数据集的绘制与远距定位对照")));
        }
        var sheet = new WorksheetSnapshot(SheetId.New(), "Benchmark", cells);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var raw = new Dictionary<string, List<double>>(); var memory = new Dictionary<string, object>();
        FastGridSurface? measuredCandidate = null;
        var inputRaw = new Dictionary<string, List<double>>();
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            if (measuredCandidate is not null) await WaitForFastGridPaintAsync(measuredCandidate.Control);
        }
        async Task Sample(string name, Action<int> action, int count = 20)
        {
            var samples = new List<double>(); raw[name] = samples;
            var inputSamples = new List<double>(); inputRaw[name] = inputSamples;
            for (var i = -2; i < count; i++)
            {
                var timer = Stopwatch.StartNew(); action(Math.Max(0, i)); var inputMs = timer.Elapsed.TotalMilliseconds;
                await Idle(); timer.Stop();
                if (i >= 0) { samples.Add(timer.Elapsed.TotalMilliseconds); inputSamples.Add(inputMs); }
            }
        }
        var host = new Window { Owner = this, SizeToContent = SizeToContent.WidthAndHeight, ShowInTaskbar = false, Resources = Resources, Title = "Grid benchmark" };
        try
        {
            host.Show(); await Idle();
            long allocated;
            var timer = new Stopwatch();

            var candidate = new FastGridSurface((_, _) => false); candidate.Element.Width = 960; candidate.Element.Height = 500;
            measuredCandidate = candidate;
            host.Content = candidate.Element; await Idle(); allocated = GC.GetTotalAllocatedBytes(); timer.Restart();
            candidate.Bind(snapshot, 0, 100_000, 256); await Idle(); timer.Stop(); raw["FastWpfGrid.bind"] = [timer.Elapsed.TotalMilliseconds];
            await Sample("FastWpfGrid.navigate", i => candidate.Select(new(i * 4703 % 99_999, i * 13 % 256)));
            await Sample("FastWpfGrid.select20000", _ => candidate.SelectRange(new(0, 0, 999, 19), new(0, 0)), 10);
            await Sample("FastWpfGrid.select25600000", _ => candidate.SelectRange(new(0, 0, 99_999, 255), new(0, 0)));
            candidate.Select(new(0, 0)); await Idle();
            var drag = DragSelectionProbe(candidate.Control);
            await Sample("FastWpfGrid.dragSelectionStep", i => drag(4 + i % 8, 2 + i % 3));
            await Sample("FastWpfGrid.dragSameCell100Moves", _ => { for (var i = 0; i < 100; i++) drag(11, 3); });
            memory["FastWpfGrid"] = new { allocatedBytes = GC.GetTotalAllocatedBytes() - allocated, workingSet = Environment.WorkingSet,
                managedBytes = GC.GetTotalMemory(false), readCount = candidate.ReadCount, rectangles = candidate.Control.SelectedRanges.Count };
            var region = new DataRegion(Guid.NewGuid(), sheet.Id, new(0, 0, 99_999, 2));
            var filters = ImmutableDictionary<int, ColumnFilter>.Empty.Add(1, new(First: new(FilterOperator.GreaterOrEqual, "50000")));
            await Sample("Filter100000.evaluate", _ => DataFilterEngine.Evaluate(snapshot, region, filters));
            await Sample("Filter100000.candidates", _ => DataFilterEngine.Candidates(snapshot, region, filters, 0));
            candidate.View.SetRegion(region); candidate.View.Apply(snapshot, 1, filters[1]); candidate.Refresh(snapshot); await Idle();
            await Sample("FastWpfGrid.filteredNavigate", i => candidate.Select(new(50_000 + i * 10, 1)));
        }
        finally { host.Close(); }
        double P95(string key) => raw[key].Order().ElementAt((int)Math.Ceiling(raw[key].Count * .95) - 1);
        var results = raw.ToDictionary(pair => pair.Key, pair => new { p50Ms = pair.Value.Order().ElementAt(pair.Value.Count / 2), p95Ms = P95(pair.Key), samples = pair.Value.Count });
        var passed = P95("FastWpfGrid.navigate") <= budgets.navigationP95Ms &&
            P95("FastWpfGrid.select25600000") <= budgets.hugeSelectionP95Ms && P95("FastWpfGrid.dragSelectionStep") <= budgets.dragStepP95Ms &&
            P95("Filter100000.evaluate") <= budgets.filterP95Ms;
        File.WriteAllText(Path.Combine(folder, "fast-grid-benchmark.json"), JsonSerializer.Serialize(new { budgets, passed, results, raw, inputRaw, memory,
            gc = new[] { GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2) }, viewport = "960x500 DIP", warmup = 2,
            endpoint = "Dispatcher ApplicationIdle plus pending bitmap retry completion; inputRaw measures synchronous selection/hover handling only, not physical input-to-display latency", os = Environment.OSVersion.ToString(), runtime = Environment.Version.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "fast-grid-benchmark-result.txt"), $"PASS: benchmark completed; predefined performance budgets {(passed ? "met" : "NOT met")}. Native IME and cross-monitor DPI evidence remain separate.");
    }
}
