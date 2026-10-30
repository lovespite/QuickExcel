using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunHorizontalScrollStressAsync(string folder, string source)
    {
        Directory.CreateDirectory(folder);
        var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(source)));
        _suppressAutoCalculation = true;
        var results = new List<object>();
        using var process = Process.GetCurrentProcess();
        try
        {
            InstallDocument(WorkbookDocument.Open(source, _services), null, Path.GetFileName(source));
            for (var sheet = 0; sheet < _document.Snapshot.Sheets.Length; sheet++)
            {
                SheetPicker.SelectedIndex = sheet;
                await WaitForFastGridPaintAsync(_grid.Control);
                var bar = (ScrollBar)_grid.Control.FindName("hscroll");
                var widestColumn = _document.Snapshot.Sheets[sheet].Columns
                    .Where(p => !p.Value.Hidden && p.Value.Width.HasValue && p.Key < _grid.ColumnCount)
                    .OrderByDescending(p => p.Value.Width).Select(p => p.Key).DefaultIfEmpty(0).First();
                foreach (var scenario in new[] { (Name: "full", Interval: 16), (Name: "full", Interval: 2), (Name: "wide-column", Interval: 8) })
                {
                    var interval = scenario.Interval;
                    var maximum = bar.Maximum;
                    var start = 0d; var end = maximum;
                    if (scenario.Name == "wide-column")
                    {
                        _grid.Select(new(0, widestColumn));
                        await WaitForFastGridPaintAsync(_grid.Control);
                        start = bar.Value;
                        var width = (_grid.Control.GetModelCellBounds(0, widestColumn)?.Width ?? 32) * VisualTreeHelper.GetDpi(_grid.Control).DpiScaleX;
                        end = Math.Min(maximum, start + Math.Max(0, width - 32));
                    }
                    var inputMs = new List<double>(); var delayMs = new List<double>();
                    var frames = new List<double>();
                    var positions = new List<double?>();
                    var operations = new ConcurrentQueue<DispatcherOperation>();
                    var allocated = GC.GetTotalAllocatedBytes(); var drawings = _grid.Artwork.CellDrawingCount;
                    var timer = Stopwatch.StartNew();
                    var final = 0d;
                    var content = (System.Windows.Controls.Image)_grid.Control.FindName("image");
                    var lastSource = content.Source;
                    void Rendering(object? sender, EventArgs e)
                    {
                        if (ReferenceEquals(lastSource, content.Source)) return;
                        lastSource = content.Source;
                        frames.Add(timer.Elapsed.TotalMilliseconds);
                        positions.Add(_grid.Control.GetModelCellBounds(0, widestColumn)?.Left);
                    }
                    CompositionTarget.Rendering += Rendering;
                    try
                    {
                        // Produce independently of painting. Waiting for idle after
                        // every request hides a UI thread that falls behind a real drag.
                        await Task.Run(async () =>
                        {
                            for (var i = 0; i <= 240; i++)
                            {
                                var pass = i / 60; var step = i % 60;
                                var value = start + (end - start) * (pass % 2 == 0 ? step / 60d : 1 - step / 60d);
                                if (i == 240) value = end;
                                var posted = timer.Elapsed.TotalMilliseconds;
                                operations.Enqueue(Dispatcher.InvokeAsync(() =>
                                {
                                    delayMs.Add(timer.Elapsed.TotalMilliseconds - posted);
                                    var input = Stopwatch.StartNew();
                                    bar.Value = value;
                                    bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.ThumbTrack, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
                                    inputMs.Add(input.Elapsed.TotalMilliseconds);
                                    final = value;
                                }, DispatcherPriority.Input));
                                await Task.Delay(interval);
                            }
                        });
                        foreach (var operation in operations) await operation;
                        bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.EndScroll, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
                        await WaitForFastGridPaintAsync(_grid.Control);
                    }
                    finally { CompositionTarget.Rendering -= Rendering; }
                    var completion = timer.Elapsed.TotalMilliseconds;
                    var idleProbe = Stopwatch.StartNew();
                    await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Input);
                    var endResponse = idleProbe.Elapsed.TotalMilliseconds;
                    process.Refresh();
                    var frameIntervals = frames.Zip(frames.Skip(1), (a, b) => b - a).ToArray();
                    results.Add(new { sheet, scenario = scenario.Name, intervalMs = interval, completionMs = completion, endResponseMs = endResponse,
                        lastValue = bar.Value, expectedLastValue = end, finalRequest = final,
                        firstColumn = _grid.Control.FirstVisibleColumnScrollIndex,
                        inputMaxMs = inputMs.Max(), queueP95Ms = delayMs.Order().ElementAt((int)(delayMs.Count * .95)), queueMaxMs = delayMs.Max(),
                        allocatedBytes = GC.GetTotalAllocatedBytes() - allocated, privateBytes = process.PrivateMemorySize64,
                        newStyledDrawings = _grid.Artwork.CellDrawingCount - drawings, cachedDrawings = _grid.Artwork.CachedCellDrawingCount,
                        frames = frames.Count, frameGapP95Ms = frameIntervals.Length == 0 ? 0 : frameIntervals.Order().ElementAt((int)(frameIntervals.Length * .95)),
                        frameGapMaxMs = frameIntervals.DefaultIfEmpty().Max(), frameTimesMs = frames, frameIntervalsMs = frameIntervals,
                        wideColumn = widestColumn, wideColumnPositionsDip = positions, inputMs, queueDelayMs = delayMs });
                    File.WriteAllText(Path.Combine(folder, "horizontal-scroll-stress.json"), JsonSerializer.Serialize(new
                    {
                        sourceHash, results,
                        processRenderMode = RenderOptions.ProcessRenderMode.ToString(), renderCapabilityTier = RenderCapability.Tier >> 16,
                        endpoint = "Independent full-width and widest-column input requests; frame gaps observe changed content at WPF Rendering, not physical display latency. Includes cold first traversal, no samples discarded."
                    }, new JsonSerializerOptions { WriteIndented = true }));
                    if (bar.Value != end || final != end) throw new Exception("Horizontal drag lost its final position.");
                }
            }
            if (sourceHash != Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(source))))
                throw new Exception("Scroll stress changed its source workbook.");
            File.WriteAllText(Path.Combine(folder, "horizontal-scroll-benchmark-result.txt"), "PASS: full-width source workbook scroll stress completed and source unchanged; inspect backlog and latency in JSON.");
        }
        finally { _suppressAutoCalculation = false; }
    }
}
