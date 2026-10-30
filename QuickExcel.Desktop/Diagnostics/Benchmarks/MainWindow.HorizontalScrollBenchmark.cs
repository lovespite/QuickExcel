using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls.Primitives;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunHorizontalScrollBenchmarkAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        const double paintBudgetP95Ms = 16.67;
        var results = new Dictionary<string, object>();
        var raw = new Dictionary<string, object>();
        var surface = new FastGridSurface((_, _) => false);
        var grid = surface.Control;
        var host = new Window { Owner = this, Width = 1400, Height = 850, ShowInTaskbar = false, Content = surface.Element };
        try
        {
            host.Show();
            foreach (var styled in new[] { false, true })
            {
                var cells = new Dictionary<CellPosition, CellSnapshot>();
                for (var row = 0; row < 100; row++)
                for (var column = 0; column < 100; column++)
                    cells[new(row, column)] = new(CellContent.FromLiteral(CellValue.FromText($"横向滚动 {row}:{column}"), styled ? 1 : 0));
                var sheet = new WorksheetSnapshot(SheetId.New(), "Scroll", cells);
                var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles:
                    [new(), new() { FontSize = 12, FillColor = "FFFFF2CC", FontColor = "FF234567", WrapText = true, Vertical = 1 }]);
                surface.Bind(snapshot, 0, 100_000, 256);
                await WaitForFastGridPaintAsync(grid);
                foreach (var artwork in styled ? new[] { true, false } : new[] { true })
                {
                    surface.Artwork.Visibility = artwork ? Visibility.Visible : Visibility.Hidden;
                    var bar = (ScrollBar)grid.FindName("hscroll");
                    void Scroll(double position)
                    {
                        bar.Value = position;
                        bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.ThumbTrack, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
                    }
                    Scroll(0); await WaitForFastGridPaintAsync(grid);
                    var input = new List<double>(); var paint = new List<double>(); var travel = new List<double>();
                    var allocations = GC.GetTotalAllocatedBytes(); var reads = surface.ReadCount; var notifications = 0;
                    var cellDrawings = surface.Artwork.CellDrawingCount;
                    void Changed(object? sender, EventArgs e) => notifications++;
                    grid.ViewportChanged += Changed;
                    // Equal cell travel for item- and pixel-based scrollbar implementations.
                    var pixelBased = bar.ViewportSize == grid.GridScrollAreaWidth;
                    var unit = pixelBased ? 112 * System.Windows.Media.VisualTreeHelper.GetDpi(grid).DpiScaleX : 1;
                    try
                    {
                        for (var i = -5; i < 60; i++)
                        {
                            await Task.Delay(16);
                            var step = i + 5;
                            var timer = Stopwatch.StartNew();
                            Scroll((8 + (step % 40) * .125) * unit);
                            var inputMs = timer.Elapsed.TotalMilliseconds;
                            await WaitForFastGridPaintAsync(grid);
                            if (i >= 0)
                            {
                                input.Add(inputMs); paint.Add(timer.Elapsed.TotalMilliseconds);
                                travel.Add(grid.GetModelCellBounds(0, 14)?.X ?? double.NaN);
                            }
                        }
                    }
                    finally { grid.ViewportChanged -= Changed; }
                    // Hidden removes composition, but WPF can still invoke OnRender.
                    var name = styled ? artwork ? "styled" : "styledHiddenArtwork" : "plain";
                    var p95 = paint.Order().ElementAt((int)Math.Ceiling(paint.Count * .95) - 1);
                    results[name] = new { inputP95Ms = input.Order().ElementAt(56), paintP50Ms = paint.Order().ElementAt(30),
                        paintP95Ms = p95, withinBudget = p95 <= paintBudgetP95Ms, modelReads = surface.ReadCount - reads,
                        allocatedBytes = GC.GetTotalAllocatedBytes() - allocations, viewportNotifications = notifications, pixelBased,
                        newStyledDrawings = surface.Artwork.CellDrawingCount - cellDrawings, cachedStyledDrawings = surface.Artwork.CachedCellDrawingCount };
                    raw[name] = new { inputMs = input, paintMs = paint, column14PositionDip = travel };
                }
                surface.Artwork.Visibility = Visibility.Visible;
            }
        }
        finally { host.Close(); }
        File.WriteAllText(Path.Combine(folder, "horizontal-scroll-benchmark.json"), JsonSerializer.Serialize(new
        {
            paintBudgetP95Ms, results, raw, samples = 60, warmup = 5,
            endpoint = "16 ms paced real WPF ThumbTrack handler to content/artwork/selection paint completion; not physical display latency",
            processRenderMode = System.Windows.Media.RenderOptions.ProcessRenderMode.ToString(),
            renderCapabilityTier = System.Windows.Media.RenderCapability.Tier >> 16,
            viewport = "1400x850 DIP window", runtime = Environment.Version.ToString(), os = Environment.OSVersion.ToString()
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "horizontal-scroll-benchmark-result.txt"), "PASS: horizontal scroll benchmark completed; inspect per-scenario performance budgets in JSON.");
    }
}
