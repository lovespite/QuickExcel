using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task BenchmarkPdfAsync(string directory, ProcessTreeMemory memory, List<object> results, Action<bool, string> check)
    {
        static double P95(IEnumerable<double> values) { var sorted = values.Order().ToArray(); return sorted[(int)Math.Ceiling(sorted.Length * .95) - 1]; }
        var visual = Program.WebBenchmarkKind == "pdf-visual";
        var root = new DirectoryInfo(AppContext.BaseDirectory); while (root is not null && !File.Exists(Path.Combine(root.FullName, "QuickExcel.slnx"))) root = root.Parent;
        if (root is null) throw new InvalidOperationException("Benchmark repository root is unavailable.");
        var source = Path.Combine(root.FullName, "artifacts/e4/excel-final-10/excel-source.xlsx");
        if (!File.Exists(source)) throw new FileNotFoundException("E4 Excel-generated reference fixture is required; see docs/57-e4-reporting-and-capacity.md.", source);
        using var e4 = new WorkbookDocument(_services.Storage.Open(source), _services); await Task.Run(() => e4.Calculate());
        var fixtures = new List<(string Name, WorkbookSnapshot Book, SheetId Sheet, int Pages)>();
        // Existing E4 suite is five pages in three PDFs: quickexcel=1, native=2, paged=2.
        var quickSource = Path.Combine(root.FullName, "artifacts/validation/m5-grid/20261102-130243-16cf45f2/Reporting/ui-report.xlsx");
        using var quickDocument = new WorkbookDocument(_services.Storage.Open(quickSource), _services); await Task.Run(() => quickDocument.Calculate());
        fixtures.Add(("e4-quickexcel", quickDocument.Snapshot, quickDocument.Snapshot.ActiveSheetId, 1));
        fixtures.Add(("e4-native", e4.Snapshot, e4.Snapshot.Sheets[0].Id, 2));
        fixtures.Add(("e4-paged", e4.Snapshot, e4.Snapshot.Sheets.Single(s => s.Name == "Paged").Id, 2));
        foreach (var count in new[] { 100, 2000 }) { var book = LightPdfBook(count); fixtures.Add(("light-" + count, book, book.ActiveSheetId, count)); }
        IEnumerable<(string Name, WorkbookSnapshot Book, SheetId Sheet, int Pages)> AllFixtures()
        {
            foreach (var fixture in fixtures) yield return fixture;
            // Do not allocate the high-density image fixture before measuring ordinary E4 documents.
            var images = DenseImagePdfBook(); yield return ("images-64MiB", images, images.ActiveSheetId, 1);
        }
        var window = new Window { Title = "Web PDF benchmark", Width = 1120, Height = 820, WindowStartupLocation = WindowStartupLocation.CenterOwner };
        var content = new ContentControl(); window.Content = content; window.Show(this);
        try
        {
            foreach (var fixture in AllFixtures())
            {
                if (visual && !fixture.Name.StartsWith("e4-")) break;
                var startedAt = memory.Samples.LastOrDefault()?.Milliseconds ?? 0; var generation = Stopwatch.StartNew();
                await using var session = await PrintSession.CreateAsync(fixture.Book, fixture.Sheet, fixture.Name + ".xlsx", Path.Combine(directory, "printing"));
                var generationMs = generation.Elapsed.TotalMilliseconds;
                if (session.Job.Layout.Pages.Length != fixture.Pages) throw new InvalidDataException($"{fixture.Name}: expected {fixture.Pages} pages, got {session.Job.Layout.Pages.Length}.");
                var copy = Path.Combine(directory, fixture.Name + ".pdf"); File.Copy(session.PdfPath, copy);
                string hash; using (var input = File.OpenRead(copy)) hash = Convert.ToHexString(await SHA256.HashDataAsync(input));
                var loads = new List<double>(); var cold = new List<double>(); var zooms = new List<double>();
                for (var run = 0; run < 5; run++)
                {
                    var created = Stopwatch.StartNew();
                    await using var page = new PdfPreviewPage(WebAssets, Document.Session.SessionId.ToString("N"), ++_webGeneration, Path.Combine(directory, "profiles", "normal-app-profile"));
                    content.Content = page.Browser.View; await page.Browser.Ready; cold.Add(created.Elapsed.TotalMilliseconds);
                    var delivered = Stopwatch.StartNew(); await page.LoadAsync(session); loads.Add(delivered.Elapsed.TotalMilliseconds);
                    if (visual)
                    {
                        for (var number = 1; number <= fixture.Pages; number++)
                        {
                            var data = new System.Text.StringBuilder(); var length = 1;
                            while (data.Length < length)
                            {
                                var snapshot = (await page.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "snapshot", ["page"] = number, ["offset"] = data.Length })).Payload;
                                length = snapshot.GetProperty("length").GetInt32(); if (length > 16 * 1024 * 1024) throw new InvalidDataException("Diagnostic PNG is too large.");
                                if (data.Length == 0) await File.WriteAllTextAsync(Path.Combine(directory, $"{fixture.Name}-{number}-canvas.json"), JsonSerializer.Serialize(new { width = snapshot.GetProperty("width").GetInt32(), height = snapshot.GetProperty("height").GetInt32(), scaleX = snapshot.GetProperty("scaleX").GetDouble(), scaleY = snapshot.GetProperty("scaleY").GetDouble() }));
                                data.Append(snapshot.GetProperty("data").GetString());
                            }
                            await File.WriteAllBytesAsync(Path.Combine(directory, $"{fixture.Name}-{number}.png"), Convert.FromBase64String(data.ToString()));
                        }
                        check(true, fixture.Name + " actual PDF.js page canvases captured for independent comparison"); break;
                    }
                    if (run == 4 && fixture.Name == "images-64MiB")
                        for (var index = 0; index < 65; index++)
                        {
                            var measured = await page.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "zoomSample", ["value"] = index % 2 == 0 ? "1" : "1.25" });
                            if (index >= 5) zooms.Add(measured.Payload.GetProperty("elapsedMs").GetDouble());
                        }
                    if (run == 4)
                    {
                        await page.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "page", ["page"] = fixture.Pages }); await Task.Delay(150);
                        var stats = (await page.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "inspect" })).Payload;
                        check(stats.GetProperty("canvases").GetInt32() <= 8 && stats.GetProperty("canvasBytes").GetInt64() <= 64 * 1024 * 1024 && stats.GetProperty("thumbnails").GetInt32() <= 12, fixture.Name + " canvas/thumbnail budgets");
                    }
                }
                var peak = memory.Samples.Where(s => s.Milliseconds >= startedAt).Max(s => s.Bytes);
                var budget = fixture.Name.StartsWith("e4-") ? 768L * 1024 * 1024 : 1536L * 1024 * 1024;
                results.Add(new { scenario = fixture.Name, fixture.Pages, fixtureSha256 = hash, generationMs, rawFirstPageMs = loads, firstPageP95 = P95(loads), rawBrowserReadyMs = cold, browserReadyP95 = P95(cold), rawZoomMs = zooms, zoomP95 = zooms.Count == 0 ? (double?)null : P95(zooms), peakBytes = peak, memoryBudgetBytes = budget });
                if (visual) continue;
                check(P95(loads) <= 5000, fixture.Name + " generated PDF delivery to first visible page p95 <= 5 s");
                check(P95(cold) <= 5000, fixture.Name + " browser creation to ready p95 <= 5 s (normal application profile; cold-process samples separate)");
                check(peak <= budget, fixture.Name + " full process-tree private commit budget");
                if (zooms.Count > 0) check(P95(zooms) <= 250, "High-image-density zoom p95 <= 250 ms");
                await File.WriteAllTextAsync(Path.Combine(directory, "pdf-progress.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
            }
        }
        finally { window.Close(); }
    }
    private static WorkbookSnapshot LightPdfBook(int pages)
    {
        var id = SheetId.New(); var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < pages; row++)
        { cells[new(row, 0)] = new(CellInput.Parse($"QuickExcel 中文轻量页 {row + 1}")); cells[new(row, 1)] = new(CellInput.Parse((row + 1).ToString())); }
        var sheet = new WorksheetSnapshot(id, "Light", cells, defaultRowHeight: 500,
            columns: [new(0, new(40)), new(1, new(15))], printing: new() { Setup = new() { PaperSize = 9, Scale = 100 }, Options = new() { GridLines = true } });
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], "");
    }
    private static WorkbookSnapshot DenseImagePdfBook()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Images", columns: Enumerable.Range(0, 8).Select(i => KeyValuePair.Create(i, new ColumnMetadata(9))), printing: new()
        { Setup = new() { PaperSize = 9, Scale = 100 }, Areas = [new(PrintRangeKind.Cells, 0, 0, 30, 7)] });
        var images = new List<SheetImage>();
        for (var index = 0; index < 4; index++)
        {
            using var bitmap = new SKBitmap(2048, 2048, SKColorType.Bgra8888, SKAlphaType.Premul);
            using (var canvas = new SKCanvas(bitmap))
            {
                canvas.Clear(new SKColor((byte)(50 + index * 45), 100, 160));
                using var paint = new SKPaint { Color = SKColors.Gold, IsAntialias = true };
                for (var stripe = 0; stripe < 128; stripe++) canvas.DrawCircle(stripe * 17, (stripe * 67) % 2048, 28, paint);
            }
            using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100);
            images.Add(new SheetImage(encoded.ToArray(), "png", 0, 0).WithAnchor(DrawingGeometry.FromBounds(sheet, new((index % 2) * 270, (index / 2) * 300, 250, 280), DrawingAnchorKind.OneCell)));
        }
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet.WithImages(images)], "");
    }
}
