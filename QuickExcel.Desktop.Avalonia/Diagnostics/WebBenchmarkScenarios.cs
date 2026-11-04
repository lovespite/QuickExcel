using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task BenchmarkWebScenarioAsync(string directory, ProcessTreeMemory memory, List<object> results, Action<bool, string> check)
    {
        var mode = Program.WebBenchmarkKind;
        if (mode is "cold-editor" or "concurrent")
        {
            var boundary = Enumerable.Range(0, 64).ToDictionary(i => "m" + i, i => new ScriptModule($"scripts/m{i}.js",
                (i == 0 ? "host.functions.register('SMOKE', value => value);\n//" : $"export const value = {i};\n//").PadRight(16 * 1024, ' ')));
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(boundary, ["m0"], scriptApiVersion: 3)));
            var initialized = Stopwatch.StartNew(); await OpenScriptsAsync();
            results.Add(new { scenario = "first-editor-in-fresh-process-profile", elapsedMs = initialized.Elapsed.TotalMilliseconds, processId = Environment.ProcessId });
            check(initialized.Elapsed.TotalMilliseconds <= 5000, "Fresh-process Monaco with 64 modules and actual language worker ready <= 5 s");
            if (mode == "cold-editor") { await CloseScriptWindowAsync(); return; }
        }
        var book = mode == "concurrent" ? DenseImagePdfBook() : LightPdfBook(100);
        await using var session = await PrintSession.CreateAsync(book, book.ActiveSheetId, mode + ".xlsx", Path.Combine(directory, "printing"));
        var content = new ContentControl();
        var window = new Window { Title = "QuickExcel PDF " + mode, Width = 1120, Height = 820, Content = content };
        window.Show(this);
        PdfPreviewPage NewPage() => new(WebAssets, Document.Session.SessionId.ToString("N"), ++_webGeneration, _userDataDirectory);
        async Task OpenCloseAsync()
        {
            await using var page = NewPage(); content.Content = page.Browser.View; await page.LoadAsync(session);
            await page.DisposeAsync(); content.Content = null;
        }
        try
        {
            if (mode == "cold-pdf")
            {
                var initialized = Stopwatch.StartNew(); await using var page = NewPage(); content.Content = page.Browser.View;
                await page.Browser.Ready; var ready = initialized.Elapsed.TotalMilliseconds;
                var delivered = Stopwatch.StartNew(); await page.LoadAsync(session);
                results.Add(new { scenario = "first-pdf-in-fresh-process-profile", browserReadyMs = ready, firstPageMs = delivered.Elapsed.TotalMilliseconds, totalMs = initialized.Elapsed.TotalMilliseconds, processId = Environment.ProcessId });
                check(ready <= 5000 && delivered.Elapsed.TotalMilliseconds <= 5000, "Fresh-process PDF browser and 100-page PDF delivery each <= 5 s");
            }
            else if (mode == "retention")
            {
                await OpenCloseAsync(); await Task.Delay(30000); var baseline = memory.Samples.Last();
                for (var cycle = 0; cycle < 10; cycle++) await OpenCloseAsync();
                await Task.Delay(30000); var after = memory.Samples.Last();
                results.Add(new { scenario = "100-page-pdf-ten-cycles", baselineBytes = baseline.Bytes, finalBytes = after.Bytes, growthBytes = after.Bytes - baseline.Bytes, processes = after.Processes });
                check(after.Bytes - baseline.Bytes <= 64L * 1024 * 1024, "Ten 100-page PDF cycles + 30 s idle grow private commit by <= 64 MiB");
                check(after.Processes.Length == 1, "No PDF browser or worker process remains after ten closes");
            }
            else if (mode == "concurrent")
            {
                // Keep the native browser surfaces exposed while the grid gets keyboard focus.
                var bounds = Screens.Primary!.WorkingArea;
                window.Position = new(bounds.Right - (int)(window.Width * RenderScaling), bounds.Y);
                window.Topmost = true;
                _scriptWindow!.Position = new(bounds.Right - (int)(_scriptWindow.Width * RenderScaling), bounds.Bottom - (int)(_scriptWindow.Height * RenderScaling));
                _scriptWindow.Topmost = true;
                await using var page = NewPage(); content.Content = page.Browser.View; await page.LoadAsync(session);
                using var cancel = new CancellationTokenSource(); var completions = new List<double>(); var zooms = new List<double>();
                async Task ExerciseAsync()
                {
                    var index = 0;
                    while (!cancel.IsCancellationRequested)
                    {
                        var completed = _scriptPage!.Browser.RequestAsync("test", "testResult", new { action = "completionSample" });
                        var zoomed = page.Browser.RequestAsync("test", "testResult", new { action = "zoomSample", value = index++ % 2 == 0 ? "1" : "1.25" });
                        await Task.WhenAll(completed, zoomed);
                        completions.Add((await completed).Payload.GetProperty("elapsedMs").GetDouble());
                        zooms.Add((await zoomed).Payload.GetProperty("elapsedMs").GetDouble());
                        await Task.Delay(100);
                    }
                }
                var work = ExerciseAsync(); bool gridPassed;
                try { gridPassed = await RunGridBenchmarkAsync(Path.Combine(directory, "grid"), shutdown: false); }
                finally { cancel.Cancel(); await work; }
                static double P95(List<double> values) { var sorted = values.Order().ToArray(); return sorted[(int)Math.Ceiling(sorted.Length * .95) - 1]; }
                results.Add(new { scenario = "grid-with-monaco-completion-and-image-pdf-zoom", screen = bounds.ToString(), pdfPosition = window.Position.ToString(), editorPosition = _scriptWindow.Position.ToString(), completionRawMs = completions, zoomRawMs = zooms, completionP95 = P95(completions), zoomP95 = P95(zooms) });
                check(gridPassed, "Existing grid scrolling, drag and sparse loads meet unchanged budgets during real Monaco completions and PDF zoom");
                check(completions.Count >= 10 && P95(completions) <= 500 && P95(zooms) <= 250, "Concurrent language completion and PDF zoom meet their response budgets");
                await CloseScriptWindowAsync();
            }
        }
        finally { window.Close(); }
    }
}
