using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunWebBenchmarkAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); var results = new List<object>(); Exception? failure = null;
        await using var memory = new ProcessTreeMemory();
        static double P95(IEnumerable<double> values) { var sorted = values.Order().ToArray(); return sorted[(int)Math.Ceiling(sorted.Length * .95) - 1]; }
        void Check(bool condition, string label) { checks.Add((condition ? "PASS " : "FAIL ") + label); File.AppendAllText(Path.Combine(directory, "progress.log"), checks[^1] + "\n"); }
        try
        {
            Width = 1400; Height = 850; ShowBook(); _workspace.New();
            if (Program.WebBenchmarkKind is "all" or "editor")
            {
            var boundary = Enumerable.Range(0, 64).ToDictionary(i => "m" + i, i => new ScriptModule($"scripts/m{i}.js",
                (i == 0 ? "host.functions.register('SMOKE', value => value);\n//" : $"export const value = {i};\n//").PadRight(16 * 1024, ' ')));
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(boundary, ["m0"], scriptApiVersion: 3)));
            var cold = new List<double>();
            for (var run = 0; run < 5; run++)
            {
                var started = Stopwatch.StartNew(); await OpenScriptsAsync(); cold.Add(started.Elapsed.TotalMilliseconds);
                await CloseScriptWindowAsync();
            }
            results.Add(new { scenario = "monaco-64-modules-1MiB-open", rawMs = cold, p95 = P95(cold), budgetMs = 5000 });
            Check(P95(cold) <= 5000, "Monaco 64-module initialization p95 <= 5 s (five opens; separate cold-process samples pending)");
            await OpenScriptsAsync(); var input = new List<double>(); var completion = new List<double>();
            for (var index = 0; index < 65; index++)
            {
                var changed = await _scriptPage!.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "inputSample", ["character"] = index % 2 == 0 ? "x" : "y" });
                var completed = await _scriptPage.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?> { ["action"] = "completionSample" });
                if (!completed.Payload.GetProperty("registered").GetBoolean()) throw new InvalidOperationException("Language completion returned no QuickExcel register API.");
                if (index >= 5) { input.Add(changed.Payload.GetProperty("elapsedMs").GetDouble()); completion.Add(completed.Payload.GetProperty("elapsedMs").GetDouble()); }
            }
            results.Add(new { scenario = "monaco-input-confirmed-paint", rawMs = input, p95 = P95(input), budgetMs = 50 });
            results.Add(new { scenario = "monaco-completion", rawMs = completion, p95 = P95(completion), budgetMs = 500 });
            Check(P95(input) <= 50, "64-module input + host acknowledgment + animation frame p95 <= 50 ms"); Check(P95(completion) <= 500, "64-module language completion p95 <= 500 ms");
            await CloseScriptWindowAsync();
            await Task.Delay(30000); var baseline = memory.Samples.Last().Bytes;
            for (var cycle = 0; cycle < 10; cycle++) { await OpenScriptsAsync(); await CloseScriptWindowAsync(); }
            await Task.Delay(30000); var after = memory.Samples.Last();
            results.Add(new { scenario = "monaco-ten-cycle-retention", baselineBytes = baseline, finalBytes = after.Bytes, growthBytes = after.Bytes - baseline, processes = after.Processes });
            Check(after.Bytes - baseline <= 64L * 1024 * 1024, "Ten editor cycles + 30 s idle grow private commit by <= 64 MiB");
            Check(after.Processes.Length == 1, "No browser or editor worker process remains after ten closes");
            Check(memory.Samples.Max(s => s.Bytes) <= 768L * 1024 * 1024, "Ordinary editor process tree peak private commit <= 768 MiB");
            }
            if (Program.WebBenchmarkKind is "all" or "pdf" or "pdf-visual") await BenchmarkPdfAsync(directory, memory, results, Check);
            if (Program.WebBenchmarkKind is "cold-editor" or "cold-pdf" or "retention" or "concurrent") await BenchmarkWebScenarioAsync(directory, memory, results, Check);
            if (Program.WebBenchmarkKind is not ("all" or "editor" or "pdf" or "pdf-visual" or "cold-editor" or "cold-pdf" or "retention" or "concurrent")) throw new ArgumentException("Unknown Web benchmark mode.");
        }
        catch (Exception ex) { failure = ex; }
        finally { try { await CloseScriptWindowAsync(); await ClosePrintPreviewAsync(); } catch (Exception ex) { failure ??= ex; } }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null && checks.All(c => c.StartsWith("PASS")), checks, results, error = failure?.ToString(), memory = memory.Samples,
            endpoint = "Synthetic Monaco edit to host checksum acknowledgment and requestAnimationFrame; whole Windows process tree private commit sampled at 100 ms; not physical presentation latency",
            scaling = RenderScaling, os = Environment.OSVersion.ToString(), webManifest = WebAssets.ManifestSha256 }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null && checks.All(c => c.StartsWith("PASS")) ? 0 : 1);
    }
}
