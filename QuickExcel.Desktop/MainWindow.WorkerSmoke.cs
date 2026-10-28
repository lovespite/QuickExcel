using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunWorkerSmokeTestsAsync(string folder)
    {
        var factory = new SmokeWorkerFactory();
        InstallDocument(WorkbookDocument.CreateBlank(new(_services.Storage, new WorkbookCalculator(factory))), null, "worker-smoke.xlsq");
        ScriptEditor.Text = "quickExcel.functions.register('VALUE',()=>42);";
        ApplyValues([(0, 0, "=VALUE()")]);
        await RunScriptsAsync();
        if (_document.GetCellValue(0, 0, 0) != "42") throw new Exception("Worker smoke initial calculation failed.");

        ScriptEditor.Text = "quickExcel.functions.register('VALUE',()=>{while(true){}});";
        factory.Ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        var running = RunScriptsAsync();
        var runtime = await factory.Ready.Task.WaitAsync(TimeSpan.FromSeconds(4));
        if (runtime.ProcessId == Environment.ProcessId || !StopCalculationButton.IsEnabled || !StopCalculationButton.IsVisible)
            throw new Exception("Worker or actual Stop button is not available during calculation.");
        var timer = Stopwatch.StartNew();
        StopCalculationButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        await running.WaitAsync(TimeSpan.FromSeconds(3));
        timer.Stop();
        if (_busy || _document.GetCellValue(0, 0, 0) != "42" || timer.Elapsed > TimeSpan.FromSeconds(2))
            throw new Exception("Stop did not preserve old results and restore editable UI.");

        factory.Ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        var crashing = RunScriptsAsync();
        var crashRuntime = await factory.Ready.Task.WaitAsync(TimeSpan.FromSeconds(4));
        using (var process = Process.GetProcessById(crashRuntime.ProcessId)) process.Kill();
        await crashing.WaitAsync(TimeSpan.FromSeconds(3));
        if (_busy || _document.GetCellValue(0, 0, 0) != "42" || _document.ScriptsEnabled)
            throw new Exception("Worker exit damaged UI, cache or authorization state.");
        ApplyValues([(1, 0, "still editable")]);
        _filePath = Path.Combine(folder, "worker-recovery.xlsq");
        if (!await SaveAsync(false)) throw new Exception("Saving after worker exit failed.");
        ScriptEditor.Text = "quickExcel.functions.register('VALUE',()=>84);";
        await RunScriptsAsync();
        if (_document.GetCellValue(0, 0, 0) != "84") throw new Exception("Explicit restart after worker exit failed.");
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", "import { value } from 'helper'; quickExcel.functions.register('VALUE',()=>value);"),
            ["helper"] = new("scripts/lib/helper.js", "export const value = 123;")
        }, ["main"], scriptApiVersion: 2);
        var modulePath = Path.Combine(folder, "worker-modules.xlsq");
        _services.Storage.Save(modulePath, _document.Snapshot.WithScriptModules(modules));
        InstallDocument(WorkbookDocument.Open(modulePath, _services), modulePath, "worker-modules.xlsq");
        if (_document.ScriptsEnabled || _document.GetCellValue(0, 0, 0) != "84")
            throw new Exception("Opening modules executed source or changed cached values.");
        await RunScriptsAsync();
        if (_document.GetCellValue(0, 0, 0) != "123") throw new Exception("Desktop package module calculation failed.");
        ScriptEditor.Text += "\n// primary entry edit";
        if (_document.ScriptsEnabled || _document.Snapshot.ScriptModules.ScriptApiVersion != 2)
            throw new Exception("Editing an entry retained authorization or lost module mode.");
        if (!await SaveAsync(false)) throw new Exception("Saving module package failed.");
        using var reopened = WorkbookDocument.Open(modulePath, _services);
        if (reopened.ScriptsEnabled || reopened.Snapshot.ScriptModules.Modules["helper"] != modules.Modules["helper"])
            throw new Exception("Module roundtrip lost helpers or restored authorization.");
        await RunActionSmokeTestsAsync(folder);
        File.WriteAllText(Path.Combine(folder, "worker-smoke-result.txt"),
            $"PASS: separate worker; actual WPF Stop button returns in {timer.ElapsedMilliseconds} ms; strong-kill preserves cache; editing/saving and explicit restart recover; API v2 package module open/apply/edit/save/reopen through desktop composition.");
    }

    private sealed class SmokeWorkerFactory : IScriptRuntimeFactory
    {
        public TaskCompletionSource<ScriptWorkerRuntime> Ready { get; set; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public IScriptRuntime Create(string source) => throw new InvalidOperationException("Stamped request required.");
        public IScriptRuntime Create(ScriptRuntimeRequest request)
        {
            var runtime = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create(request);
            Ready.TrySetResult(runtime); return runtime;
        }
    }
}
