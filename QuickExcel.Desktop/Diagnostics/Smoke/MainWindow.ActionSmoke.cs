using System.IO;
using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunActionSmokeTestsAsync(string folder)
    {
        const string source = """
            quickExcel.actions.register({name:'整理商品',description:'批量去空格并写入公式'},({snapshot,patch})=>{
              const id=snapshot.sheets[0].id;
              patch.setValues(id,'B1:B2',snapshot.getRange(id,'A1:A2').map(row=>row.map(value=>value.trim())));
              patch.setFormula(id,'C1','=1+2');
              console.log('整理完成');
            });
            """;
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) },
            ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);
        var path = Path.Combine(folder, "worker-actions.xlsq");
        using (var seed = WorkbookDocument.CreateBlank(_services))
        {
            seed.SetCell(0, 0, 0, "  Apple  "); seed.SetCell(0, 1, 0, "  Pear  ");
            seed.Session.Execute(new UpdateScriptModulesCommand(modules)); seed.Save(path);
        }
        InstallDocument(WorkbookDocument.Open(path, _services), path, "worker-actions.xlsq", fileVersion: QuickExcel.Storage.FileVersions.Read(path));
        if (_document.ScriptsEnabled || RunActionButton.IsEnabled || ActionPicker.Items.Count != 0)
            throw new Exception("Opening an action package restored authorization or a live registry.");
        await RunScriptsAsync();
        ScriptPage_Click(this, new());
        if (ActionPicker.Items.Count != 1 || !RunActionButton.IsEnabled || !ActionsPanel.IsVisible)
            throw new Exception("The action selector is not ready after explicit application.");
        var before = _document.Snapshot.Revision;
        RunActionButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        // WorkAsync clears _busy before the button handler appends diagnostics. Await the whole handler operation.
        await _actionExecution.WaitAsync(TimeSpan.FromSeconds(6));
        if (_busy || _document.Snapshot.Revision != before + 1 || _document.GetCellValue(0, 0, 1) != "Apple" ||
            _document.GetCellValue(0, 1, 1) != "Pear" || _document.Snapshot.Sheets[0].GetCell(0, 2)?.Content.FormulaText != "=1+2")
            throw new Exception("Actual WPF action button did not commit the entire patch once.");
        if (!LogText.Text.Contains("整理完成")) throw new Exception("Action logs were not displayed.");
        _document.DisableScripts(); // Recorded undo/redo must work with execution disabled.
        await UndoAsync(false);
        if (_document.Session.CanUndo || _document.Snapshot.Sheets[0].GetCell(0, 1) is not null || _document.Snapshot.Sheets[0].GetCell(0, 2) is not null)
            throw new Exception("One undo did not remove every action edit.");
        await UndoAsync(true);
        if (_document.GetCellValue(0, 0, 1) != "Apple" || _document.GetCellValue(0, 0, 2) != "3" || _document.ScriptsEnabled)
            throw new Exception("Redo failed to restore recorded content with scripts disabled.");
        if (!await SaveAsync(false)) throw new Exception("Saving action package failed.");
        InstallDocument(WorkbookDocument.Open(path, _services), path, "worker-actions.xlsq", fileVersion: QuickExcel.Storage.FileVersions.Read(path));
        if (!_document.Snapshot.ScriptModules.HasActions || _document.ScriptsEnabled || _document.GetCellValue(0, 0, 1) != "Apple")
            throw new Exception("Action package roundtrip changed bindings, data or authorization.");

        ScriptEditor.Text = "quickExcel.actions.register('等待',({snapshot,patch})=>{patch.setValue(snapshot.sheets[0].id,'B1','wrong');while(true){}});";
        await RunScriptsAsync();
        var running = RunSelectedActionAsync();
        await Task.Delay(150);
        if (!StopCalculationButton.IsEnabled || !StopCalculationButton.IsVisible) throw new Exception("Action has no usable Stop control.");
        StopCalculationButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        await running.WaitAsync(TimeSpan.FromSeconds(3));
        if (_busy || _document.GetCellValue(0, 0, 1) != "Apple" || _document.Session.CanUndo)
            throw new Exception("Stopping an action left a partial patch or blocked the UI.");
        File.WriteAllText(Path.Combine(folder, "actions-smoke-result.txt"),
            "PASS: package open without authorization; explicit action discovery; actual WPF Run button; snapshot patch committed once; one undo/redo with scripts disabled; action logs; save/reopen; actual Stop button prevents partial writes.");
    }
}
