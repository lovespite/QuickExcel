using System.Text;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckEventFilesAsync(string directory, Action<bool, string> check)
    {
        _eventTimer.Stop(); _calculationTimer.Stop();
        _workspace.New(); ShowBook();
        var source = """
            const names=['workbook.open','workbook.activate','workbook.deactivate','sheet.added','sheet.deleted','sheet.activate','sheet.deactivate','cell.click','cell.edit','cell.commit','selection.changed','workbook.saved'];
            for(const name of names) quickExcel.events.on(name,({fs,event})=>{
                const entries=fs.exists('//mem/events')?fs.readJson('//mem/events'):[];
                entries.push(event.name); fs.writeJson('//mem/events',entries);
            });
            quickExcel.events.on('cell.commit',({snapshot,patch})=>patch.setValue(snapshot.sheets[0].id,'B1','event result'));
            quickExcel.events.on('workbook.beforeSave',({fs})=>fs.writeText('//xlsq/data/saved.txt','保存事件'));
            quickExcel.events.on('workbook.beforeClose',({fs})=>fs.writeText('//xlsq/data/closing.txt','关闭前'));
            quickExcel.events.on('workbook.close',({event})=>console.log(event.name));
            """;
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) },
            ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")])));
        await ApplyScriptsAsync();
        check(Document.ScriptsEnabled && _availableScriptActions.Count == 0, "Event registrations load with script consent and stay out of the action picker");
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var canvas = _grid.GetVisualDescendants().OfType<Canvas>().Single();
        var pointer = new Pointer(81, PointerType.Mouse, true);
        canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, canvas, new Point(82, 44), 0,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
        _grid.CancelInputGesture();
        _grid.Select(new(0, 0)); check(_grid.BeginEdit(), "Grid begins an event-observed edit");
        _grid.CellEditor.Text = "event input"; check(await CommitEditorsAsync(), "Grid commit succeeds with events queued");
        await DispatchEventsAsync();
        check(Document.GetCellValue(0, 0, 1) == "event result", "Cell commit handler writes through one validated worker transaction");
        await AddSheetAsync(); var added = Document.Snapshot.ActiveSheetId;
        _workspace.Editor.Execute(new DeleteSheetCommand(added)); Synchronize();
        ShowHome(); ShowBook(); await DispatchEventsAsync();
        var logs = JsonSerializer.Deserialize<string[]>(Encoding.UTF8.GetString(Document.Session.TransientFiles.Entries["//mem/events"].Data.AsSpan()))!;
        check(new[] { "workbook.open", "workbook.activate", "workbook.deactivate", "sheet.added", "sheet.deleted", "sheet.activate", "sheet.deactivate", "cell.click", "cell.edit", "cell.commit", "selection.changed" }.All(logs.Contains),
            "Avalonia pointer, editor, selection, sheet and workspace handlers deliver script events");
        var path = Path.Combine(directory, "events.xlsq"); _files = new SmokeFiles { NextSave = path };
        check(await SaveAsync(true), "Before-save event participates in checked save");
        check(new TypedWorkbookStorage().Open(path).PackageFiles.Entries.ContainsKey("//xlsq/data/saved.txt"), "Save persists files created by the before-save handler");
        _discardChoice = () => Task.FromResult(SaveDecision.Cancel);
        check(!await ConfirmDiscardAsync() && Document.IsDirty && Document.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/data/closing.txt"),
            "Before-close edits are included in the unsaved prompt and cancellation keeps the session");
        var retiring = Document;
        _workspace.New();
        check(retiring.Session.IsClosed && retiring.ScriptErrors.Contains("workbook.close"), "Accepted replacement delivers read-only close event before session cleanup");
        _discardChoice = null;
        var rewrite = "quickExcel.actions.register('更新源码',({fs})=>fs.writeText('//xlsq/scripts/main.js','// 文件 API 更新后的源码'));";
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", rewrite) },
            ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")])));
        await OpenScriptsAsync(); await ApplyScriptsAsync();
        await RunScriptActionAsync(_availableScriptActions.Single());
        check(!Document.ScriptsEnabled && _scriptPage!.Drafts.Models["main"].Committed == Document.ScriptSource && await CommitScriptsAsync(),
            "Virtual source edits revoke trust and refresh the open editor without restoring stale source");
        await CloseScriptWindowAsync(); _workspace.New();
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.ts", """
            let calls=0;
            function main(workbook: ExcelScript.Workbook) {
              const cell=$.getActiveCell(), range=workbook.getActiveRange();
              if(cell.getAddress()!=='B2' || range.getAddress()!=='B2:C3') throw Error('live selection');
              cell.getFormat().getFill().setColor('#ff0000');
              range.getFormat().getFill().setColor('#00ff00');
              workbook.getActiveWorksheet().getCell('A1').setValue(++calls);
            }
            """, ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")])));
        await ApplyScriptsAsync();
        _grid.SelectRange(new(1, 1, 2, 2), new(1, 1));
        check(_availableScriptActions.Single().Name == "main", "Typed main(workbook) is discovered as a runnable entry");
        await RunScriptActionAsync(_availableScriptActions.Single());
        var book = Document.Snapshot; var sheet = book.GetSheet(book.ActiveSheetId);
        check(book.Styles[sheet.GetStyleId(1, 1)].FillColor == "FF00FF00" && book.Styles[sheet.GetStyleId(2, 2)].FillColor == "FF00FF00" &&
            Document.GetCellValue(0, 0, 0) == "1", "Live objects use Avalonia's current selection and commit overlapping fill and value writes");
        await RunScriptActionAsync(_availableScriptActions.Single());
        check(Document.GetCellValue(0, 0, 0) == (OperatingSystem.IsWindows() ? "2" : "1"), "Windows workbook module state survives consecutive main executions");
        check(Document.Undo() && Document.GetCellValue(0, 0, 0) == (OperatingSystem.IsWindows() ? "1" : ""), "One undo restores the complete preceding main transaction");
        _workspace.New();
    }
}
