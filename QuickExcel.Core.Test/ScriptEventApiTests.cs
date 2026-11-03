using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptEventApiTests
{
    private static readonly (string Constant, string Name, string Property)[] Events = [
        ("ON_OPEN", "workbook.open", "onWorkbookOpen"), ("ON_BEFORE_CLOSE", "workbook.beforeClose", "onWorkbookBeforeClose"),
        ("ON_CLOSE", "workbook.close", "onWorkbookClose"), ("ON_ACTIVATE", "workbook.activate", "onWorkbookActivate"),
        ("ON_DEACTIVATE", "workbook.deactivate", "onWorkbookDeactivate"), ("ON_BEFORE_SAVE", "workbook.beforeSave", "onWorkbookBeforeSave"),
        ("ON_SAVED", "workbook.saved", "onWorkbookSaved"), ("ON_CALCULATED", "workbook.calculated", "onWorkbookCalculated"),
        ("ON_WORKSHEET_ADDED", "sheet.added", "onWorksheetAdded"), ("ON_WORKSHEET_DELETED", "sheet.deleted", "onWorksheetDeleted"),
        ("ON_WORKSHEET_RENAMED", "sheet.renamed", "onWorksheetRenamed"), ("ON_WORKSHEET_ACTIVATE", "sheet.activate", "onWorksheetActivate"),
        ("ON_WORKSHEET_DEACTIVATE", "sheet.deactivate", "onWorksheetDeactivate"), ("ON_SELECTION_CHANGED", "selection.changed", "onSelectionChanged"),
        ("ON_CELL_CLICK", "cell.click", "onCellClick"), ("ON_CELL_EDIT", "cell.edit", "onCellEdit"),
        ("ON_CELL_COMMIT", "cell.commit", "onCellCommit"), ("ON_CELLS_CHANGED", "cells.changed", "onCellsChanged")
    ];
    private static ScriptModuleSet Modules(string source, ScriptLanguage language = ScriptLanguage.JavaScript) => new(
        new Dictionary<string, ScriptModule> { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js", source, language) },
        ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);

    [Fact]
    public void AllExistingEventsHaveConstantsAndPropertyHandlersWithCorrectArguments()
    {
        Assert.True(ScriptEventNames.All.SetEquals(Events.Select(e => e.Name)));
        var entries = JsonSerializer.Serialize(Events.Select(e => new[] { e.Constant, e.Name, e.Property }));
        var source = "const entries=" + entries + ";" + """
            if (!Object.isFrozen(QxEventNames) || Object.keys(QxEventNames).length !== entries.length) throw Error('constants');
            for (const [key, name, property] of entries) {
              if (QxEventNames[key] !== name || $.events[property] !== null) throw Error('mapping');
              $.events[property] = (argument, context) => {
                if (!Object.isFrozen(argument) || context.event.name !== name) throw Error('payload');
                if (name.startsWith('sheet.')) {
                  if (argument.id !== context.event.sheetId || argument.name !== 'Deleted or current') throw Error('sheet');
                } else if (argument !== context.event) throw Error('event');
                console.log('property:' + name);
              };
              $.events.on(QxEventNames[key], context => console.log('listener:' + context.event.name));
            }
            """;
        var sheet = Sheet("Data"); var snapshot = Book(sheet).WithScriptModules(Modules(source)); var executor = new ScriptWorkerActionExecutor();
        foreach (var (_, name, _) in Events)
        {
            var stamp = new CalculationStamp(Guid.NewGuid(), snapshot.Revision, 1, 1);
            var ev = new ScriptEvent(name, stamp.SessionId, 1, snapshot.Revision, sheet.Id.ToString(), "Deleted or current", []);
            var result = executor.Run(new(stamp, Guid.NewGuid(), ScriptEventNames.Prefix + name, snapshot, default) { Event = ev }, default);
            Assert.Equal(["property:" + name, "listener:" + name], result.Logs.ToArray());
        }
    }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)]
    [InlineData(ScriptLanguage.TypeScript)]
    public void ReplacementNullRemovalAndExistingListenersShareOneOrderedTransaction(ScriptLanguage language)
    {
        var source = """
            $.events.onWorksheetActivate = () => { throw Error('replaced'); };
            $.events.on(QxEventNames.ON_WORKSHEET_ACTIVATE, ({patch,snapshot}) => patch.setValue(snapshot.sheets[0].id,'B1','listener'));
            $.events.onWorksheetActivate = (sheet, {patch}) => patch.setValue(sheet.id,'A1',sheet.name);
            $.events.onCellClick = () => { throw Error('removed'); };
            $.events.onCellClick = null;
            $.events.on(QxEventNames.ON_CELL_CLICK, ({patch,snapshot}) => patch.setValue(snapshot.sheets[0].id,'C1','retained'));
            $.events.onBeforeSave = () => { throw Error('alias replaced'); };
            $.events.onWorkbookBeforeSave = (event, {fs}) => fs.writeText('//xlsq/saved.txt',event.name);
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(Modules(source, language)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); Assert.Empty(doc.ListActions());
        doc.DispatchScriptEvents();
        Assert.Equal("Data", doc.GetCellValue(0, 0, 0)); Assert.Equal("listener", doc.GetCellValue(0, 0, 1));
        doc.Session.RaiseScriptEvent("cell.click", doc.Snapshot.ActiveSheetId); doc.DispatchScriptEvents();
        Assert.Equal("retained", doc.GetCellValue(0, 0, 2));
        doc.Session.RaiseScriptEvent("workbook.beforeSave"); doc.DispatchScriptEvents();
        Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/saved.txt"));
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.PackageFiles.Entries);
    }

    [Theory]
    [InlineData("$.events.onCellClick = 42;")]
    [InlineData("$.events.onCellClick = undefined;")]
    [InlineData("for (let i=0;i<32;i++) $.events.on(QxEventNames.ON_CELL_CLICK,()=>{}); $.events.onCellClick=()=>{};")]
    public void InvalidPropertyAssignmentsAndCombinedHandlerLimitAreRejected(string source)
    {
        var snapshot = Book(Sheet("Data")).WithScriptModules(Modules(source));
        Assert.Throws<ScriptRuntimeFailureException>(() => new ScriptWorkerActionExecutor().List(
            new(new(Guid.NewGuid(), snapshot.Revision, 1, 1), Guid.NewGuid(), null, snapshot, default), default));
    }

    [Fact]
    public void RemovingOnlyPropertyHandlerRemovesDiscoveryAndLateAssignmentsFail()
    {
        var modules = Modules("$.events.onCellClick=()=>{}; $.events.onCellClick=null; $.actions.register('late',()=>{$.events.onCellClick=()=>{};});");
        var snapshot = Book(Sheet("Data")).WithScriptModules(modules); var executor = new ScriptWorkerActionExecutor();
        var request = new ScriptActionRequest(new(Guid.NewGuid(), snapshot.Revision, 1, 1), Guid.NewGuid(), null, snapshot, default);
        Assert.Equal("late", Assert.Single(executor.List(request, default)).Name);
        Assert.Throws<ScriptRuntimeFailureException>(() => executor.Run(request with { Name = "late" }, default));
    }
}
