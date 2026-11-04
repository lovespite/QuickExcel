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
    private static readonly string[] Events = [
        "workbook.open", "workbook.beforeClose", "workbook.close", "workbook.activate",
        "workbook.deactivate", "workbook.beforeSave", "workbook.saved", "workbook.calculated",
        "sheet.added", "sheet.deleted", "sheet.renamed", "sheet.activate", "sheet.deactivate",
        "selection.changed", "cell.click", "cell.edit", "cell.commit", "cells.changed"
    ];
    private static ScriptModuleSet Modules(string source, ScriptLanguage language = ScriptLanguage.JavaScript) => new(
        new Dictionary<string, ScriptModule> { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js", source, language) },
        ["main"], scriptApiVersion: 3, bindings: [new("main", "actions")]);

    [Fact]
    public void AllExistingEventsHaveContextHandlersWithCorrectArguments()
    {
        Assert.True(ScriptEventNames.All.SetEquals(Events));
        var names = JsonSerializer.Serialize(Events);
        var source = "const names=" + names + ";" + """
            for (const name of names) host.events.on(name, context => {
              if (!Object.isFrozen(context) || !Object.isFrozen(context.event) || context.event.name !== name) throw Error('payload');
              if (name.startsWith('sheet.') && (context.event.sheetId === null || context.event.sheetName !== 'Deleted or current')) throw Error('sheet');
              console.log('listener:' + name);
            });
            """;
        var sheet = Sheet("Data"); var snapshot = Book(sheet).WithScriptModules(Modules(source)); var executor = new ScriptWorkerActionExecutor();
        foreach (var name in Events)
        {
            var stamp = new CalculationStamp(Guid.NewGuid(), snapshot.Revision, 1, 1);
            var ev = new ScriptEvent(name, stamp.SessionId, 1, snapshot.Revision, sheet.Id.ToString(), "Deleted or current", []);
            var result = executor.Run(new(stamp, Guid.NewGuid(), ScriptEventNames.Prefix + name, snapshot, default) { Event = ev }, default);
            Assert.Equal(["listener:" + name], result.Logs.ToArray());
        }
    }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)]
    [InlineData(ScriptLanguage.TypeScript)]
    public void MultipleHandlersShareOneOrderedTransaction(ScriptLanguage language)
    {
        var source = """
            host.events.on('sheet.activate', context => context.workbook.getActiveWorksheet().getRange('A1').setValue(context.event.sheetName));
            host.events.on('sheet.activate', context => context.workbook.getActiveWorksheet().getRange('B1').setValue('listener'));
            host.events.on('cell.click', context => context.workbook.getActiveWorksheet().getRange('C1').setValue('retained'));
            host.events.on('workbook.beforeSave', context => context.host.files.writeText('//xlsq/saved.txt',context.event.name));
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
    [InlineData("host.events.on('missing.event',()=>{});")]
    [InlineData("host.events.on('cell.click',42);")]
    [InlineData("for (let i=0;i<32;i++) host.events.on('cell.click',()=>{}); host.events.on('cell.click',()=>{});")]
    public void InvalidRegistrationsAndCombinedHandlerLimitAreRejected(string source)
    {
        var snapshot = Book(Sheet("Data")).WithScriptModules(Modules(source));
        Assert.Throws<ScriptRuntimeFailureException>(() => new ScriptWorkerActionExecutor().List(
            new(new(Guid.NewGuid(), snapshot.Revision, 1, 1), Guid.NewGuid(), null, snapshot, default), default));
    }

    [Fact]
    public void LateRegistrationInsideMainFailsWithoutCommitting()
    {
        var modules = Modules("host.events.on('cell.click',()=>{}); function main(){ host.events.on('cell.click',()=>{}); }");
        var snapshot = Book(Sheet("Data")).WithScriptModules(modules); var executor = new ScriptWorkerActionExecutor();
        var request = new ScriptActionRequest(new(Guid.NewGuid(), snapshot.Revision, 1, 1), Guid.NewGuid(), null, snapshot, default);
        Assert.Equal("main", Assert.Single(executor.List(request, default), a => !a.Name.StartsWith("@event/", StringComparison.Ordinal)).Name);
        Assert.Throws<ScriptRuntimeFailureException>(() => executor.Run(request with { Name = "main" }, default));
    }
}
