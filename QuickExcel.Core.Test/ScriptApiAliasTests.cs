using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptApiAliasTests
{
    private static ScriptModuleSet Modules(string source, int api, ScriptLanguage language) => new(
        new Dictionary<string, ScriptModule> { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js", source, language) },
        ["main"], scriptApiVersion: api);
    private static IScriptRuntime Runtime(ScriptModuleSet modules)
    {
        try { return new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 1, 1, 1), TimeSpan.FromSeconds(5), default)); }
        catch (ScriptCompilationException error)
        { throw new InvalidOperationException(string.Join("\n", error.Diagnostics.Select(d => $"{d.Module}:{d.Line}:{d.Column} TS{d.Code}: {d.Message}")), error); }
    }

    [Theory]
    [InlineData(1, ScriptLanguage.JavaScript)]
    [InlineData(2, ScriptLanguage.JavaScript)]
    [InlineData(2, ScriptLanguage.TypeScript)]
    public void AliasAndOriginalShareOneFrozenApiAndFunctionRegistry(int api, ScriptLanguage language)
    {
        using var runtime = Runtime(Modules("""
            $.functions.register('SAME', () => $ === quickExcel && $.functions === quickExcel.functions
              && $.actions === quickExcel.actions && $.events === quickExcel.events);
            quickExcel.functions.register('ORIGINAL', n => n * 2);
            $.functions.register('FROZEN', () => {
              const d = Object.getOwnPropertyDescriptor(globalThis, '$');
              return !!d && !d.writable && !d.configurable && Object.isFrozen($) && Object.isFrozen($.functions);
            });
            """, api, language));
        Assert.Equal(true, runtime.Invoke("SAME", []));
        Assert.Equal(12d, runtime.Invoke("ORIGINAL", [6d]));
        Assert.Equal(true, runtime.Invoke("FROZEN", []));
    }

    [Theory]
    [InlineData(1, ScriptLanguage.JavaScript)]
    [InlineData(2, ScriptLanguage.JavaScript)]
    [InlineData(2, ScriptLanguage.TypeScript)]
    public void ExistingLocalDollarBindingsCanShadowTheAlias(int api, ScriptLanguage language)
    {
        using var runtime = Runtime(Modules("const $ = 42; quickExcel.functions.register('LOCAL', () => $);", api, language));
        Assert.Equal(42d, runtime.Invoke("LOCAL", []));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void RegistrationGuardsAreSharedAcrossBothNames(int api)
    {
        Assert.Throws<ScriptRuntimeFailureException>(() => Runtime(Modules(
            "$.functions.register('SAME',()=>1); quickExcel.functions.register('SAME',()=>2);", api, ScriptLanguage.JavaScript)));
        using var runtime = Runtime(Modules("$.functions.register('LATE',()=>{$.functions.register('OTHER',()=>1);return 0;});", api, ScriptLanguage.JavaScript));
        Assert.Throws<InvalidOperationException>(() => runtime.Invoke("LATE", []));
    }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)]
    [InlineData(ScriptLanguage.TypeScript)]
    public void AliasActionsAndEventsCommitAtomicPatchesAndKeepUndo(ScriptLanguage language)
    {
        var modules = Modules("""
            $.actions.register('write', ({snapshot,patch,fs}) => {
              if ($ !== quickExcel) throw Error('Different API');
              patch.setValue(snapshot.sheets[0].id, 'A1', 42);
              fs.writeText('//xlsq/value.txt', '42');
            });
            $.events.on('cell.click', ({snapshot,patch,fs,event}) => {
              if (event.name !== 'cell.click') throw Error('Different event');
              patch.setValue(snapshot.sheets[0].id, 'B1', fs.readText('//xlsq/value.txt'));
            });
            """, 2, language);
        modules = new(modules.Modules, modules.EntryModules, scriptApiVersion: 2, bindings: [new("main", "actions")]);
        using var document = new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        Assert.Throws<InvalidOperationException>(() => document.RunAction("write"));
        document.EnableScriptsAndCalculate();
        document.RunAction("write");
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.True(document.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/value.txt"));
        Assert.True(document.Undo());
        Assert.Empty(document.Snapshot.PackageFiles.Entries);
        Assert.Equal("", document.GetCellValue(0, 0, 0));
        Assert.True(document.Redo());
        document.Session.RaiseScriptEvent("cell.click", document.Snapshot.ActiveSheetId, ["A1"]);
        document.DispatchScriptEvents();
        Assert.Equal("42", document.GetCellValue(0, 0, 1));
        Assert.True(document.Undo());
        Assert.Equal("", document.GetCellValue(0, 0, 1));
        Assert.True(document.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/value.txt"));
    }
}
