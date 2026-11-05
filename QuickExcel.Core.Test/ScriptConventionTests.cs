using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation.Web;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptConventionTests
{
    private static ScriptModuleSet Modules(params (string Id, string Role, string Source)[] modules) => new(
        modules.Select(m => KeyValuePair.Create(m.Id, new ScriptModule($"scripts/{m.Id}.js", m.Source))),
        modules.Where(m => m.Role != "common").Select(m => m.Id), scriptApiVersion: 4,
        bindings: modules.Where(m => m.Role != "common").Select(m => new ScriptBinding(m.Id, m.Role)));

    private static WorkbookDocument Document(ScriptModuleSet modules, WorksheetSnapshot? sheet = null) => new(
        Book(sheet ?? Sheet("Data")).WithScriptModules(modules),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));

    [Fact]
    public void MainRunsOncePerApplyAndActionsEventsShareStateWhileUdfIsIsolated()
    {
        using var document = Document(Modules(
            ("main", "initialization", "export function main(w) { globalThis.counter = 10; const c=w.getActiveWorksheet().getRange('D1'); c.setValue(Number(c.getValue())+1); }"),
            ("action", "actions", "export function main(w) { w.getActiveWorksheet().getRange('A1').setValue(++globalThis.counter); }"),
            ("events", "events", "export function onCellClick(c) { c.workbook.getActiveWorksheet().getRange('B1').setValue(++globalThis.counter); }"),
            ("udf", "formula-functions", "export function STATE() { return typeof globalThis.counter; }")),
            Sheet("Data", (0, 2, F("=JS.udf.STATE()"))));
        document.EnableScriptsAndCalculate();
        Assert.Equal("1", document.GetCellValue(0, 0, 3));
        Assert.Equal("undefined", document.GetCellValue(0, 0, 2));
        Assert.Equal("action", Assert.Single(document.ListActions()).Name);
        document.RunAction("action"); Assert.Equal("11", document.GetCellValue(0, 0, 0));
        document.Calculate(); document.ListActions();
        document.Session.RaiseScriptEvent("cell.click", document.Snapshot.ActiveSheetId);
        document.DispatchScriptEvents(); Assert.Equal("12", document.GetCellValue(0, 0, 1));
        Assert.Equal("1", document.GetCellValue(0, 0, 3));
        document.EnableScriptsAndCalculate();
        Assert.Equal("2", document.GetCellValue(0, 0, 3));
        document.RunAction("action"); Assert.Equal("11", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void InitializationFailureRollsBackCellsFilesAndAuthorization()
    {
        using var document = Document(Modules(("main", "initialization", "export function main(w,h) { w.getActiveCell().setValue(42); h.files.writeText('//xlsq/start.txt','bad'); throw Error('startup failed'); }")));
        Assert.ThrowsAny<Exception>(() => document.EnableScriptsAndCalculate());
        Assert.Equal("", document.GetCellValue(0, 0, 0));
        Assert.Empty(document.Snapshot.PackageFiles.Entries);
        Assert.False(document.ScriptsEnabled);
    }

    [Fact]
    public void FailedActionDoesNotReplayInitializationOrRecreateSharedState()
    {
        using var document = Document(Modules(
            ("main", "initialization", "export function main(w) { w.getActiveCell().setValue(7); }"),
            ("bad", "actions", "export function main() { throw Error('bad'); }")));
        document.EnableScriptsAndCalculate();
        Assert.ThrowsAny<Exception>(() => document.RunAction("bad")); Assert.False(document.ScriptsEnabled);
        Assert.ThrowsAny<Exception>(() => document.RunAction("bad"));
        Assert.Equal("7", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void PublicRegistrationIsAbsentAndUnboundEventExportsAreNotMounted()
    {
        using var document = Document(Modules(
            ("main", "initialization", "export function main(w,h) { if ('functions' in h || 'events' in h || 'actions' in h) throw Error('registration API'); }"),
            ("action", "actions", "export function main(w) { w.getActiveCell().setValue(1); } export function onCellClick(c) { throw Error('must not run'); }"),
            ("helper", "common", "export function onCellClick(c) { throw Error('must not run'); }"),
            ("events", "events", "import 'helper'; export function onCellClick(c) { c.workbook.getActiveCell().setValue(5); }")));
        document.EnableScriptsAndCalculate(); document.RunAction("action");
        document.Session.RaiseScriptEvent("cell.click", document.Snapshot.ActiveSheetId); document.DispatchScriptEvents();
        Assert.Equal("5", document.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData("export async function onCellClick(c) {}")]
    [InlineData("export function onCellClick() {}")]
    [InlineData("export function onCellClick(a,b) {}")]
    [InlineData("export function onCellClik(c) {}")]
    [InlineData("export const onCellClick = 1;")]
    public void InvalidEventSignaturesFailLoading(string source)
    {
        using var document = Document(Modules(("main", "initialization", ""), ("events", "events", source)));
        Assert.ThrowsAny<Exception>(() => document.EnableScriptsAndCalculate()); Assert.False(document.ScriptsEnabled);
    }

    [Fact]
    public void EventsWithSameNameExecuteInEntryOrderInOneUndoableTransaction()
    {
        using var document = Document(Modules(("first", "events", "export function onCellClick(c){ c.workbook.getActiveCell().setValue(4); }"),
            ("second", "events", "export function onCellClick(c){ const r=c.workbook.getActiveCell(); r.setValue(Number(r.getValue())+3); }")));
        document.EnableScriptsAndCalculate(); document.ListActions();
        document.Session.RaiseScriptEvent("cell.click", document.Snapshot.ActiveSheetId); document.DispatchScriptEvents();
        Assert.Equal("7", document.GetCellValue(0, 0, 0)); Assert.True(document.Undo()); Assert.Equal("", document.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void GeneratedTemplatesCompileAndRunWithoutRegistration(ScriptLanguage language)
    {
        var items = new[] { ("main", "initialization"), ("sample", "formula-functions"), ("action", "actions"), ("events", "events") };
        var modules = new ScriptModuleSet(items.Select(m => KeyValuePair.Create(m.Item1, new ScriptModule(
            $"scripts/{m.Item1}.{(language == ScriptLanguage.TypeScript ? "ts" : "js")}",
            m.Item2 == "events" ? string.Join("\n", ScriptEventCatalog.All.Select(e => e.Code(language))) :
                ScriptModuleTemplates.CreateConventional(m.Item1, language, m.Item2), language))),
            items.Select(m => m.Item1), scriptApiVersion: 4, bindings: items.Select(m => new ScriptBinding(m.Item1, m.Item2)));
        using var document = Document(modules, Sheet("Data", (0, 1, F("=JS.sample.DOUBLE(21)"))));
        document.EnableScriptsAndCalculate(); Assert.Equal("42", document.GetCellValue(0, 0, 1));
        document.RunAction("action"); Assert.Equal("42", document.GetCellValue(0, 0, 0));
        document.DispatchScriptEvents();
    }

    [Fact]
    public void EverySupportedEventHasSearchableBilingualTemplate()
    {
        Assert.True(ScriptEventNames.All.SetEquals(ScriptEventCatalog.All.Select(e => e.Name)));
        foreach (var item in ScriptEventCatalog.All)
        {
            Assert.NotEmpty(item.DescriptionZh); Assert.NotEmpty(item.DescriptionEn); Assert.NotEmpty(item.TimingZh); Assert.NotEmpty(item.TimingEn);
            Assert.True(item.Matches(item.Name)); Assert.True(item.Matches(item.DescriptionEn)); Assert.True(item.Matches(item.DescriptionZh));
            Assert.Contains("export function " + item.Handler, item.Code(ScriptLanguage.JavaScript));
        }
    }

    [Fact]
    public async Task ApiFourTypeCheckingDoesNotExposeManualRegistration()
    {
        var modules = new ScriptModuleSet([KeyValuePair.Create("main", new ScriptModule("scripts/main.ts",
            "host.functions.register('F', () => 1); host.events.on('cell.click', () => {});", ScriptLanguage.TypeScript))],
            ["main"], scriptApiVersion: 4, bindings: [new("main", "events")]);
        var error = await Assert.ThrowsAsync<ScriptCompilationException>(() => ScriptCompilation.PrepareAsync(modules,
            new(Guid.NewGuid(), 0, 0, 0), default));
        Assert.Contains(error.Diagnostics, d => d.Message.Contains("functions"));
        Assert.Contains(error.Diagnostics, d => d.Message.Contains("events"));
    }

    [Fact]
    public void SaveAndReopenPreserveRolesAndRunInitializationOnlyAfterAuthorization()
    {
        var modules = Modules(("main", "initialization", "export function main(w) { const c=w.getActiveCell(); c.setValue(Number(c.getValue())+1); }"),
            ("events", "events", "export function onCellClick(c) {}"), ("action", "actions", "export function main(w) {}"));
        var path = Path.Combine(Path.GetTempPath(), "quickexcel-conventions-" + Guid.NewGuid().ToString("N") + ".xlsq");
        try
        {
            using (var document = Document(modules)) { document.EnableScriptsAndCalculate(); document.Save(path); }
            var reopened = new TypedWorkbookStorage().Open(path);
            Assert.Equal(4, reopened.ScriptModules.ScriptApiVersion); Assert.Equal(modules.Bindings.ToArray(), reopened.ScriptModules.Bindings.ToArray());
            using var restored = new WorkbookDocument(reopened, new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
            Assert.False(restored.ScriptsEnabled); Assert.Equal("1", restored.GetCellValue(0, 0, 0));
            restored.EnableScriptsAndCalculate(); Assert.Equal("2", restored.GetCellValue(0, 0, 0));
            restored.Calculate(); restored.ListActions(); Assert.Equal("2", restored.GetCellValue(0, 0, 0));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void RejectedEventCommitDisablesTheLostInitializationContext()
    {
        using var document = Document(Modules(("main", "initialization", "export function main() {}"),
            ("events", "events", "export function onWorkbookSaved(c) { c.workbook.getActiveCell().setValue(99); }")));
        document.EnableScriptsAndCalculate(); document.Session.RaiseScriptEvent("workbook.saved");
        Assert.Throws<InvalidOperationException>(() => document.DispatchScriptEvents());
        Assert.False(document.ScriptsEnabled); Assert.Equal("", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void UdfCannotImportTheInitializationModuleIntoItsIsolatedContext()
    {
        using var document = Document(Modules(("main", "initialization", "export const value = 7; export function main() {}"),
            ("udf", "formula-functions", "import {value} from 'main'; export function READ() { return value; }")),
            Sheet("Data", (0, 0, F("=JS.udf.READ()"))));
        Assert.ThrowsAny<Exception>(() => document.EnableScriptsAndCalculate());
        Assert.NotEqual("7", document.GetCellValue(0, 0, 0));
    }
}
