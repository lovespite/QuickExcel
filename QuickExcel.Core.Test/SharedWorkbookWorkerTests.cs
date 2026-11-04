using System.Diagnostics;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SharedWorkbookWorkerTests
{
    private static WorkbookDocument Document(string source) => new(Book(Sheet("Data")).WithScriptModules(M4ActionTests.Actions(source)),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));

    private static int WorkerId(WorkbookDocument document, bool actions = true)
    {
        var request = document.Session.BeginCalculation();
        using var lease = SharedScriptWorkers.Acquire(Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"),
            new ScriptRuntimeRequest(request.Snapshot.ScriptModules, request.Stamp, TimeSpan.FromSeconds(5), request.CancellationToken)
            { RuntimeLifetime = request.RuntimeLifetime }, actions);
        return lease.Runtime.ProcessId;
    }

    [Fact]
    public async Task WindowsKeepsOneIsolatedProcessAndContextAcrossIdleTimeAndReleasesItWhenDisabled()
    {
        if (!OperatingSystem.IsWindows()) return;
        using var document = Document("let count=0; function main(workbook){ workbook.getActiveCell().setValue(++count); }");
        document.EnableScriptsAndCalculate();
        document.ListActions(); var first = WorkerId(document);
        document.RunAction("A"); Assert.Equal("1", document.GetCellValue(0, 0, 0));
        Assert.Equal(first, WorkerId(document));
        await Task.Delay(5200); // The deadline is per active round, not time since process creation.
        document.RunAction("A"); Assert.Equal("2", document.GetCellValue(0, 0, 0));
        Assert.Equal(first, WorkerId(document));
        document.DisableScripts();
        try { using var process = Process.GetProcessById(first); Assert.True(process.WaitForExit(2000)); } catch (ArgumentException) { }
        document.EnableScriptsAndCalculate(); document.RunAction("A"); Assert.Equal("1", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void WorkbookStateIsSeparateAndFailureRebuildsTheContextWithoutReplayingWrites()
    {
        if (!OperatingSystem.IsWindows()) return;
        const string source = "let n=0; function main(workbook){ workbook.getActiveCell().setValue(++n); if(n===3) throw Error('fail'); }";
        using var first = Document(source); using var second = Document(source);
        first.EnableScriptsAndCalculate(); second.EnableScriptsAndCalculate();
        first.RunAction("A"); first.RunAction("A"); second.RunAction("A");
        Assert.Equal("2", first.GetCellValue(0, 0, 0)); Assert.Equal("1", second.GetCellValue(0, 0, 0));
        Assert.NotEqual(WorkerId(first), WorkerId(second));
        Assert.Throws<ScriptRuntimeFailureException>(() => first.RunAction("A"));
        Assert.Equal("2", first.GetCellValue(0, 0, 0));
        first.RunAction("A"); Assert.Equal("1", first.GetCellValue(0, 0, 0));
        second.RunAction("A"); Assert.Equal("2", second.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void StateOnlyActionsCannotChangeFormulaStateOrInvalidateItsCache()
    {
        if (!OperatingSystem.IsWindows()) return;
        var source = "let n=0; host.functions.register('STATE',()=>n); function main(){n++;}";
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, F("=STATE()")))).WithScriptModules(M4ActionTests.Actions(source, true)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        Assert.Equal(CalculationState.Fresh, doc.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
        Assert.False(doc.RunAction("A").Changed); Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.IsDirty);
        Assert.Equal(CalculationState.Fresh, doc.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
        doc.Calculate(); Assert.Equal("0", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void RejectedHostCommitDiscardsProvisionalExecutionStateWithoutPublishingWrites()
    {
        if (!OperatingSystem.IsWindows()) return;
        var source = """
            let n=0;
            host.functions.register('STATE',()=>n);
            function main(){n++;}
            host.events.on('workbook.calculated',context=>{n=99;context.workbook.getActiveWorksheet().getCell('B1').setValue(99);});
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, F("=STATE()")))).WithScriptModules(M4ActionTests.Actions(source, true)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); doc.ListActions(); doc.RunAction("A"); doc.Calculate();
        Assert.Equal("0", doc.GetCellValue(0, 0, 0));
        var revision = doc.Snapshot.Revision;
        Assert.NotNull(Record.Exception(() => doc.DispatchScriptEvents()));
        Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.CanUndo);
        Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
        Assert.NotEqual(CalculationState.Fresh, doc.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
        doc.Calculate(); Assert.Equal("0", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void SharedModuleRoleCannotBeForgedByReplacingGlobalPropertiesOrRegisteringFromAnotherEntry()
    {
        if (!OperatingSystem.IsWindows()) return;
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["a"] = new("scripts/a.js", "globalThis.__qxRegistrationRole=()=>3; host.functions.register('WRONG',()=>77); export function main(workbook){ workbook.getActiveCell().setValue(1); }"),
            ["f"] = new("scripts/f.js", "host.functions.register('BOUNDVALUE',()=>42);"),
            ["unbound"] = new("scripts/unbound.js", "throw Error('unbound entry must not initialize');")
        }, ["a", "f"], scriptApiVersion: 3, bindings: [new("a", "actions"), new("f", "formula-functions")]);
        using var doc = new WorkbookDocument(Book(Sheet("Data", (1, 0, F("=BOUNDVALUE()")), (2, 0, F("=WRONG()")))).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate();
        Assert.Equal("a", Assert.Single(doc.ListActions()).Name);
        Assert.Equal("42", doc.GetCellValue(0, 1, 0));
        Assert.Equal(CalculationState.Blocked, doc.Snapshot.Sheets[0].GetCell(2, 0)!.Calculation!.State);
        doc.Session.RaiseScriptEvent("workbook.activate"); Assert.False(doc.Session.HasScriptEvents);
    }

    [Fact]
    public void FormulaAndActionContextsAreBothReusedButTheirClosuresAndGlobalsAreSeparate()
    {
        if (!OperatingSystem.IsWindows()) return;
        const string source = """
            let calls=0;
            globalThis.marker=10;
            host.functions.register('NEXT',()=>++calls + globalThis.marker);
            function main(){calls=100;globalThis.marker=1000;}
            host.events.on('workbook.activate',()=>{calls=200;globalThis.marker=2000;});
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, F("=NEXT()")))).WithScriptModules(M4ActionTests.Actions(source, true)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); Assert.Equal("11", doc.GetCellValue(0, 0, 0));
        var formulaWorker = WorkerId(doc, actions: false); var actionWorker = WorkerId(doc);
        Assert.NotEqual(formulaWorker, actionWorker);
        doc.ListActions(); doc.RunAction("A");
        doc.Session.RaiseScriptEvent("workbook.activate"); doc.DispatchScriptEvents();
        doc.Calculate(); Assert.Equal("12", doc.GetCellValue(0, 0, 0));
        doc.Calculate(); Assert.Equal("13", doc.GetCellValue(0, 0, 0));
        Assert.Equal(formulaWorker, WorkerId(doc, actions: false)); Assert.Equal(actionWorker, WorkerId(doc));
        doc.DisableScripts();
        foreach (var id in new[] { formulaWorker, actionWorker })
            try { using var process = Process.GetProcessById(id); Assert.True(process.WaitForExit(2000)); } catch (ArgumentException) { }
    }

    [Fact]
    public void FormulaWorkerDoesNotInitializeActionRootsAndCannotAccessWorkbookObjects()
    {
        if (!OperatingSystem.IsWindows()) return;
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["actions"] = new("scripts/actions.js", "throw Error('action root must not run during calculation');"),
            ["formulas"] = new("scripts/formulas.js", """
                host.functions.register('SAFE',()=>21);
                host.functions.register('READ',()=>workbook.getActiveWorksheet().getCell('A1').getValue());
                host.functions.register('WRITE',()=>{workbook.getActiveWorksheet().getCell('A1').setValue(99);return 0;});
                """)
        }, ["actions", "formulas"], scriptApiVersion: 3, bindings: [new("actions", "actions"), new("formulas", "formula-functions")]);
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, N(7)), (0, 1, F("=SAFE()")), (0, 2, F("=READ()")), (0, 3, F("=WRITE()")))).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate();
        Assert.Equal("7", doc.GetCellValue(0, 0, 0)); Assert.Equal("21", doc.GetCellValue(0, 0, 1));
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), doc.Snapshot.Sheets[0].GetCell(0, 2)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), doc.Snapshot.Sheets[0].GetCell(0, 3)!.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task ReusableWorkerRejectsMixedBindingsAndChangingItsRole(bool actions, bool mixedBindings)
    {
        if (!OperatingSystem.IsWindows()) return;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        using var process = WorkerProcess.Start(Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"), timeout.Token);
        var identity = M4ProtocolTests.Identity() with { CalculationGeneration = actions ? null : 1 };
        async Task Send<T>(string kind, T payload)
        {
            identity = identity with { RequestId = Guid.NewGuid() };
            await WorkerProtocol.WriteAsync(process.Input, WorkerProtocol.Message(identity, kind, payload), timeout.Token);
        }
        await Send("hello", new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [3], WorkerProtocol.MaxFrameBytes));
        _ = await WorkerProtocol.ReadAsync(process.Output, timeout.Token);
        var bundle = new LoadModules([new("main", "scripts/main.js", mixedBindings ? "while(true){}" : "")], ["main"], "")
        { FunctionEntries = !actions || mixedBindings ? ["main"] : [], ActionEntries = actions || mixedBindings ? ["main"] : [] };
        await Send("loadWorkbook", bundle with { Sha256 = bundle.ComputeDigest() });
        if (!mixedBindings)
        {
            _ = await WorkerProtocol.ReadAsync(process.Output, timeout.Token);
            await Send("endRound", new EmptyPayload());
            _ = await WorkerProtocol.ReadAsync(process.Output, timeout.Token);
            identity = identity with { RoundId = Guid.NewGuid(), CalculationGeneration = actions ? 1 : null };
            await Send("beginRound", new BeginWorkerRound());
        }
        await Assert.ThrowsAsync<EndOfStreamException>(() => WorkerProtocol.ReadAsync(process.Output, timeout.Token));
    }
}
