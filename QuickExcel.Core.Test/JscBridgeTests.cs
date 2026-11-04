using System.Diagnostics;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Scripting;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class JscBridgeTests
{
    [Fact(Skip = "JSC-MEM-001: 64 MiB engine allocation cap deferred by user; not a progression gate.")]
    public void EngineAllocationHardLimitRemainsOpen()
    {
        using var runtime = new ScriptWorkerRuntimeFactory().Create("host.functions.register('F',()=>new Uint8Array(80*1024*1024).length);");
        Assert.Throws<ScriptRuntimeFailureException>(() => runtime.Invoke("F", []));
    }

    [Fact]
    public void LegacyQueuedJobsStayInertIncludingOnDispose()
    {
        using var runtime = new JscSession("let n=0; Promise.resolve().then(()=>{while(true){}}); host.functions.register('F',()=>++n);");
        Assert.Equal(1d, runtime.Invoke("F", []));
        Assert.Equal(2d, runtime.Invoke("F", []));
    }

    [Theory]
    [InlineData("throw {get stack(){while(true){}},get message(){while(true){}},toString(){while(true){}}};")]
    [InlineData("throw new Error('out of memory; call stack size; time limit');")]
    public void ErrorPropertiesAndTextCannotManufactureFatalNativeStatus(string body)
    {
        using var runtime = new ScriptWorkerRuntimeFactory().Create("host.functions.register('BAD',()=>{" + body + "});host.functions.register('GOOD',()=>9);");
        var clock = Stopwatch.StartNew();
        Assert.Throws<InvalidOperationException>(() => runtime.Invoke("BAD", []));
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(1));
        Assert.Equal(9d, runtime.Invoke("GOOD", []));
    }

    [Fact]
    public void NonTailRecursionInvalidatesNativeVm()
    {
        using var runtime = new JscSession("host.functions.register('F',()=>{function f(){return 1+f()}return f();});");
        Assert.Contains("stack", Assert.Throws<InvalidOperationException>(() => runtime.Invoke("F", [])).Message);
        Assert.True(runtime.LastOperationInvalidated);
        Assert.False(runtime.LastOperationInterrupted);
        Assert.Throws<InvalidOperationException>(() => runtime.Invoke("F", []));
    }

    [Theory]
    [InlineData("try {while(true){}} catch(e) {return 3;}")]
    [InlineData("function f(){return f()} return f();")]
    public void UncatchableWallDeadlineKillsRound(string body)
    {
        using var runtime = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create("host.functions.register('F',()=>{" + body + "});");
        var clock = Stopwatch.StartNew();
        Assert.Throws<TimeoutException>(() => runtime.Invoke("F", []));
        Assert.InRange(clock.Elapsed.TotalMilliseconds, 850, 2000);
        Assert.Throws<ScriptRuntimeFailureException>(() => runtime.Invoke("F", []));
    }

    [Fact]
    public void PendingUserJobsInvalidateWholeBatchWithoutPublishingEarlierResults()
    {
        var sheet = Sheet("S", (0, 0, F("=F(1)", 41)), (1, 0, F("=F(2)", 42)));
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", "host.functions.register('F',x=>{if(x===2)Promise.resolve().then(()=>{while(true){}});return x;});")
        }, ["main"], scriptApiVersion: 3);
        using var session = new DocumentSession(Book(sheet).WithScriptModules(modules));
        session.SetScriptsEnabled(true);
        var before = session.Snapshot;
        var error = Assert.Throws<ScriptRuntimeFailureException>(() => new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(session.BeginCalculation()));
        Assert.Equal("Invalidated", error.Code);
        Assert.Same(before, session.Snapshot);
    }
}
