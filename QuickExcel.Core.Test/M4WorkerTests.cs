using System.Diagnostics;
using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4WorkerTests
{
    private const string IdentitySource = "host.functions.register('IDENTITY', x => x);";
    private static ScriptWorkerRuntime Create(string source, CancellationToken token = default, TimeSpan? limit = null) =>
        (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(source,
            new(Guid.NewGuid(), 1, 2, 3), limit ?? TimeSpan.FromSeconds(5), token));

    [Fact]
    public void WorkerHasSeparateProcessAndPreservesScalarTypesAndBatchOrder()
    {
        using var runtime = Create(IdentitySource);
        Assert.NotEqual(Environment.ProcessId, runtime.ProcessId);
        Assert.Equal(["IDENTITY"], runtime.FunctionNames);
        var values = new object?[] { "00123", 123d, true, null, "", "#VALUE!", "中文 🧮" };
        var results = runtime.InvokeBatch(values.Select(v => new FunctionCall("IDENTITY", [WireValue.FromObject(v)])).ToArray());
        Assert.Equal(values, results.Select(r => r.Value!.ToObject(true)).ToArray());
        Assert.All(results, r => Assert.Null(r.Error));
        var pid = runtime.ProcessId;
        runtime.Dispose();
        AssertExited(pid);
    }

    [Fact]
    public void WorkerAndNativeBridgeAgreeForV1V2ErrorsAndLazyBranches()
    {
        var source = """
            host.functions.register('FLAT', r => r.length === 4 && r[1] === '' && r[2] === '');
            host.functions.register('FAIL', () => { throw new Error('bad'); });
            host.functions.register({name:'JS.SHAPE',apiVersion:2,parameters:[{name:'r',type:'range'}],returns:'text'},
              r => [r.rows,r.columns,r.values[1] === null,r.values[2] === null,Object.isFrozen(r.values)].join('|'));
            """;
        var data = Sheet("Data", (0, 0, N(2)), (1, 1, T("end")));
        var output = Sheet("Out", (0, 0, F("=FLAT(Data!A1:B2)")), (1, 0, F("=JS.SHAPE(Data!A1:B2)")),
            (2, 0, F("=IF(FALSE,FAIL(),7)")), (3, 0, F("=IFERROR(FAIL(),8)")), (4, 0, F("=FAIL()")), (5, 0, F("=\"#SCRIPT!\"")));
        var request = Request(Book(data, output).WithScriptSource(source), true);
        var native = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(request);
        var worker = new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(request);
        Assert.Equal(native.Cells.OrderBy(p => p.Key.Row).Select(p => p.Value.Value), worker.Cells.OrderBy(p => p.Key.Row).Select(p => p.Value.Value));
        Assert.Equal(CellValue.FromText("2|2|true|true|true"), C(worker, output, 1).Value);
    }

    [Fact]
    public void ContextsAndLogsAreIsolatedAndConservativeRoundsNeverReuseUdfResults()
    {
        const string source = "let n=0; host.functions.register('NEXT',()=>++n);console.log('private A');";
        using var first = Create(source);
        using var second = Create("host.functions.register('NEXT',()=>40);console.log('private B');");
        Assert.NotEqual(first.WorkerInstanceId, second.WorkerInstanceId);
        Assert.Equal(1d, first.Invoke("NEXT", [])); Assert.Equal(2d, first.Invoke("NEXT", []));
        Assert.Equal(40d, second.Invoke("NEXT", []));
        Assert.Equal(["private A"], first.GetLogs()); Assert.Equal(["private B"], second.GetLogs());
        var sheet = Sheet("S", (0, 0, F("=NEXT()")), (1, 0, F("=A1+NEXT()")));
        using var session = new DocumentSession(Book(sheet).WithScriptSource(source));
        session.SetScriptsEnabled(true);
        var calculator = new WorkbookCalculator(new ScriptWorkerRuntimeFactory());
        for (var i = 0; i < 2; i++)
        {
            var result = calculator.Calculate(session.BeginCalculation());
            Assert.Equal(0, result.Statistics!.ReusedCount);
            Assert.Equal(CellValue.FromNumber(OperatingSystem.IsWindows() ? 3 + 4 * i : 3), C(result, sheet, 1).Value);
            Assert.True(session.TryCommitCalculation(result));
        }
    }

    [Fact]
    public void UnicodeLogsHaveCountAndUtf8ByteLimits()
    {
        using var runtime = Create("for(let i=0;i<200;i++) console.log('中🧮'.repeat(2000));");
        var logs = runtime.GetLogs();
        Assert.Equal(100, logs.Count);
        Assert.All(logs, log => Assert.InRange(Encoding.UTF8.GetByteCount(log), 1, 4000));
    }

    [Theory]
    [InlineData("while(true) {}")]
    [InlineData("host.functions.register('F',()=>1); throw new Error('bad load');")]
    [InlineData("this is invalid js !!")]
    public void LoadingFailureNeverPublishesPartialRegistryOrFallsBack(string source)
    {
        Assert.Throws<ScriptRuntimeFailureException>(() => Create(source));
        using var next = Create(IdentitySource);
        Assert.Equal(3d, next.Invoke("IDENTITY", [3d]));
    }

    [Fact]
    public void MissingWorkerDoesNotRunInProcess()
    {
        var factory = new ScriptWorkerRuntimeFactory(Path.Combine(AppContext.BaseDirectory, "missing-worker.exe"));
        Assert.Equal("WorkerMissing", Assert.Throws<ScriptRuntimeFailureException>(() => factory.Create(IdentitySource)).Code);
    }

    [Fact]
    public void SynchronousFacadeNeverNeedsCallerSynchronizationContextToPump()
    {
        var previous = SynchronizationContext.Current;
        var context = new CountingContext();
        try
        {
            SynchronizationContext.SetSynchronizationContext(context);
            using var runtime = Create(IdentitySource);
            Assert.Equal(42d, runtime.Invoke("IDENTITY", [42d]));
            Assert.Equal(0, context.Posts);
        }
        finally { SynchronizationContext.SetSynchronizationContext(previous); }
    }

    private sealed class CountingContext : SynchronizationContext
    {
        public int Posts;
        public override void Post(SendOrPostCallback callback, object? state)
        { Interlocked.Increment(ref Posts); ThreadPool.QueueUserWorkItem(_ => callback(state)); }
    }

    [Fact]
    public void NativeTimeoutKillsRoundAndNeverReusesHandle()
    {
        using var runtime = Create("host.functions.register('LOOP',()=>{while(true){}});");
        var clock = Stopwatch.StartNew();
        Assert.Throws<TimeoutException>(() => runtime.Invoke("LOOP", []));
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(4));
        AssertExited(runtime.ProcessId);
        Assert.Throws<ScriptRuntimeFailureException>(() => runtime.Invoke("LOOP", []));
    }

    [Fact]
    public async Task CancellationTerminatesWorkerBeforeNativeDeadline()
    {
        using var cancel = new CancellationTokenSource();
        using var runtime = Create("host.functions.register('LOOP',()=>{while(true){}});", cancel.Token);
        var invoke = Task.Run(() => runtime.Invoke("LOOP", []));
        await Task.Delay(50);
        var clock = Stopwatch.StartNew(); cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => invoke.WaitAsync(TimeSpan.FromSeconds(3)));
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(2));
        AssertExited(runtime.ProcessId);
    }

    // JSC deliberately does not invoke exception accessors; its corresponding
    // tests also distinguish proper tail calls from actual stack growth.

    [Fact]
    public async Task HostDeadlineKillsIdleWorkerWithoutAnotherInvocation()
    {
        using var runtime = Create(IdentitySource, limit: TimeSpan.FromMilliseconds(800));
        using var process = Process.GetProcessById(runtime.ProcessId);
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
        Assert.Throws<TimeoutException>(() => runtime.Invoke("IDENTITY", [1d]));
    }

    [Fact]
    public async Task ForcedExitRejectsWholeRoundButDocumentRemainsEditableAndSavable()
    {
        var factory = new ObservedFactory();
        var sheet = Sheet("S", (0, 0, F("=9", 12)), (0, 1, F("=LOOP()", 13)));
        using var document = new WorkbookDocument(Book(sheet).WithScriptSource("host.functions.register('LOOP',()=>{while(true){}});"),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(factory)));
        document.Session.SetScriptsEnabled(true);
        var work = Task.Run(() => document.Calculate(CancellationToken.None));
        var runtime = await factory.Ready.Task.WaitAsync(TimeSpan.FromSeconds(4));
        using (var process = Process.GetProcessById(runtime.ProcessId)) process.Kill();
        await Assert.ThrowsAsync<ScriptRuntimeFailureException>(() => work.WaitAsync(TimeSpan.FromSeconds(4)));
        Assert.Equal("12", document.GetCellValue(0, 0, 0)); Assert.Equal("13", document.GetCellValue(0, 0, 1));
        document.SetCell(0, 1, 0, "editable");
        var file = Path.Combine(Path.GetTempPath(), "quickexcel-worker-" + Guid.NewGuid() + ".xlsq");
        try { document.Save(file); using var reopened = WorkbookDocument.Open(file, new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()))); Assert.False(reopened.ScriptsEnabled); Assert.Equal("editable", reopened.GetCellValue(0, 1, 0)); }
        finally { File.Delete(file); }
        document.ScriptSource = "host.functions.register('LOOP',()=>7);";
        document.EnableScriptsAndCalculate();
        Assert.Equal("7", document.GetCellValue(0, 0, 1));
    }

    [Theory]
    [InlineData("edit")] [InlineData("source")] [InlineData("disable")] [InlineData("close")] [InlineData("new-round")]
    public async Task SessionInvalidationTerminatesItsWorkerAndRejectsLateResults(string change)
    {
        var factory = new ObservedFactory();
        var sheet = Sheet("S", (0, 0, F("=LOOP()", 13)));
        using var document = new WorkbookDocument(Book(sheet).WithScriptSource("host.functions.register('LOOP',()=>{while(true){}});"),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(factory)));
        document.Session.SetScriptsEnabled(true);
        var work = Task.Run(() => document.Calculate(CancellationToken.None));
        var runtime = await factory.Ready.Task.WaitAsync(TimeSpan.FromSeconds(4));
        switch (change)
        {
            case "edit": document.SetCell(0, 1, 0, "new"); break;
            case "source": document.ScriptSource = "// replacement"; break;
            case "disable": document.DisableScripts(); break;
            case "close": document.Dispose(); break;
            case "new-round": document.Session.BeginCalculation(); break;
        }
        Assert.False(await work.WaitAsync(TimeSpan.FromSeconds(4)));
        AssertExited(runtime.ProcessId);
        if (!document.Session.IsClosed) Assert.Equal("13", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public async Task OverlappingRequestsAreRejectedWithoutGrowingQueue()
    {
        using var cancellation = new CancellationTokenSource();
        using var runtime = Create("host.functions.register('LOOP',()=>{while(true){}});", cancellation.Token);
        var work = Task.Run(() => runtime.Invoke("LOOP", []));
        await Task.Delay(100);
        Assert.Equal("Busy", Assert.Throws<ScriptRuntimeFailureException>(() => runtime.GetLogs()).Code);
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work);
    }

    private sealed class ObservedFactory : IScriptRuntimeFactory
    {
        public TaskCompletionSource<ScriptWorkerRuntime> Ready { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public IScriptRuntime Create(string source) => throw new InvalidOperationException("Stamped request required.");
        public IScriptRuntime Create(ScriptRuntimeRequest request)
        {
            var runtime = new ScriptWorkerRuntimeFactory().Create(request);
            Ready.TrySetResult(runtime is ScriptWorkerLease lease ? lease.Runtime : (ScriptWorkerRuntime)runtime);
            return runtime;
        }
    }
    private static void AssertExited(int pid)
    {
        try { using var process = Process.GetProcessById(pid); Assert.True(process.WaitForExit(1000)); }
        catch (ArgumentException) { }
    }
}
