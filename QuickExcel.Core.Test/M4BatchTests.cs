using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4BatchTests
{
    private sealed class ObservedFactory(string? executable = null) : IScriptRuntimeFactory
    {
        public ScriptWorkerRuntime? Runtime { get; private set; }
        public IScriptRuntime Create(string source) => throw new NotSupportedException();
        public IScriptRuntime Create(ScriptRuntimeRequest request)
        {
            var runtime = new ScriptWorkerRuntimeFactory(executable).Create(request);
            Runtime = runtime as ScriptWorkerRuntime; // Metrics below use explicit one-shot requests.
            return runtime;
        }
    }

    private static CalculationRequest Request(WorkbookSnapshot book) => new(new(Guid.NewGuid(), 0, 1, 1), book, true);

    [Fact]
    public void IndependentCallsBatchWithinReplyBudgetAndKeepStatefulOrderAcrossRounds()
    {
        var sheet = Sheet("S", Enumerable.Range(0, 600).Select(i => (i, 0, F("=NEXT()"))).ToArray());
        var book = Book(sheet).WithScriptSource("let n=0; host.functions.register('NEXT',()=>++n);");
        var factory = new ObservedFactory();
        var calculator = new WorkbookCalculator(factory);
        var request = Request(book);
        for (var round = 0; round < 2; round++)
        {
            var result = calculator.Calculate(request with { Stamp = request.Stamp with { CalculationGeneration = round + 1 } });
            Assert.Equal(8, factory.Runtime!.InvocationBatchCount);
            Assert.Equal(600, factory.Runtime.InvocationCount);
            Assert.Equal(0, result.Statistics!.ReusedCount);
            for (var row = 0; row < 600; row++) Assert.Equal(CellValue.FromNumber(row + 1), C(result, sheet, row).Value);
        }
    }

    [Fact]
    public void DependenciesLazyBranchesNestedCallsAndErrorsMatchSerialReference()
    {
        // A newly ready, earlier address must run before previously ready later addresses.
        var a = Sheet("A", (0, 0, F("=NEXT(B!A1)")), (1, 0, F("=NEXT(A1)")),
            (2, 0, F("=IF(FALSE,NEXT(),NEXT())")), (3, 0, F("=NEXT(NEXT())+NEXT()")),
            (4, 0, F("=IFERROR(BAD(),NEXT())")), (5, 0, F("=BAD()")),
            (6, 0, F("=NEXT(1/0)")), (7, 0, F("=NEXT()")), (8, 0, F("=NEXT(\"bad\"+1)")),
            (9, 0, F("=NEXT()")), (10, 0, F("=IFERROR(NEXT(1/0),NEXT())")));
        var b = Sheet("B", (0, 0, F("=NEXT()")), (1, 0, F("=NEXT()")));
        var book = Book(a, b).WithScriptSource("let n=0; host.functions.register('NEXT',()=>++n); host.functions.register('BAD',()=>{n++;throw Error('bad');});");
        var serial = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(book));
        var factory = new ObservedFactory();
        var batched = new WorkbookCalculator(factory).Calculate(Request(book));
        foreach (var (address, cell) in serial.Cells)
        {
            Assert.Equal(cell.Value, batched.Cells[address].Value);
            Assert.Equal(cell.State, batched.Cells[address].State);
            Assert.Equal(cell.Diagnostics, batched.Cells[address].Diagnostics);
        }
        Assert.True(factory.Runtime!.InvocationBatchCount < factory.Runtime.InvocationCount);
    }

    [Fact]
    public void TypedRangesErrorsAndDescriptorFailuresPreserveCellSemantics()
    {
        var data = Sheet("Data", (0, 0, T("00123")), (0, 1, new CellSnapshot(CellContent.FromLiteral(CellValue.FromBoolean(true)))), (1, 0, N(9)), (2, 0, F("=1/0")));
        var output = Sheet("S", (0, 0, F("=JS.R(Data!A1:B2)")), (1, 0, F("=JS.R(Data!A1:B2)")),
            (2, 0, F("=JS.R(Data!A3)")), (3, 0, F("=JS.R()")), (4, 0, F("=JS.WRONG()")),
            (5, 0, F("=JS.LONG()")), (6, 0, F("=JS.R(Data!A1:B2)")));
        var book = Book(output, data).WithScriptSource("""
            host.functions.register({name:'JS.R',apiVersion:2,parameters:[{name:'r',type:'range'}],returns:'text'},r=>JSON.stringify(r));
            host.functions.register({name:'JS.WRONG',apiVersion:2,returns:'number'},()=>true);
            host.functions.register({name:'JS.LONG',apiVersion:2},()=>'x'.repeat(32768));
            """);
        var serial = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(book));
        var batched = new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(Request(book));
        foreach (var (address, cell) in serial.Cells)
        {
            Assert.Equal(cell.Value, batched.Cells[address].Value);
            Assert.Equal(cell.State, batched.Cells[address].State);
        }
    }

    [Fact]
    public void LargeEscapedArgumentsSplitBeforeFrameLimitWithoutChangingRangeShape()
    {
        var data = Sheet("Data", Enumerable.Range(0, 10).Select(i => (i, 0, T(new string('中', 32767)))).ToArray());
        var output = Sheet("S", Enumerable.Range(0, 12).Select(i => (i, 0, F("=SIZE(Data!A1:A10)"))).ToArray());
        var book = Book(output, data).WithScriptSource("host.functions.register('SIZE',r=>r[0].length);");
        var factory = new ObservedFactory();
        var result = new WorkbookCalculator(factory).Calculate(Request(book));
        Assert.Equal(2, factory.Runtime!.InvocationBatchCount);
        Assert.Equal(12, factory.Runtime.InvocationCount);
        Assert.All(result.Cells.Values, c => Assert.Equal(CellValue.FromNumber(32767), c.Value));
    }

    [Fact]
    public void BatchIsSingleUseAndIndividuallyOversizedArgumentsFailBeforeSending()
    {
        using var runtime = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create("host.functions.register('F',()=>1);");
        var batch = runtime.CreateBatch();
        Assert.True(batch.TryAdd("F", []));
        Assert.Single(batch.Invoke());
        Assert.Throws<InvalidOperationException>(() => batch.Invoke());
        Assert.Throws<InvalidOperationException>(() => batch.TryAdd("F", []));
        var large = runtime.CreateBatch();
        Assert.Equal("InputLimit", Assert.Throws<ScriptRuntimeFailureException>(() => large.TryAdd("F", [new string('中', 3_000_000)])).Code);
        Assert.Equal(1, runtime.InvocationCount);
    }

    [Theory]
    [InlineData("bad-count")]
    [InlineData("bad-value")]
    [InlineData("bad-limit")]
    public void MalformedBatchNeverCommitsEarlierCells(string fault)
    {
        var sheet = Sheet("S", (0, 0, F("=F()", 41)), (1, 0, F("=F()", 42)));
        using var session = new DocumentSession(Book(sheet).WithScriptSource(fault));
        session.SetScriptsEnabled(true);
        var factory = new ObservedFactory(Path.Combine(AppContext.BaseDirectory, "WorkerFixture", "WorkerFixture.exe"));
        var before = session.Snapshot;
        Assert.Equal("Protocol", Assert.Throws<ScriptRuntimeFailureException>(() => new WorkbookCalculator(factory).Calculate(session.BeginCalculation())).Code);
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void TimeoutAfterSuccessfulCallRejectsWholeBatchAndPreservesOldCache()
    {
        var sheet = Sheet("S", (0, 0, F("=F(1)", 41)), (1, 0, F("=F(2)", 42)));
        using var session = new DocumentSession(Book(sheet).WithScriptSource("host.functions.register('F',x=>{if(x===2)while(true){};return x;});"));
        session.SetScriptsEnabled(true);
        var before = session.Snapshot;
        Assert.Throws<TimeoutException>(() => new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(session.BeginCalculation()));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void WorstCaseValidTextRepliesFitAndOversizedTextRemainsACellError()
    {
        var sheet = Sheet("S", Enumerable.Range(0, 45).Select(i => (i, 0, F($"=PAYLOAD({i})"))).ToArray());
        var book = Book(sheet).WithScriptSource("host.functions.register('PAYLOAD',n=>'中'.repeat(n===22?32768:32767));");
        var factory = new ObservedFactory();
        var result = new WorkbookCalculator(factory).Calculate(Request(book));
        Assert.Equal(1, factory.Runtime!.InvocationBatchCount);
        for (var row = 0; row < 45; row++)
            Assert.Equal(row == 22 ? CellValue.FromError(CellErrorCode.Value) : CellValue.FromText(new string('中', 32767)), C(result, sheet, row).Value);
    }
}
