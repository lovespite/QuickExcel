using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Scripting;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M3ScriptTests
{
    [Fact]
    public void V2DescriptorReceivesExplicitRangeShapeAndDistinguishesBlankFromEmptyText()
    {
        var source = """
            quickExcel.functions.register({ name: "JS.SHAPE", apiVersion: 2, parameters: [{name:"cells", type:"range"}],
                returns: "text", profile: "conservative" }, r =>
                [r.rows, r.columns, r.values[0], r.values[1] === "", r.values[2] === null, r.values[3], Object.isFrozen(r.values)].join("|"));
            """;
        var data = Sheet("Data", (0, 0, N(2)), (0, 1, T("")), (1, 1, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))));
        var output = Sheet("S", (0, 0, F("=JS.SHAPE(Data!A1:B2)")));
        var result = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(Book(output, data).WithScriptSource(source), true));
        Assert.Equal(CellValue.FromText("2|2|2|true|true|true|true"), C(result, output).Value);
    }

    [Fact]
    public void V1FlatRangeAndScalarBlankContractIsPreserved()
    {
        var source = """
            quickExcel.functions.register("CHECK", (r, blank) => r.length === 4 && r[1] === "" && r[2] === "" && blank === "");
            """;
        var data = Sheet("Data", (0, 0, N(2)), (1, 1, T("end")));
        var output = Sheet("S", (0, 0, F("=CHECK(Data!A1:B2,Data!C1)")));
        var result = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(Book(data, output).WithScriptSource(source), true));
        Assert.Equal(CellValue.FromBoolean(true), C(result, output).Value);
    }

    [Fact]
    public void AuthorizedRoundsResetClosuresAndAlwaysRecalculateIndependentAndDownstreamCells()
    {
        var source = "let n=0; quickExcel.functions.register('NEXT', () => ++n);";
        var first = Sheet("A", (0, 0, F("=NEXT()")));
        var second = Sheet("B", (0, 0, F("=NEXT()+A!A1")), (0, 1, F("=42")));
        using var session = new DocumentSession(Book(first, second).WithScriptSource(source));
        session.SetScriptsEnabled(true);
        var calculator = new WorkbookCalculator(new JscRuntimeFactory());
        for (var i = 0; i < 3; i++)
        {
            var result = calculator.Calculate(session.BeginCalculation());
            Assert.Equal(0, result.Statistics!.ReusedCount);
            Assert.Equal(3, result.Statistics.EvaluatedCount);
            Assert.Equal(CellValue.FromNumber(1), C(result, first).Value);
            Assert.Equal(CellValue.FromNumber(3), C(result, second).Value);
            session.TryCommitCalculation(result);
        }
        session.SetScriptsEnabled(false);
        var blocked = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(CalculationState.Blocked, C(blocked, second).State);
        Assert.Equal(CellValue.FromNumber(3), C(blocked, second).Value);
        Assert.Equal(CalculationState.Fresh, C(blocked, second, 0, 1).State);
    }

    [Fact]
    public void ScriptConflictsAreRejectedWithoutApplyingResults()
    {
        var sheet = Sheet("S", (0, 0, F("=SUM(1)", 3)));
        var book = Book(sheet).WithScriptSource("quickExcel.functions.register('SUM',()=>99)");
        Assert.Throws<InvalidOperationException>(() => new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(book, true)));
        Assert.Equal(CellValue.FromNumber(3), sheet.GetCell(0, 0)!.Value);
    }

    [Theory]
    [InlineData("quickExcel.functions.register({name:'JS.F', apiVersion:2, profile:'strict'}, ()=>1);")]
    [InlineData("quickExcel.functions.register({name:'F', apiVersion:2}, ()=>1);")]
    [InlineData("quickExcel.functions.register('F',()=>1);quickExcel.functions.register('f',()=>2);")]
    public void InvalidDescriptorOrDuplicateRegistrationFailsLoading(string source)
    {
        Assert.Throws<InvalidOperationException>(() => new JscSession(source));
    }

    [Fact]
    public void ScriptErrorsStayTypedAndUnselectedLazyBranchesAreNeverInvoked()
    {
        var source = "quickExcel.functions.register('FAIL', () => {throw new Error('failed');});";
        var sheet = Sheet("S", (0, 0, F("=IF(FALSE,FAIL(),7)")), (1, 0, F("=IFERROR(FAIL(),8)")),
            (2, 0, F("=FAIL()")), (3, 0, F("=A3+1")), (4, 0, F("=IFERROR(\"#SCRIPT!\",FAIL())")));
        var result = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(Book(sheet).WithScriptSource(source), true));
        Assert.Equal(CellValue.FromNumber(7), C(result, sheet).Value);
        Assert.Equal(CellValue.FromNumber(8), C(result, sheet, 1).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), C(result, sheet, 2).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), C(result, sheet, 3).Value);
        Assert.Equal(CellValue.FromText("#SCRIPT!"), C(result, sheet, 4).Value);
    }

    [Fact]
    public void ErrorArgumentsNeverReachJsAndV2TypesAreValidated()
    {
        var source = """
            quickExcel.functions.register({name:"JS.NUM", apiVersion:2, parameters:[{name:"n",type:"number"}], returns:"number"}, n=>n);
            """;
        var sheet = Sheet("S", (0, 0, F("=JS.NUM(#DIV/0!)")), (1, 0, F("=JS.NUM(\"00123\")")), (2, 0, F("=JS.NUM(4)")));
        var result = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(Request(Book(sheet).WithScriptSource(source), true));
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), C(result, sheet).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), C(result, sheet, 1).Value);
        Assert.Equal(CellValue.FromNumber(4), C(result, sheet, 2).Value);
    }

    [Fact]
    public void CancellationDuringScriptInvocationDisposesRuntimeAndReturnsNothing()
    {
        using var cancellation = new CancellationTokenSource();
        var factory = new CancelFactory(cancellation);
        var sheet = Sheet("S", (0, 0, F("=1", 12)), (0, 1, F("=STOP()", 13)));
        Assert.ThrowsAny<OperationCanceledException>(() => new WorkbookCalculator(factory).Calculate(Request(Book(sheet), true), cancellation.Token));
        Assert.True(factory.Runtime.Disposed);
        Assert.Equal(CellValue.FromNumber(13), sheet.GetCell(0, 1)!.Value);
    }
    private sealed class CancelFactory(CancellationTokenSource cancellation) : IScriptRuntimeFactory
    {
        public CancelRuntime Runtime { get; } = new(cancellation);
        public IScriptRuntime Create(string source) => Runtime;
    }
    private sealed class CancelRuntime(CancellationTokenSource cancellation) : IScriptRuntime
    {
        public bool Disposed { get; private set; }
        public IReadOnlyList<string> FunctionNames => ["STOP"];
        public object? Invoke(string name, IReadOnlyList<object> args) { cancellation.Cancel(); return 1d; }
        public IReadOnlyList<string> GetLogs() => [];
        public void Dispose() => Disposed = true;
    }
}
