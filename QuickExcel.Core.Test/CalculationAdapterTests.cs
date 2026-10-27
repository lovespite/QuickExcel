using QuickExcel.Calculation;
using QuickExcel.Runtime;
using System.Globalization;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class CalculationAdapterTests
{
    [Fact]
    public void SparseSnapshotCalculatesDependenciesWithoutMutatingInputsOrAllocatingRows()
    {
        var snapshot = Snapshot((99_999, 0, Literal(CellValue.FromNumber(7))),
            (0, 0, Formula("=B1+1")), (0, 1, Formula("=SUM(A100000:A100000)*2")));
        var request = Request(snapshot);

        var result = new LegacyWorkbookCalculator(new RecordingFactory()).Calculate(request);

        Assert.Equal(request.Stamp, result.Stamp);
        Assert.Equal(CellValue.FromNumber(15), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(14), Value(result, snapshot, 0, 1));
        Assert.Equal(3, snapshot.Sheets[0].Cells.Count);
        Assert.Null(snapshot.Sheets[0].GetCell(0, 0)!.Calculation);
        Assert.Equal(2, result.Cells.Count);
        Assert.All(result.Cells.Values, cell =>
        {
            Assert.Equal(CalculationState.Fresh, cell.State);
            Assert.Equal(request.Stamp.InputRevision, cell.InputRevision);
            Assert.Equal(request.Stamp.CalculationGeneration, cell.CalculationGeneration);
            Assert.Equal(request.Stamp.ScriptGeneration, cell.ScriptGeneration);
        });
    }

    [Fact]
    public void DirectReferencesKeepTextBooleanBlankAndErrorIdentities()
    {
        var snapshot = Snapshot((0, 0, Literal(CellValue.FromText("00123"))),
            (1, 0, Literal(CellValue.FromText("#VALUE!"))), (2, 0, Literal(CellValue.FromBoolean(true))),
            (0, 1, Formula("=A1")), (1, 1, Formula("=A2")), (2, 1, Formula("=A3")),
            (3, 1, Formula("=A4")), (4, 1, Formula("=UPPER(\"#value!\")")),
            (5, 0, Literal(CellValue.FromText(""))), (5, 1, Formula("=A6")));

        var result = Calculate(snapshot);

        Assert.Equal(CellValue.FromText("00123"), Value(result, snapshot, 0, 1));
        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 1, 1));
        Assert.Equal(CellValue.FromBoolean(true), Value(result, snapshot, 2, 1));
        Assert.Equal(CellValue.Blank, Value(result, snapshot, 3, 1));
        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 4, 1));
        Assert.Equal(CellValue.FromText(""), Value(result, snapshot, 5, 1));
    }

    [Fact]
    public void ErrorsPropagateAsErrorsAndIfErrorDoesNotConsumeLookalikeText()
    {
        var snapshot = Snapshot((0, 0, Formula("=1/0")), (0, 1, Formula("=A1+1")),
            (1, 0, Formula("=IFERROR(A1,5)")), (1, 1, Formula("=IFERROR(\"#VALUE!\",5)")),
            (2, 0, Formula("=IF(FALSE,1/0,42)")), (2, 1, Formula("=10^309")));

        var result = Calculate(snapshot);

        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), Value(result, snapshot, 0, 1));
        Assert.Equal(CellValue.FromNumber(5), Value(result, snapshot, 1, 0));
        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 1, 1));
        Assert.Equal(CellValue.FromNumber(42), Value(result, snapshot, 2, 0));
        Assert.Equal(CellValue.FromError(CellErrorCode.Number), Value(result, snapshot, 2, 1));
    }

    [Fact]
    public void DisabledScriptsNeverCreateRuntimeAndBlockedCachesPropagate()
    {
        var factory = new RecordingFactory();
        var snapshot = Snapshot((0, 0, Formula("=PRIVATE(1)", CellValue.FromNumber(8))),
            (0, 1, Formula("=A1+1", CellValue.FromNumber(9))), (0, 2, Formula("=SUM(A1:B1)", CellValue.FromNumber(17))),
            (1, 0, Formula("=2+3")), (1, 1, Formula("=PRIVATE(2)")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot));

        Assert.Empty(factory.Created);
        Assert.Equal(CellValue.FromNumber(8), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(9), Value(result, snapshot, 0, 1));
        Assert.Equal(CellValue.FromNumber(17), Value(result, snapshot, 0, 2));
        Assert.Equal(CalculationState.Blocked, Cell(result, snapshot, 0, 2).State);
        Assert.Equal(CellValue.FromNumber(5), Value(result, snapshot, 1, 0));
        Assert.Null(Cell(result, snapshot, 1, 1).Value);
        Assert.Equal(CalculationState.Blocked, Cell(result, snapshot, 1, 1).State);
        Assert.NotEmpty(result.Diagnostics);
    }

    [Fact]
    public void CrossSheetReferencesAndDependentsRetainBlockedCaches()
    {
        var snapshot = Snapshot((0, 0, Formula("='Other sheet'!A1", CellValue.FromNumber(25))),
            (0, 1, Formula("=A1*2", CellValue.FromNumber(50))),
            (0, 2, Formula("=\"literal!\"")));

        var result = Calculate(snapshot);

        Assert.Equal(CellValue.FromNumber(25), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(50), Value(result, snapshot, 0, 1));
        Assert.Equal(CalculationState.Blocked, Cell(result, snapshot, 0, 1).State);
        Assert.Equal(CellValue.FromText("literal!"), Value(result, snapshot, 0, 2));
    }

    [Fact]
    public void CyclesAndTheirDependentsBecomeCycleErrorsWhileIndependentCellsFinish()
    {
        var snapshot = Snapshot((0, 0, Formula("=B1+1")), (0, 1, Formula("=A1+1")),
            (0, 2, Formula("=SUM(A1:B1)")), (0, 3, Formula("=40+2")));

        var result = Calculate(snapshot);

        for (var column = 0; column < 3; column++)
        {
            Assert.Equal(CellValue.FromError(CellErrorCode.Cycle), Value(result, snapshot, 0, column));
            Assert.Equal(CalculationState.Error, Cell(result, snapshot, 0, column).State);
        }
        Assert.Equal(CellValue.FromNumber(42), Value(result, snapshot, 0, 3));
    }

    [Theory]
    [InlineData("=SUM(")]
    [InlineData("=1+")]
    [InlineData("=#VALUE!")]
    public void MalformedFormulasProduceAddressedDiagnosticsWithoutBlockingOtherCells(string formula)
    {
        var snapshot = Snapshot((0, 0, Formula(formula, CellValue.FromNumber(99))), (0, 1, Formula("=3+4")));

        var result = Calculate(snapshot);

        Assert.Equal(CalculationState.Error, Cell(result, snapshot, 0, 0).State);
        Assert.NotEmpty(Cell(result, snapshot, 0, 0).Diagnostics);
        Assert.Contains(result.Diagnostics, message => message.Contains("Sheet1!A1"));
        Assert.Equal(CellValue.FromNumber(7), Value(result, snapshot, 0, 1));
    }

    [Fact]
    public void MissingAuthorizedFunctionBecomesNameErrorInsteadOfKeepingCache()
    {
        var snapshot = Snapshot((0, 0, Formula("=MISSING()", CellValue.FromNumber(12))));

        var result = new LegacyWorkbookCalculator(new RecordingFactory()).Calculate(Request(snapshot, true));

        Assert.Equal(CellValue.FromError(CellErrorCode.Name), Value(result, snapshot, 0, 0));
        Assert.Equal(CalculationState.Error, Cell(result, snapshot, 0, 0).State);
    }

    [Fact]
    public void EveryAuthorizedRoundCreatesAndDisposesItsOwnRuntime()
    {
        var factory = new RecordingFactory();
        var calculator = new LegacyWorkbookCalculator(factory);
        var snapshot = Snapshot((0, 0, Formula("=NEXT()")), (1, 0, Formula("=NEXT()")));

        var first = calculator.Calculate(Request(snapshot, true));
        var second = calculator.Calculate(Request(snapshot, true));

        Assert.Equal(2, factory.Created.Count);
        Assert.All(factory.Created, runtime => Assert.True(runtime.Disposed));
        Assert.Equal(CellValue.FromNumber(1), Value(first, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(2), Value(first, snapshot, 1, 0));
        Assert.Equal(CellValue.FromNumber(1), Value(second, snapshot, 0, 0));
        Assert.Equal(snapshot.ScriptSource, factory.Sources[0]);
    }

    [Fact]
    public void BuiltinLazyBranchesDoNotInvokeUnselectedScriptCalls()
    {
        var calls = 0;
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, _) => { calls++; throw new InvalidOperationException("unselected branch"); }
        });
        var snapshot = Snapshot((0, 0, Formula("=IF(FALSE,NEXT(),7)")),
            (0, 1, Formula("=IFERROR(9,NEXT())")), (0, 2, Formula("=IFERROR(\"#VALUE!\",NEXT())")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(0, calls);
        Assert.Equal(CellValue.FromNumber(7), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(9), Value(result, snapshot, 0, 1));
        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 0, 2));
    }

    [Fact]
    public void ScriptFunctionOverridesBuiltinWithoutBuiltinArgumentOrLazyRules()
    {
        var names = new List<string>();
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            FunctionNames = ["SUM", "IF"],
            OnInvoke = (name, arguments) => { names.Add(name); return name == "SUM" ? "#VALUE!" : (double)arguments.Count; }
        });
        var snapshot = Snapshot((0, 0, Formula("=SUM()")), (1, 0, Formula("=IF(1)")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(new[] { "SUM", "IF" }, names);
        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(1), Value(result, snapshot, 1, 0));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteInjectedRuntimeResultsCannotEnterFreshCaches(double value)
    {
        var factory = new RecordingFactory(() => new RecordingRuntime { OnInvoke = (_, _) => value });
        var snapshot = Snapshot((0, 0, Formula("=NEXT()")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(CellValue.FromError(CellErrorCode.Number), Value(result, snapshot, 0, 0));
        Assert.Equal(CalculationState.Error, Cell(result, snapshot, 0, 0).State);
    }

    [Fact]
    public void ScriptResultsPreserveStringsBooleansAndNull()
    {
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, arguments) => arguments.Count == 0 ? null : arguments[0]
        });
        var snapshot = Snapshot((0, 0, Formula("=NEXT(\"#VALUE!\")")),
            (1, 0, Formula("=NEXT(TRUE)")), (2, 0, Formula("=NEXT()")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(CellValue.FromText("#VALUE!"), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromBoolean(true), Value(result, snapshot, 1, 0));
        Assert.Equal(CellValue.Blank, Value(result, snapshot, 2, 0));
    }

    [Fact]
    public void ScriptReceivesRectangularRangeInRowColumnOrderWithTypedValuesAndEmptyPositions()
    {
        IReadOnlyList<object>? seen = null;
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, arguments) => { seen = arguments; return 1d; }
        });
        var snapshot = Snapshot((4, 0, Literal(CellValue.FromBoolean(true))),
            (0, 1, Literal(CellValue.FromText("00123"))), (0, 0, Literal(CellValue.FromNumber(2))),
            (0, 2, Formula("=NEXT(A1:B5)")));

        new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        var values = Assert.IsType<List<object>>(Assert.Single(seen!));
        Assert.Equal(new object[] { 2d, "00123", "", "", "", "", "", "", true, "" }, values);
    }

    [Fact]
    public void ScriptRangePreservesMiddleBlankAndBoundsWholeColumnAtUsedRows()
    {
        var received = new List<List<object>>();
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, arguments) => { received.Add(Assert.IsType<List<object>>(arguments[0])); return 1d; }
        });
        var snapshot = Snapshot((0, 0, Literal(CellValue.FromNumber(1))), (2, 0, Literal(CellValue.FromNumber(3))),
            (0, 1, Formula("=NEXT(A1:A3)")), (1, 1, Formula("=NEXT(A)")));

        new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(2, received.Count);
        Assert.All(received, values => Assert.Equal(new object[] { 1d, "", 3d }, values));
    }

    [Fact]
    public void OversizeScriptRangeFailsBeforeInvokingRuntime()
    {
        var calls = 0;
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, _) => { calls++; return 1d; }
        });
        var snapshot = Snapshot((0, 20, Formula("=NEXT(A1:K100000)")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(0, calls);
        Assert.Equal(CalculationState.Error, Cell(result, snapshot, 0, 20).State);
        Assert.Contains(result.Diagnostics, message => message.Contains("范围参数最多支持"));
    }

    [Fact]
    public void ScriptFailureProducesTypedErrorAndAddressedDiagnostic()
    {
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, _) => throw new InvalidOperationException("script failure")
        });
        var snapshot = Snapshot((0, 0, Formula("=NEXT()")), (0, 1, Formula("=A1+1")));

        var result = new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true));

        Assert.Equal(CellValue.FromError(CellErrorCode.Script), Value(result, snapshot, 0, 0));
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), Value(result, snapshot, 0, 1));
        Assert.Contains(result.Diagnostics, message => message.Contains("Sheet1!A1") && message.Contains("script failure"));
        Assert.True(factory.Created[0].Disposed);
    }

    [Fact]
    public void CompilationFailureEscapesWithoutChangingCachedSnapshot()
    {
        var factory = new RecordingFactory(() => throw new InvalidOperationException("compile failure"));
        var snapshot = Snapshot((0, 0, Formula("=NEXT()", CellValue.FromNumber(77))));

        var error = Assert.Throws<InvalidOperationException>(() => new LegacyWorkbookCalculator(factory).Calculate(Request(snapshot, true)));

        Assert.Equal("compile failure", error.Message);
        Assert.Equal(CellValue.FromNumber(77), snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void CancellationBeforeStartDoesNotCreateRuntime()
    {
        var factory = new RecordingFactory();
        var snapshot = Snapshot((0, 0, Formula("=NEXT()")));

        Assert.ThrowsAny<OperationCanceledException>(() => new LegacyWorkbookCalculator(factory)
            .Calculate(Request(snapshot, true), new CancellationToken(true)));

        Assert.Empty(factory.Created);
    }

    [Fact]
    public void CancellationDuringInvocationDisposesRuntimeAndReturnsNoPartialResults()
    {
        using var cancellation = new CancellationTokenSource();
        var factory = new RecordingFactory(() => new RecordingRuntime
        {
            OnInvoke = (_, _) => { cancellation.Cancel(); return 99d; }
        });
        var snapshot = Snapshot((0, 0, Formula("=1+1", CellValue.FromNumber(12))),
            (0, 1, Formula("=NEXT()", CellValue.FromNumber(13))));

        Assert.ThrowsAny<OperationCanceledException>(() => new LegacyWorkbookCalculator(factory)
            .Calculate(Request(snapshot, true), cancellation.Token));

        Assert.True(factory.Created[0].Disposed);
        Assert.Equal(CellValue.FromNumber(12), snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(13), snapshot.Sheets[0].GetCell(0, 1)!.Value);
    }

    [Fact]
    public void FormulaNumbersUseInvariantCultureAndRestoreCallingCulture()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            var culture = CultureInfo.GetCultureInfo("fr-FR");
            CultureInfo.CurrentCulture = culture;
            var snapshot = Snapshot((0, 0, Formula("=1.25+2.5")));

            var result = Calculate(snapshot);

            Assert.Equal(CellValue.FromNumber(3.75), Value(result, snapshot, 0, 0));
            Assert.Same(culture, CultureInfo.CurrentCulture);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void LookupsUseRequestedRangeAndReturnTypedValues()
    {
        var snapshot = Snapshot((0, 0, Literal(CellValue.FromText("key"))), (0, 1, Literal(CellValue.FromNumber(999))),
            (3, 2, Literal(CellValue.FromText("key"))), (3, 3, Literal(CellValue.FromText("00123"))),
            (4, 2, Literal(CellValue.FromText("next"))), (4, 3, Literal(CellValue.FromBoolean(true))),
            (0, 5, Formula("=VLOOKUP(\"key\",C4:D5,2,FALSE)")),
            (1, 5, Formula("=INDEX(C4:D5,2,2)")), (2, 5, Formula("=MATCH(\"next\",C4:C5,0)")));

        var result = Calculate(snapshot);

        Assert.Equal(CellValue.FromText("00123"), Value(result, snapshot, 0, 5));
        Assert.Equal(CellValue.FromBoolean(true), Value(result, snapshot, 1, 5));
        Assert.Equal(CellValue.FromNumber(2), Value(result, snapshot, 2, 5));
    }

    [Fact]
    public void NativeAdapterStartsCleanJavaScriptContextForEveryRound()
    {
        var snapshot = Snapshot((0, 0, Formula("=NEXT()"))).WithScriptSource(
            "let count = 0; quickExcel.functions.register('NEXT', () => ++count);");
        var calculator = new LegacyWorkbookCalculator(new QuickJsRuntimeFactory());

        Assert.Equal(CellValue.FromNumber(1), Value(calculator.Calculate(Request(snapshot, true)), snapshot, 0, 0));
        Assert.Equal(CellValue.FromNumber(1), Value(calculator.Calculate(Request(snapshot, true)), snapshot, 0, 0));
    }

    private static CalculationResult Calculate(WorkbookSnapshot snapshot) =>
        new LegacyWorkbookCalculator(new RecordingFactory()).Calculate(Request(snapshot));

    private static CalculationRequest Request(WorkbookSnapshot snapshot, bool scriptsEnabled = false) =>
        new(new CalculationStamp(Guid.NewGuid(), snapshot.Revision, 7, 3), snapshot, scriptsEnabled);

    private static WorkbookSnapshot Snapshot(params (int Row, int Column, CellSnapshot Cell)[] cells) =>
        new(Guid.NewGuid(), 11, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "Sheet1", cells.Select(cell =>
                new KeyValuePair<CellPosition, CellSnapshot>(new(cell.Row, cell.Column), cell.Cell)))], "source");

    private static CellSnapshot Literal(CellValue value) => new(CellContent.FromLiteral(value));
    private static CellSnapshot Formula(string formula, CellValue? cache = null) => new(CellContent.FromFormula(formula),
        cache is null ? null : new CellCalculation(cache, CalculationState.Stale, 0, 0, 0));
    private static CellCalculation Cell(CalculationResult result, WorkbookSnapshot snapshot, int row, int column) =>
        result.Cells[new(snapshot.Sheets[0].Id, row, column)];
    private static CellValue? Value(CalculationResult result, WorkbookSnapshot snapshot, int row, int column) =>
        Cell(result, snapshot, row, column).Value;

    private sealed class RecordingFactory(Func<RecordingRuntime>? create = null) : IScriptRuntimeFactory
    {
        public List<RecordingRuntime> Created { get; } = [];
        public List<string> Sources { get; } = [];
        public IScriptRuntime Create(string source)
        {
            Sources.Add(source);
            var runtime = create?.Invoke() ?? new RecordingRuntime();
            Created.Add(runtime);
            return runtime;
        }
    }

    private sealed class RecordingRuntime : IScriptRuntime
    {
        private int _count;
        public Func<string, IReadOnlyList<object>, object?>? OnInvoke { get; init; }
        public bool Disposed { get; private set; }
        public IReadOnlyList<string> FunctionNames { get; init; } = ["NEXT"];
        public object? Invoke(string name, IReadOnlyList<object> arguments) => OnInvoke is null ? (double)++_count : OnInvoke(name, arguments);
        public IReadOnlyList<string> GetLogs() => [];
        public void Dispose() => Disposed = true;
    }
}

