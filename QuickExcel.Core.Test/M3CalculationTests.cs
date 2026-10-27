using System.Globalization;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Runtime;

namespace QuickExcel.Core.Test;

public sealed class M3CalculationTests
{
    internal static CellSnapshot F(string text, double? cache = null) => new(CellContent.FromFormula(text),
        cache is { } n ? new(CellValue.FromNumber(n), CalculationState.Stale, 0, 0, 0) : null);
    internal static CellSnapshot N(double value) => new(CellContent.FromLiteral(CellValue.FromNumber(value)));
    internal static CellSnapshot T(string value) => new(CellContent.FromLiteral(CellValue.FromText(value)));
    internal static WorksheetSnapshot Sheet(string name, params (int Row, int Col, CellSnapshot Cell)[] cells) =>
        new(SheetId.New(), name, cells.Select(c => new KeyValuePair<CellPosition, CellSnapshot>(new(c.Row, c.Col), c.Cell)));
    internal static WorkbookSnapshot Book(params WorksheetSnapshot[] sheets) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, sheets);
    internal static CalculationRequest Request(WorkbookSnapshot book, bool scripts = false) => new(new(Guid.NewGuid(), book.Revision, 1, 0), book, scripts);
    internal static CellCalculation C(CalculationResult result, WorksheetSnapshot sheet, int row = 0, int column = 0) => result.Cells[new(sheet.Id, row, column)];
    internal sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Scripts must not run."); }
    private static CalculationResult Calculate(WorkbookSnapshot book) => new WorkbookCalculator(new NoScripts()).Calculate(Request(book));

    [Fact]
    public void LegacyBareColumnSpellingRemainsAvailableAndNamesTakePriority()
    {
        var sheet = Sheet("S", (0, 0, N(2)), (2, 0, N(4)), (0, 1, F("=SUM(A)")), (1, 1, F("=SUM(A:A)")));
        var book = Book(sheet);
        Assert.Equal(CellValue.FromNumber(6), C(Calculate(book), sheet, 0, 1).Value);
        Assert.Equal(CellValue.FromNumber(9), C(Calculate(book.WithDefinedNames([new("A", "=9")])), sheet, 0, 1).Value);
        Assert.Equal(CellValue.FromNumber(6), C(Calculate(book.WithDefinedNames([new("A", "=9")])), sheet, 1, 1).Value);
    }

    [Theory]
    [InlineData("=INDEX(A1:A2,B1:B2)")]
    [InlineData("=VLOOKUP(1,A1:B2,C1:C2)")]
    [InlineData("=MATCH(1,A1:A2,B1:B2)")]
    public void RangeUsedInScalarLookupArgumentIsCellErrorNotFailedRound(string formula)
    {
        var sheet = Sheet("S", (5, 0, F(formula)), (5, 1, F("=7")));
        var result = Calculate(Book(sheet));
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), C(result, sheet, 5).Value);
        Assert.Equal(CellValue.FromNumber(7), C(result, sheet, 5, 1).Value);
    }

    [Fact]
    public void RecursiveAndExpandingNamesFailWithinBoundedDiagnosticBudget()
    {
        var sheet = Sheet("S", (0, 0, F("=NameA")));
        var recursive = Book(sheet).WithDefinedNames([new("NameA", "=NameB"), new("NameB", "=NameA")]);
        Assert.Equal(CalculationState.Error, C(Calculate(recursive), sheet).State);
        var names = Enumerable.Range(0, 15).Select(i => new DefinedName("Name_" + i, i == 14 ? "=1" : $"=Name_{i + 1}+Name_{i + 1}"));
        var exponential = Book(Sheet("S", (0, 0, F("=Name_0")))).WithDefinedNames(names);
        var result = Calculate(exponential);
        Assert.Equal(CalculationState.Error, C(result, exponential.Sheets[0]).State);
        Assert.Contains(result.Diagnostics, d => d.Contains("16384"));
    }

    [Fact]
    public void CrossSheetChainQuotedUnicodeEscapedNamesMixedReferencesAndRanges()
    {
        var input = Sheet("O'Brien 数据", (0, 0, N(3)), (1, 0, N(4)));
        var subtotal = Sheet("Sheet2", (0, 0, F("=SUM('O''Brien 数据'!$A$1:$A2)")), (0, 1, F("='O''Brien 数据'!A1")));
        var output = Sheet("结果", (0, 0, F("=Sheet2!A1+Sheet2!B1")), (0, 1, F("=SUM('O''Brien 数据'!A:A)")));
        var result = Calculate(Book(output, subtotal, input));
        Assert.Equal(CellValue.FromNumber(10), C(result, output).Value);
        Assert.Equal(CellValue.FromNumber(7), C(result, output, 0, 1).Value);
        Assert.Equal(CellValue.FromNumber(3), C(result, subtotal, 0, 1).Value);
    }

    [Theory]
    [InlineData("=SUM(1,,2)")]
    [InlineData("=1 trailing")]
    [InlineData("=1+")]
    [InlineData("=\"unclosed")]
    [InlineData("=A0")]
    [InlineData("=A1:B")]
    [InlineData("=A2:A1")]
    [InlineData("=[book.xlsx]Sheet!A1")]
    [InlineData("='Sheet!A1")]
    public void UnsupportedOrMalformedSyntaxNeverSilentlyDropsTokens(string formula)
    {
        var sheet = Sheet("S", (0, 2, F(formula)), (1, 2, F("=6*7")));
        var result = Calculate(Book(sheet));
        Assert.Equal(CalculationState.Error, C(result, sheet, 0, 2).State);
        Assert.Equal(CellValue.FromNumber(42), C(result, sheet, 1, 2).Value);
        Assert.Contains(result.Diagnostics, d => d.Contains("S!C1"));
    }

    [Theory]
    [InlineData("=-2^2", "4")]
    [InlineData("=2^3^2", "64")]
    [InlineData("=-(2^2)", "-4")]
    [InlineData("=1E2+2.5e-1", "100.25")]
    [InlineData("=ROUND(2.5,0)", "3")]
    [InlineData("=ROUND(-2.5,0)", "-3")]
    [InlineData("=ROUND(125,-1)", "130")]
    [InlineData("=IF(FALSE,1/0,42)", "42")]
    [InlineData("=IFERROR(\"oops\"+1,7)", "7")]
    [InlineData("=1+2&3", "33")]
    [InlineData("=IF(1&2=12,0,8)", "8")]
    public void TypedOperatorsAndLazyErrors(string formula, string expected)
    {
        var sheet = Sheet("S", (0, 0, F(formula)));
        Assert.Equal(expected, CellInput.DisplayText(C(Calculate(Book(sheet)), sheet).Value!.Value));
    }

    [Fact]
    public void ReferenceAggregatesDoNotGuessNumericTextAndCountAIncludesEmptyText()
    {
        var sheet = Sheet("S", (0, 0, N(3)), (1, 0, T("00123")), (2, 0, T("")),
            (3, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (0, 1, F("=SUM(A1:A4)")), (1, 1, F("=COUNT(A1:A4)")), (2, 1, F("=COUNTA(A1:A5)")),
            (3, 1, F("=SUM(\"2\",TRUE,A2)")), (4, 1, F("=AVERAGE(C1:C3)")), (5, 1, F("=A2=123")));
        var result = Calculate(Book(sheet));
        Assert.Equal(CellValue.FromNumber(3), C(result, sheet, 0, 1).Value);
        Assert.Equal(CellValue.FromNumber(1), C(result, sheet, 1, 1).Value);
        Assert.Equal(CellValue.FromNumber(4), C(result, sheet, 2, 1).Value);
        Assert.Equal(CellValue.FromNumber(3), C(result, sheet, 3, 1).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), C(result, sheet, 4, 1).Value);
        Assert.Equal(CellValue.FromBoolean(false), C(result, sheet, 5, 1).Value);
    }

    [Fact]
    public void TypedReferencesPreserveTextErrorsAndBlankWithoutReinterpretation()
    {
        var sheet = Sheet("S", (0, 0, T("#DIV/0!")), (1, 0, T("")),
            (0, 1, F("=IFERROR(A1,9)")), (1, 1, F("=A2")), (2, 1, F("=A3")), (3, 1, F("=#DIV/0!")),
            (4, 1, F("=\"a\"\"b\"")), (5, 1, F("=IFERROR(#REF!,8)")));
        var result = Calculate(Book(sheet));
        Assert.Equal(CellValue.FromText("#DIV/0!"), C(result, sheet, 0, 1).Value);
        Assert.Equal(CellValue.FromText(""), C(result, sheet, 1, 1).Value);
        Assert.Equal(CellValue.Blank, C(result, sheet, 2, 1).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), C(result, sheet, 3, 1).Value);
        Assert.Equal(CellValue.FromText("a\"b"), C(result, sheet, 4, 1).Value);
        Assert.Equal(CellValue.FromNumber(8), C(result, sheet, 5, 1).Value);
    }

    [Fact]
    public void NamedRangesResolveLocalBeforeWorkbookAndTrackCrossSheetDependencies()
    {
        var data = Sheet("Data", (0, 0, N(5)), (1, 0, N(7)));
        var a = Sheet("A", (0, 0, F("=SUM(Rates)")), (0, 1, F("=B!Rate")));
        var b = Sheet("B", (0, 0, F("=Rate")));
        var book = Book(a, b, data).WithDefinedNames([
            new("Rates", "='Data'!$A$1:$A$2"), new("Rate", "=2"), new("Rate", "=3", b.Id)]);
        var result = Calculate(book);
        Assert.Equal(CellValue.FromNumber(12), C(result, a).Value);
        Assert.Equal(CellValue.FromNumber(3), C(result, a, 0, 1).Value);
        Assert.Equal(CellValue.FromNumber(3), C(result, b).Value);
    }

    [Fact]
    public void StronglyConnectedMembersAndDownstreamAreDiagnosedSeparately()
    {
        var a = Sheet("A", (0, 0, F("=B!A1")), (0, 1, F("=7")), (0, 2, F("=C1")));
        var b = Sheet("B", (0, 0, F("=A!A1")), (0, 1, F("=A1+1")));
        var result = Calculate(Book(a, b));
        Assert.Contains(C(result, a).Diagnostics, d => d.Contains("属于循环强连通分量"));
        Assert.Contains(C(result, b, 0, 1).Diagnostics, d => d.Contains("不属于循环分量"));
        Assert.Contains(C(result, a, 0, 2).Diagnostics, d => d.Contains("属于循环强连通分量"));
        Assert.Equal(CellValue.FromNumber(7), C(result, a, 0, 1).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Cycle), C(result, b, 0, 1).Value);
    }

    [Fact]
    public void MissingSheetIsReferenceErrorAndUnknownNameIsNameError()
    {
        var sheet = Sheet("S", (0, 0, F("=Missing!A1")), (0, 1, F("=MissingName")));
        var result = Calculate(Book(sheet));
        Assert.Equal(CellValue.FromError(CellErrorCode.Reference), C(result, sheet).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Name), C(result, sheet, 0, 1).Value);
    }

    [Fact]
    public void IncrementalRecalculatesOnlyAffectedCrossSheetClosureIncludingNewBlankRangeInput()
    {
        var input = Sheet("Input", (0, 0, N(3)));
        var output = Sheet("Output", (0, 0, F("=SUM(Input!A1:A9)")), (0, 1, F("=A1+1")), (0, 2, F("=17")));
        using var session = new DocumentSession(Book(input, output));
        var calculator = new WorkbookCalculator(new NoScripts());
        Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        session.Execute(new SetCellsCommand([new(new(input.Id, 8, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        var result = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(2, result.Statistics!.EvaluatedCount);
        Assert.Equal(1, result.Statistics.ReusedCount);
        Assert.Equal(CellValue.FromNumber(13), C(result, output, 0, 1).Value);
        Assert.True(session.TryCommitCalculation(result));
        Assert.All(session.Snapshot.Sheets[1].Cells.Values, c => Assert.Equal(session.Revision, c.Calculation!.InputRevision));
    }

    [Fact]
    public void RandomEditsFormulaReplacementRemovalUndoAndRedoMatchFreshFullReference()
    {
        var data = Sheet("Data", Enumerable.Range(0, 30).Select(i => (i, 0, N(i))).ToArray());
        var output = Sheet("Output", Enumerable.Range(0, 30).Select(i =>
            (i, 0, F(i == 0 ? "=Data!A1*2" : $"=Data!A{i + 1}+A{i}"))).Append((0, 1, F("=SUM(Data!A:A)"))).ToArray());
        using var session = new DocumentSession(Book(data, output));
        var incremental = new WorkbookCalculator(new NoScripts());
        var random = new Random(42703);
        for (var round = 0; round < 100; round++)
        {
            var row = random.Next(30);
            if (round % 13 == 0 && session.CanUndo) session.Undo();
            else if (round % 17 == 0 && session.CanRedo) session.Redo();
            else
            {
                var content = round % 9 == 0 ? CellContent.FromLiteral(CellValue.Blank)
                    : round % 7 == 0 ? CellContent.FromFormula($"=Data!A{row + 1}*3")
                    : CellContent.FromLiteral(CellValue.FromNumber(random.Next(-100, 100)));
                session.Execute(new SetCellsCommand([new(new(round % 7 == 0 ? output.Id : data.Id, row, 0), content)]));
            }
            var request = session.BeginCalculation();
            var actual = incremental.Calculate(request);
            var reference = new WorkbookCalculator(new NoScripts()).Calculate(request with { ForceFull = true });
            Assert.Equal(reference.Cells.Count, actual.Cells.Count);
            foreach (var (address, cell) in reference.Cells)
            {
                Assert.Equal(cell.Value, actual.Cells[address].Value);
                Assert.Equal(cell.State, actual.Cells[address].State);
            }
            Assert.True(session.TryCommitCalculation(actual));
        }
    }

    [Fact]
    public void FormulaEdgeRemovalCycleRepairAndNewCycleMatchFullReference()
    {
        var sheet = Sheet("S", (0, 0, F("=B1")), (0, 1, F("=A1")), (0, 2, F("=A1+1")));
        using var session = new DocumentSession(Book(sheet));
        var calculator = new WorkbookCalculator(new NoScripts());
        session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation()));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 1), CellContent.FromLiteral(CellValue.FromNumber(8)))]));
        var result = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromNumber(9), C(result, sheet, 0, 2).Value);
        session.TryCommitCalculation(result);
        session.Undo();
        result = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromError(CellErrorCode.Cycle), C(result, sheet, 0, 2).Value);
    }

    [Fact]
    public void VolatileDateCellsUseWorkbookDateSystemAndNeverReuseDownstream()
    {
        var sheet = Sheet("S", (0, 0, F("=TODAY()")), (0, 1, F("=A1+1")), (0, 2, F("=9")));
        var book = Book(sheet);
        using var session = new DocumentSession(book);
        var calculator = new WorkbookCalculator(new NoScripts());
        session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation()));
        var result = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(2, result.Statistics!.EvaluatedCount);
        Assert.Equal(1, result.Statistics.ReusedCount);
        var mac = new WorkbookSnapshot(book.DocumentId, 0, ExcelDateSystem.Mac1904, book.Sheets);
        var macResult = Calculate(mac);
        Assert.Equal(1462, C(result, sheet).Value!.Value.NumberValue - C(macResult, sheet).Value!.Value.NumberValue);
    }

    [Fact]
    public void LongSparseDependencyChainDoesNotUseRecursiveGraphTraversal()
    {
        var sheet = Sheet("S", Enumerable.Range(0, 4000).Select(i => (i, 0, F(i == 3999 ? "=1" : $"=A{i + 2}+1"))).ToArray());
        Assert.Equal(CellValue.FromNumber(4000), C(Calculate(Book(sheet)), sheet).Value);
    }

    [Fact]
    public void CancellationDoesNotPublishOrSeedPartialCacheAndStampsRejectOldResult()
    {
        var sheet = Sheet("S", (0, 0, F("=1+1", 19)));
        var book = Book(sheet);
        using var session = new DocumentSession(book);
        var calculator = new WorkbookCalculator(new NoScripts());
        var request = session.BeginCalculation();
        Assert.ThrowsAny<OperationCanceledException>(() => calculator.Calculate(request, new CancellationToken(true)));
        Assert.Equal(CellValue.FromNumber(19), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        var old = calculator.Calculate(session.BeginCalculation());
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromFormula("=3"))]));
        Assert.False(session.TryCommitCalculation(old));
        var result = calculator.Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromNumber(3), C(result, sheet).Value);
    }

    [Fact]
    public void CulturesCannotChangeFormulaTokensOrTypedConversions()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var sheet = Sheet("S", (0, 0, F("=1.5+\"2.5\"")));
            Assert.Equal(CellValue.FromNumber(4), C(Calculate(Book(sheet)), sheet).Value);
            Assert.Equal("fr-FR", CultureInfo.CurrentCulture.Name);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void CrossSheetLookupsKeepOriginalTypes()
    {
        var data = Sheet("Data", (0, 0, N(1)), (1, 0, N(2)), (0, 1, T("00123")), (1, 1, T("next")));
        var sheet = Sheet("S", (0, 0, F("=VLOOKUP(1,Data!A1:B2,2,FALSE)")), (0, 1, F("=MATCH(2,Data!A1:A2,0)")),
            (0, 2, F("=INDEX(Data!A1:B2,2,2)")));
        var result = Calculate(Book(sheet, data));
        Assert.Equal(CellValue.FromText("00123"), C(result, sheet).Value);
        Assert.Equal(CellValue.FromNumber(2), C(result, sheet, 0, 1).Value);
        Assert.Equal(CellValue.FromText("next"), C(result, sheet, 0, 2).Value);
    }
}
