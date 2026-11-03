using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core.Test;

public sealed class DynamicArrayTests
{
    [Fact]
    public void ArrayEvaluationMatchesIndependentExcelSpillCells()
    {
        using var spec = ScalarFormulaFixture.Load("dynamic-array-spec");
        using var oracle = ScalarFormulaFixture.Load("excel-dynamic-arrays");
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", spec.RootElement.GetProperty("Inputs").EnumerateArray().Select(input =>
            KeyValuePair.Create(ScalarFormulaFixture.Position(input.GetProperty("Address").GetString()!), new CellSnapshot(ScalarFormulaFixture.Content(input)))));
        var resultSheet = new WorksheetSnapshot(SheetId.New(), "Results");
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet, resultSheet]);
        var scalarResults = new WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(M3CalculationTests.Request(book));
        var evaluator = new TypedEvaluator(book, scalarResults.Cells, null, new Dictionary<string, ScriptFunctionDescriptor>(), () => { })
            { ArrayMode = true, CurrentSheet = resultSheet.Id };
        var expected = oracle.RootElement.GetProperty("Initial").EnumerateArray().ToDictionary(e => e.GetProperty("Id").GetString()!);
        var differences = new List<string>();
        foreach (var item in spec.RootElement.GetProperty("Arrays").EnumerateArray())
        {
            var formula = item.GetProperty("Formula").GetString()!; var id = item.GetProperty("Id").GetString()!;
            var node = FormulaSyntax.Map(FormulaSyntax.Parse(formula), n => n is ReferenceNode r ? r with { Sheet = r.Sheet ?? resultSheet.Name } : n);
            object value;
            try { value = evaluator.FormulaResult(node); }
            catch (FormulaValueException error) { value = CellValue.FromError(error.Code); }
            for (var row = 0; row < item.GetProperty("Rows").GetInt32(); row++)
            for (var column = 0; column < item.GetProperty("Columns").GetInt32(); column++)
            {
                var wanted = ScalarFormulaFixture.Value(expected[$"{id}-{row}-{column}"]);
                var actual = value is FormulaArray array ? row < array.Rows && column < array.Columns ? array[row,column] : CellValue.Blank
                    : row == 0 && column == 0 ? (CellValue)value : CellValue.Blank;
                if (wanted != actual) differences.Add($"{id} [{row},{column}] {formula}: expected {wanted}, actual {actual}");
            }
        }
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
    }

    [Fact]
    public void ArrayValuesOwnTheirDataAndRejectMismatchedOrUnboundedShapes()
    {
        var values = new[] { CellValue.FromNumber(1), CellValue.FromNumber(2) };
        var array = new FormulaArray(1, 2, values);
        values[0] = CellValue.FromNumber(99);
        Assert.Equal(CellValue.FromNumber(1), array[0, 0]);
        Assert.Throws<ArgumentException>(() => new FormulaArray(1, 1, values));
        Assert.Throws<ArgumentException>(() => new FormulaArray(2, 2, values));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaArray(0, 2, values));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaArray(int.MaxValue, int.MaxValue, values));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaArray(1, FormulaArray.MaxElements + 1, values));
        Assert.Throws<ArgumentException>(() => new FormulaArray(200, 1,
            Enumerable.Repeat(CellValue.FromText(new string('x', WorkbookLimits.MaxCellTextLength)), 200)));
    }

    [Fact]
    public void SequenceComposesWithArrayArithmeticAndScalarAggregates()
    {
        var result = Array("=SEQUENCE(2,3,10,-2)*2+1");
        Assert.Equal((2, 3), (result.Rows, result.Columns));
        Assert.Equal(new double[] { 21, 17, 13, 9, 5, 1 }, result.Values.Select(v => v.NumberValue));
        Assert.Equal(CellValue.FromNumber(21), Evaluate("=SUM(SEQUENCE(2,3))"));
        Assert.Equal(CellValue.FromError(CellErrorCode.Calculation), Evaluate("=SEQUENCE(0)"));
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), Evaluate("=SEQUENCE(-1)"));
        Assert.Equal(CellValue.FromError(CellErrorCode.Calculation), Evaluate("=SEQUENCE(100001)"));
        Assert.All(Array("=SEQUENCE(2)/0").Values, v => Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), v));
    }

    [Fact]
    public void FilterSupportsRowsColumnsCompositionAndEmptyResults()
    {
        Assert.Equal(new double[] { 3, 4, 5, 6 }, Array("=FILTER(SEQUENCE(3,2),SEQUENCE(3)>1)").Values.Select(v => v.NumberValue));
        Assert.Equal(new double[] { 2, 4, 6 }, Array("=FILTER(SEQUENCE(3,2),SEQUENCE(1,2)=2)").Values.Select(v => v.NumberValue));
        Assert.Equal(CellValue.FromError(CellErrorCode.Calculation), Evaluate("=FILTER(SEQUENCE(3),FALSE)"));
        Assert.Equal(CellValue.FromText("empty"), Evaluate("=FILTER(SEQUENCE(3),SEQUENCE(3)>10,\"empty\")"));
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), Evaluate("=FILTER(SEQUENCE(3,2),SEQUENCE(2))"));
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), Evaluate("=FILTER(SEQUENCE(3),SEQUENCE(3)/0)"));
    }

    [Fact]
    public void SortIsStableAndUniqueComparesEntireRowsIgnoringTextCase()
    {
        var cells = new (int Row, int Column, CellValue Value)[]
        {
            (0,0,CellValue.FromText("b")), (0,1,CellValue.FromNumber(2)),
            (1,0,CellValue.FromText("A")), (1,1,CellValue.FromNumber(1)),
            (2,0,CellValue.FromText("a")), (2,1,CellValue.FromNumber(1)),
            (3,0,CellValue.FromText("b")), (3,1,CellValue.FromNumber(3))
        };
        var sorted = Array("=SORT(A1:B4)", cells);
        Assert.Equal(new[] { "A", "a", "b", "b" }, Enumerable.Range(0, 4).Select(i => sorted[i,0].TextValue));
        Assert.Equal(new double[] { 1, 1, 2, 3 }, Enumerable.Range(0, 4).Select(i => sorted[i,1].NumberValue));
        var unique = Array("=UNIQUE(A1:B4)", cells);
        Assert.Equal(3, unique.Rows);
        Assert.Equal(new double[] { 2, 1, 3 }, Enumerable.Range(0, 3).Select(i => unique[i,1].NumberValue));
        Assert.Equal(new double[] { 2, 3 }, Array("=UNIQUE(A1:B4,FALSE,TRUE)", cells).Values.Where(v => v.Kind == CellValueKind.Number).Select(v => v.NumberValue));
        Assert.Equal(new double[] { 3, 2, 1, 6, 5, 4 }, Array("=SORT(SEQUENCE(2,3),1,-1,TRUE)").Values.Select(v => v.NumberValue));
    }

    [Fact]
    public void ArrayOperationsObserveCancellationInsideLoops()
    {
        var checks = 0;
        Assert.Throws<OperationCanceledException>(() => Evaluate("=SORT(SEQUENCE(10000,1,10000,-1))", [], () =>
        { if (++checks == 15000) throw new OperationCanceledException(); }));
    }

    [Fact]
    public void SpillCommitCreatesDerivedCellsAndRejectsPartialEditsAtomically()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [KeyValuePair.Create(new CellPosition(0, 0),
            new CellSnapshot(CellContent.FromFormula("=SEQUENCE(3)")))]);
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
        var owner = new CellAddress(sheet.Id, 0, 0);
        void Publish(int count)
        {
            var request = session.BeginCalculation();
            var array = new FormulaArray(count, 1, Enumerable.Range(1, count).Select(i => CellValue.FromNumber(i)));
            var calculation = new CellCalculation(array[0,0], CalculationState.Fresh, request.Stamp.InputRevision,
                request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration, array: array);
            Assert.True(session.TryCommitCalculation(new(request.Stamp, [KeyValuePair.Create(owner, calculation)])));
        }
        Publish(3);
        var child = session.Snapshot.GetSheet(sheet.Id).GetCell(2, 0)!;
        Assert.Equal(owner, child.SpillOwner);
        Assert.Equal(CellValue.Blank, child.Content.LiteralValue);
        Assert.Equal(CellValue.FromNumber(3), child.Value);
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([
            new(new(sheet.Id, 5, 0), CellContent.FromLiteral(CellValue.FromNumber(5))),
            new(new(sheet.Id, 2, 0), CellContent.FromLiteral(CellValue.FromNumber(99)))])));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ClearCellsCommand([new(sheet.Id, 2, 0)])));
        Assert.Same(before, session.Snapshot);
        Assert.False(session.IsDirty);
        Publish(2);
        Assert.Null(session.Snapshot.GetSheet(sheet.Id).GetCell(2, 0));
        session.Execute(new ClearCellsCommand([owner]));
        Assert.Null(session.Snapshot.GetSheet(sheet.Id).GetCell(1, 0));
        Assert.True(session.Undo());
        Assert.Equal(owner, session.Snapshot.GetSheet(sheet.Id).GetCell(1, 0)!.SpillOwner);
        Assert.Equal(CalculationState.Stale, session.Snapshot.GetSheet(sheet.Id).GetCell(1, 0)!.Calculation!.State);
    }

    [Fact]
    public void SpillAllocationChecksContentMergesEdgesAndCancellationBeforeAdmission()
    {
        var id = SheetId.New();
        WorkbookSnapshot Book(CellSnapshot? obstacle = null, string[]? merges = null) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(id, "Data", new[] { KeyValuePair.Create(new CellPosition(0,0), new CellSnapshot(CellContent.FromFormula("=SEQUENCE(3)"))) }
                .Concat(obstacle is null ? [] : new[] { KeyValuePair.Create(new CellPosition(2,0), obstacle) }), mergedCells: merges)]);
        var spill = new FormulaSpill(new(id,0,0), new(3,1,Enumerable.Range(1,3).Select(i => CellValue.FromNumber(i))));
        foreach (var book in new[] { Book(new(CellContent.FromLiteral(CellValue.FromText("")))), Book(merges: ["A2:B2"]) })
        {
            var allocator = new FormulaSpillAllocator(book);
            Assert.False(allocator.TryReserve(spill, out var reason));
            Assert.NotNull(reason); Assert.Empty(allocator.Owners); Assert.Empty(allocator.Spills);
        }
        var available = new FormulaSpillAllocator(Book()); var checks = 0;
        Assert.Throws<OperationCanceledException>(() => available.TryReserve(spill, out _, () =>
        { if (++checks == 2) throw new OperationCanceledException(); }));
        Assert.Empty(available.Owners);
        Assert.True(available.TryReserve(spill, out _));
        Assert.Equal(3, available.Owners.Count);
    }

    [Fact]
    public void InvalidSpillPublicationNeverChangesSession()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [
            KeyValuePair.Create(new CellPosition(0,0), new CellSnapshot(CellContent.FromFormula("=SEQUENCE(2)"))),
            KeyValuePair.Create(new CellPosition(1,0), new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("occupied"))))]);
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
        var before = session.Snapshot; var request = session.BeginCalculation();
        var array = new FormulaArray(2,1,[CellValue.FromNumber(1),CellValue.FromNumber(2)]);
        var calculation = new CellCalculation(array[0,0],CalculationState.Fresh,request.Stamp.InputRevision,
            request.Stamp.CalculationGeneration,request.Stamp.ScriptGeneration,array:array);
        Assert.Throws<InvalidOperationException>(() => session.TryCommitCalculation(new(request.Stamp,
            [KeyValuePair.Create(new CellAddress(sheet.Id,0,0),calculation)])));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void WorkbookRoundReplansConsumersOfSpillChildrenAndShrinksAfterInputChanges()
    {
        var sheet = M3CalculationTests.Sheet("Data",
            (0,0,M3CalculationTests.F("=B4+1")), (0,2,new CellSnapshot(CellContent.FromLiteral(CellValue.FromNumber(3)))),
            (1,0,M3CalculationTests.F("=SUM(B2#)")), (1,1,M3CalculationTests.F("=SEQUENCE(C1)")),
            (5,0,M3CalculationTests.F("=SUM(B3:B4)")));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet));
        var calculator = new WorkbookCalculator(new M3CalculationTests.NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        CellValue Value(int row, int column) => session.Snapshot.GetSheet(sheet.Id).GetCell(row,column)?.Value ?? CellValue.Blank;
        Calculate();
        Assert.Equal(CellValue.FromNumber(4), Value(0,0));
        Assert.Equal(CellValue.FromNumber(6), Value(1,0));
        Assert.Equal(CellValue.FromNumber(5), Value(5,0));
        session.Execute(new SetCellsCommand([new(new(sheet.Id,0,2),CellContent.FromLiteral(CellValue.FromNumber(1)))]));
        Calculate();
        Assert.Equal(CellValue.FromNumber(1), Value(0,0));
        Assert.Equal(CellValue.FromNumber(1), Value(1,0));
        Assert.Equal(CellValue.FromNumber(0), Value(5,0));
        Assert.Null(session.Snapshot.GetSheet(sheet.Id).GetCell(3,1));
        Assert.True(session.Undo()); Calculate(); Assert.Equal(CellValue.FromNumber(4), Value(0,0));
    }

    [Fact]
    public void SpillCyclesAndObstaclesDoNotPublishGhostValues()
    {
        var sheet = M3CalculationTests.Sheet("Data", (0,0,M3CalculationTests.F("=SEQUENCE(2)+A2")),
            (0,2,M3CalculationTests.F("=SEQUENCE(2)")), (1,2,new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("occupied")))),
            (3,0,M3CalculationTests.F("=A1+1")), (3,2,M3CalculationTests.F("=C1#")));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet));
        var result = new WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation());
        Assert.True(session.TryCommitCalculation(result));
        Assert.Equal(CellValue.FromError(CellErrorCode.Cycle), session.Snapshot.GetSheet(sheet.Id).GetCell(0,0)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Cycle), session.Snapshot.GetSheet(sheet.Id).GetCell(3,0)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill), session.Snapshot.GetSheet(sheet.Id).GetCell(0,2)!.Value);
        Assert.Equal(CellValue.FromText("occupied"), session.Snapshot.GetSheet(sheet.Id).GetCell(1,2)!.Value);
        Assert.Null(session.Snapshot.GetSheet(sheet.Id).GetCell(1,0));
    }

    [Theory]
    [InlineData("=SUM(A1#)")]
    [InlineData("='Data'!$B$2#")]
    public void SpillReferencesRoundTripAndTransformTheirAnchor(string source)
    {
        var parsed = FormulaSyntax.Parse(source);
        Assert.True(FormulaSyntax.TryPrintPortable(parsed, out var printed));
        Assert.Equal(FormulaSyntax.Print(parsed), FormulaSyntax.Print(FormulaSyntax.Parse(printed)));
        Assert.Contains(FormulaSyntax.Walk(parsed), n => n is UnaryNode { Operator: "#" });
    }

    private static FormulaArray Array(string formula, params (int Row, int Column, CellValue Value)[] cells) =>
        Assert.IsType<FormulaArray>(Evaluate(formula, cells));

    private static object Evaluate(string formula, (int Row, int Column, CellValue Value)[]? cells = null, Action? check = null)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", (cells ?? []).Select(c =>
            KeyValuePair.Create(new CellPosition(c.Row, c.Column), new CellSnapshot(CellContent.FromLiteral(c.Value)))));
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var evaluator = new TypedEvaluator(snapshot, new Dictionary<CellAddress, CellCalculation>(), null,
            new Dictionary<string, ScriptFunctionDescriptor>(), check ?? (() => { })) { ArrayMode = true, CurrentSheet = sheet.Id };
        var expression = FormulaSyntax.Map(FormulaSyntax.Parse(formula), n => n is ReferenceNode r ? r with { Sheet = sheet.Name } : n);
        return evaluator.FormulaResult(expression);
    }
}
