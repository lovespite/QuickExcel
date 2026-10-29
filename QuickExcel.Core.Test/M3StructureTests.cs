using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M3StructureTests
{
    private static CellValue Result(DocumentSession session, WorksheetSnapshot sheet, int row = 0, int col = 0)
    {
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        session.TryCommitCalculation(result);
        return C(result, sheet, row, col).Value!.Value;
    }

    [Fact]
    public void RenameUpdatesQuotedCrossSheetAndNamesWithoutTouchingStringLiteralsAndUndoRestores()
    {
        var data = Sheet("Data", (0, 0, N(7)));
        var report = Sheet("Report", (0, 0, F("=Data!$A1+Rate")), (1, 0, F("=\"Data!A1\"")));
        var book = Book(data, report).WithDefinedNames([new("Rate", "='Data'!$A$1")]);
        using var session = new DocumentSession(book);
        session.Execute(new RenameSheetCommand(data.Id, "O'Brien 数据"));
        Assert.Equal(data.Id, session.Snapshot.Sheets[0].Id);
        Assert.Equal(CellValue.FromNumber(14), Result(session, report));
        Assert.Equal("=\"Data!A1\"", session.Snapshot.Sheets[1].GetCell(1, 0)!.Content.FormulaText);
        Assert.Contains("'O''Brien 数据'!", session.Snapshot.DefinedNames[0].Formula);
        Assert.True(session.Undo());
        Assert.Equal("Data", session.Snapshot.Sheets[0].Name);
        Assert.Equal(book.Sheets[1].GetCell(0, 0)!.Content, session.Snapshot.Sheets[1].GetCell(0, 0)!.Content);
        Assert.True(session.Redo());
        Assert.Equal(CellValue.FromNumber(14), Result(session, report));
    }

    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void InsertDeleteAndUndoKeepAbsoluteMixedAndRelativeTargets(SheetAxis axis)
    {
        var data = Sheet("Data", (1, 1, N(7)));
        var output = Sheet("Report", (0, 0, F("=Data!$B$2+Data!B$2+Data!$B2+Data!B2")));
        using var session = new DocumentSession(Book(data, output));
        session.Execute(new ShiftAxisCommand(data.Id, axis, 0, 2));
        Assert.Equal(CellValue.FromNumber(28), Result(session, output));
        var refs = FormulaSyntax.Walk(FormulaSyntax.Parse(session.Snapshot.Sheets[1].GetCell(0, 0)!.Content.FormulaText)).OfType<ReferenceNode>().ToArray();
        Assert.True(refs[0].ColumnAbsolute && refs[0].RowAbsolute);
        Assert.False(refs[1].ColumnAbsolute); Assert.True(refs[1].RowAbsolute);
        Assert.True(refs[2].ColumnAbsolute); Assert.False(refs[2].RowAbsolute);
        session.Execute(new ShiftAxisCommand(data.Id, axis, 3, 1, delete: true));
        Assert.Equal(CellValue.FromError(CellErrorCode.Reference), Result(session, output));
        session.Undo();
        Assert.Equal(CellValue.FromNumber(28), Result(session, output));
        session.Undo();
        Assert.Equal(CellValue.FromNumber(28), Result(session, output));
    }

    [Fact]
    public void RangesExpandAndShrinkWithInsertDeleteAndNamesFollow()
    {
        var data = Sheet("Data", (0, 0, N(2)), (1, 0, N(3)), (2, 0, N(5)));
        var output = Sheet("R", (0, 0, F("=SUM(Values)")), (1, 0, F("=SUM(Data!A:A)")));
        using var session = new DocumentSession(Book(data, output).WithDefinedNames([new("Values", "='Data'!$A$1:$A$3")]));
        session.Execute(new ShiftAxisCommand(data.Id, SheetAxis.Row, 1, 1));
        session.Execute(new SetCellsCommand([new(new(data.Id, 1, 0), CellContent.FromLiteral(CellValue.FromNumber(7)))]));
        Assert.Equal(CellValue.FromNumber(17), Result(session, output));
        session.Execute(new ShiftAxisCommand(data.Id, SheetAxis.Row, 2, 1, delete: true));
        Assert.Equal(CellValue.FromNumber(14), Result(session, output));
        Assert.Equal(CellValue.FromNumber(14), Result(session, output, 1));
    }

    [Fact]
    public void MoveRowsKeepsReferencedCellsAndMovesMetadata()
    {
        var source = Sheet("Data", (0, 0, N(3)), (1, 0, N(5)), (2, 0, N(7)));
        source = new(source.Id, source.Name, source.Cells, rows: [new(0, new RowMetadata(31))]);
        var output = Sheet("R", (0, 0, F("=Data!A1")), (0, 1, F("=SUM(Data!A1:A3)")));
        using var session = new DocumentSession(Book(source, output));
        session.Execute(new MoveAxisCommand(source.Id, SheetAxis.Row, 0, 1, 2));
        Assert.Equal(CellValue.FromNumber(3), Result(session, output));
        Assert.Equal(CellValue.FromNumber(15), Result(session, output, 0, 1));
        Assert.Equal(31, session.Snapshot.Sheets[0].Rows[2].Height);
        Assert.Equal(CellValue.FromNumber(5), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void MovingPartialReferencedRangeFailsAtomically()
    {
        var data = Sheet("Data", (0, 0, N(1)), (1, 0, N(2)), (2, 0, N(3)));
        var output = Sheet("R", (0, 0, F("=SUM(Data!A1:A2)")));
        using var session = new DocumentSession(Book(data, output));
        var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveAxisCommand(data.Id, SheetAxis.Row, 0, 1, 2)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveCellsCommand(data.Id, new(0, 0), new(0, 0), data.Id, new(5, 0))));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void CrossSheetCutKeepsFormulaTargetsAndUpdatesReferencesToMovedCells()
    {
        var data = Sheet("Data", (0, 0, N(7)), (0, 1, F("=$A1*2")));
        var target = Sheet("Target", (0, 0, N(99)));
        var output = Sheet("R", (0, 0, F("=Data!B1")));
        using var session = new DocumentSession(Book(data, target, output));
        session.Execute(new MoveCellsCommand(data.Id, new(0, 1), new(0, 1), target.Id, new(2, 2)));
        Assert.Equal(CellValue.FromNumber(14), Result(session, output));
        Assert.Equal(CellValue.FromNumber(14), Result(session, target, 2, 2));
        Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 1));
        session.Undo();
        Assert.NotNull(session.Snapshot.Sheets[0].GetCell(0, 1));
        Assert.Null(session.Snapshot.Sheets[1].GetCell(2, 2));
    }

    [Fact]
    public void OverlappingMoveUsesCapturedSourceAndOverwritesBlankDestinationPositions()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (1, 0, N(2)), (3, 0, N(9)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new MoveCellsCommand(sheet.Id, new(0, 0), new(2, 0), sheet.Id, new(1, 0)));
        var current = session.Snapshot.Sheets[0];
        Assert.Null(current.GetCell(0, 0)); Assert.Null(current.GetCell(3, 0));
        Assert.Equal(CellValue.FromNumber(1), current.GetCell(1, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(2), current.GetCell(2, 0)!.Value);
    }

    [Fact]
    public void InvalidNamesOverflowAndUnsupportedStructureLeaveRevisionHistoryAndSnapshotIntact()
    {
        var data = Sheet("Data", (WorkbookLimits.MaxRows - 1, 0, N(1)));
        var output = Sheet("R");
        using var session = new DocumentSession(Book(data, output));
        var before = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new RenameSheetCommand(data.Id, "R")));
        Assert.Throws<ArgumentException>(() => session.Execute(new ShiftAxisCommand(data.Id, SheetAxis.Row, 0, 1)));
        Assert.Throws<ArgumentException>(() => session.Execute(new DefineNameCommand(new("Bad", "=Data!A1"))));
        Assert.Same(before, session.Snapshot);
        Assert.Equal(0, session.Revision); Assert.False(session.CanUndo);
        var merged = new WorksheetSnapshot(data.Id, data.Name, mergedCells: ["A1:B1"]);
        using var protectedSession = new DocumentSession(Book(merged));
        Assert.Throws<NotSupportedException>(() => protectedSession.Execute(new MoveCellsCommand(data.Id, new(0, 0), new(0, 1), data.Id, new(1, 0))));
    }
}
