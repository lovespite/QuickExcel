using System.Collections.Immutable;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class OutlineSubtotalTests
{
    [Fact]
    public void CopySheetOwnsIndependentOutlineAndActiveFiltersCannotMixWithGrouping()
    {
        var sheet = Sheet("Data", (0, 0, T("Value")), (1, 0, N(1))); using var session = new DocumentSession(Book(sheet));
        session.Execute(new GroupAxisCommand(sheet.Id, SheetAxis.Row, 1, 3));
        var copy = new CopySheetCommand(sheet.Id, "Copy", 1); session.Execute(copy);
        session.Execute(new SetOutlineCollapsedCommand(copy.Id, SheetAxis.Row, 1, 3, 1, true));
        Assert.False(session.Snapshot.GetSheet(sheet.Id).Rows[1].Hidden); Assert.True(session.Snapshot.GetSheet(copy.Id).Rows[1].Hidden);
        var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new SetAutoFilterCommand(sheet.Id, new(new(0, 0, 1, 0), [new(0, new(values: ["1"]))]))));
        Assert.Same(before, session.Snapshot);
    }
    [Fact]
    public void StructuredTableTotalUsesTheFilteredDataRange()
    {
        var sheet = Sheet("Data", (0, 0, T("Value")), (0, 1, T("Keep")), (1, 0, N(10)), (1, 1, T("yes")), (2, 0, N(20)), (2, 1, T("no")),
            (3, 0, F("=SUBTOTAL(109,Sales[Value])")));
        sheet = sheet.WithData(new() { Tables = [new(1, "Sales", new(0, 0, 3, 1), [new(1, "Value", "sum"), new(2, "Keep")], totalsRow: true,
            filter: new(new(0, 0, 2, 1), [new(1, new(values: ["yes"]))]))] });
        using var session = new DocumentSession(Book(sheet)); Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal(10, session.Snapshot.Sheets[0].GetCell(3, 0)!.Value.NumberValue);
    }
    [Fact]
    public void NestedGroupsRetainInnerCollapseWhenOuterGroupExpandsAndUndoRestoresMetadata()
    {
        var sheet = Sheet("Data", (0, 0, N(1))); using var session = new DocumentSession(Book(sheet));
        session.Execute(new GroupAxisCommand(sheet.Id, SheetAxis.Row, 2, 3));
        session.Execute(new GroupAxisCommand(sheet.Id, SheetAxis.Row, 1, 11));
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 2, 4, 2, true));
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 1, 11, 1, true));
        Assert.True(session.Snapshot.Sheets[0].Rows[1].Hidden);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 1, 1, hidden: false)));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 2, 4, 2, false)));
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 1, 11, 1, false));
        Assert.False(session.Snapshot.Sheets[0].Rows[1].Hidden); Assert.True(session.Snapshot.Sheets[0].Rows[3].Hidden);
        session.Undo(); Assert.True(session.Snapshot.Sheets[0].Rows[1].Hidden); session.Redo();
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 2, 1)));
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 2, 4, 2, false));
        session.Execute(new GroupAxisCommand(sheet.Id, SheetAxis.Row, 1, 11, true));
        Assert.Equal(1, session.Snapshot.Sheets[0].Rows[3].OutlineLevel);
    }
    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void OutlineDepthAndBoundsRejectAtomically(SheetAxis axis)
    {
        var sheet = Sheet("Data"); using var session = new DocumentSession(Book(sheet));
        for (var i = 0; i < 7; i++) session.Execute(new GroupAxisCommand(sheet.Id, axis, 1, 3));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new GroupAxisCommand(sheet.Id, axis, 1, 3))); Assert.Same(before, session.Snapshot);
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, axis, 1, 3, 1, true));
        session.Execute(new SetOutlineCollapsedCommand(sheet.Id, axis, 1, 3, 1, false));
        Assert.False(axis == SheetAxis.Row ? session.Snapshot.Sheets[0].Rows[2].Hidden : session.Snapshot.Sheets[0].Columns[2].Hidden);
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void OutlineFlagsAboveLeftAndEmptyMetadataSurviveOrdinarySave(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-outline-" + Guid.NewGuid().ToString("N") + "." + extension);
        try
        {
            var sheet = Sheet("Data").WithData(new() { Outline = new(false, false, false, true) }); using var session = new DocumentSession(Book(sheet));
            session.Execute(new GroupAxisCommand(sheet.Id, SheetAxis.Column, 2, 3)); session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Column, 2, 4, 1, true));
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var result = storage.Open(path).Sheets[0];
            Assert.Equal(sheet.Data.Outline, result.Data.Outline); Assert.True(result.Columns[1].Collapsed); Assert.True(result.Columns[3].Hidden); Assert.Equal(1, result.Columns[3].OutlineLevel);
        }
        finally { File.Delete(path); }
    }
    [Fact]
    public void VisibilityAndFilterFormulaDependenciesRecalculateIncrementallyAndInvalidateDependents()
    {
        var sheet = Sheet("Data", (0, 0, T("Value")), (0, 1, T("Keep")), (1, 0, N(10)), (2, 0, N(20)), (1, 1, F("=D1")), (2, 1, N(1)), (0, 3, N(1)),
            (0, 5, F("=SUBTOTAL(9,A2:A3)")), (1, 5, F("=SUBTOTAL(109,A2:A3)")), (2, 5, F("=F2+1")));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        double Value(int row) => session.Snapshot.Sheets[0].GetCell(row, 5)!.Value.NumberValue;
        Calculate(); Assert.Equal(30, Value(0));
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 1, 1, hidden: true)); Calculate();
        Assert.Equal(30, Value(0)); Assert.Equal(20, Value(1)); Assert.Equal(21, Value(2));
        var result = session.Execute(new SetAutoFilterCommand(sheet.Id, new(new(0, 0, 2, 1), [new(1, new(values: ["1"]))])));
        Assert.False(result.CalculationUnchanged); Calculate();
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 3), CellContent.FromLiteral(CellValue.FromNumber(0)))])); Calculate();
        Assert.Equal(20, Value(0)); Assert.Equal(20, Value(1));
        session.Undo(); Calculate(); Assert.Equal(30, Value(0));
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Column, 0, 1, hidden: true)); Calculate(); Assert.Equal(30, Value(0));
    }
    [Fact]
    public void NestedNameSubtotalsAreSkippedAndOverlappingReferencesCountRepeatedly()
    {
        var sheet = Sheet("Data", (0, 0, N(10)), (1, 0, N(20)), (2, 0, F("=MySubtotal")), (0, 2, F("=SUBTOTAL(9,A1:A3,A1)")));
        var book = Book(sheet).WithDefinedNames([new("MySubtotal", "=SUBTOTAL(9,Data!$A$1:$A$2)")]); using var session = new DocumentSession(book);
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal(40, session.Snapshot.Sheets[0].GetCell(0, 2)!.Value.NumberValue);
    }
    [Fact]
    public void GeneratorUsesContiguousGroupsMovesOutsideReferencesAndCommitsOneUndoableTransaction()
    {
        var sheet = Sheet("Data", (0, 0, T("Group")), (0, 1, T("Value")), (1, 0, T("a")), (2, 0, T("A")), (3, 0, T("b")),
            (1, 1, N(1)), (2, 1, N(2)), (3, 1, N(3)), (0, 4, F("=B4")), (5, 0, T("tail")));
        using var session = new DocumentSession(Book(sheet)); var before = session.Snapshot;
        var preview = SubtotalCommands.Preview(before, sheet.Id, new(0, 0, 3, 1), 0, [1]);
        Assert.Equal(3, preview.InsertedRows); Assert.Equal(2, preview.Groups.Length); Assert.Same(before, session.Snapshot);
        session.Execute(new ApplySubtotalsCommand(preview));
        Assert.Equal("=B5", session.Snapshot.Sheets[0].GetCell(0, 4)!.Content.FormulaText);
        Assert.Equal("=SUBTOTAL(9,B2:B5)", session.Snapshot.Sheets[0].GetCell(6, 1)!.Content.FormulaText);
        Assert.Equal("tail", session.Snapshot.Sheets[0].GetCell(8, 0)!.Value.TextValue);
        session.Undo(); Assert.Equal("=B4", session.Snapshot.Sheets[0].GetCell(0, 4)!.Content.FormulaText);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplySubtotalsCommand(preview)));
        session.Redo(); Assert.Equal(2, session.Snapshot.Sheets[0].Rows[1].OutlineLevel);
    }
    [Fact]
    public void InvalidGeneratorAndUnmodeledFiltersCannotProducePartialOrFalseFreshResults()
    {
        var sheet = Sheet("Data", (0, 0, T("Key")), (1, 0, T("a")), (1, 1, N(10)), (0, 3, F("=SUBTOTAL(9,B2:B3)")));
        var book = Book(sheet);
        Assert.Throws<NotSupportedException>(() => SubtotalCommands.Preview(book, sheet.Id, new(0, 0, 1, 1), 0, [1]));
        var state = new WorkbookFileState(Path.Combine(Path.GetTempPath(), "unknown-filter.xlsx"), 2, [], book, [], [], compatibilityDiagnostics:
            [new("unknown-filter", CompatibilityKind.PreservedOnly, "worksheet/autoFilter", "Unknown filter")]);
        var blocked = new WorkbookSnapshot(book.DocumentId, 0, book.DateSystem, [sheet], fileState: state);
        using var session = new DocumentSession(blocked); Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal(CalculationState.Blocked, session.Snapshot.Sheets[0].GetCell(0, 3)!.Calculation!.State);
    }
}
