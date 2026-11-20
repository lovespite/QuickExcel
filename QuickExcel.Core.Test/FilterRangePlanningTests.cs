using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class FilterRangePlanningTests
{
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));

    [Fact]
    public void TextHeaderSelectionIncludesCandidatesAndSortsDataWithOneUndoEntry()
    {
        var sheet = Sheet("Text records",
            (0, 0, T("X")), (0, 1, T("a")), (0, 2, T("v")), (0, 3, T("c")),
            (1, 0, T("S")), (1, 1, T("X")), (1, 2, T("X")), (1, 3, T("vv")),
            (2, 0, T("S")), (2, 1, T("X")), (2, 2, T("X")), (2, 3, T("f")),
            (3, 0, T("S")), (3, 1, T("X")), (3, 2, T("X")), (3, 3, T("g")),
            (4, 0, T("S")), (4, 1, T("X")), (4, 2, T("X")), (4, 3, T("g")),
            (5, 0, T("S")), (5, 1, T("X")), (5, 2, T("X")), (5, 3, T("g")));
        var range = PersistentFiltering.ResolveRange(sheet, new(0, 0, 0, 3), new(0, 3));
        Assert.Equal(new CellRectangle(0, 0, 5, 3), range);
        using var document = new WorkbookDocument(Book(sheet), Services);
        document.Session.Execute(new SetAutoFilterCommand(sheet.Id, new(range)));
        var candidates = PersistentFiltering.Candidates(document.Snapshot, document.Snapshot.Sheets[0], 3);
        Assert.Equal(new[] { "f", "g", "vv" }, candidates.Select(c => c.Label));
        Assert.Equal(new[] { 1, 3, 1 }, candidates.Select(c => c.Count));
        Assert.All(candidates, c => Assert.True(c.Selected));

        document.Session.Execute(new SortRangeCommand(new SortDescriptor(sheet.Id,
            new(range.Top, range.Left), new(range.Bottom, range.Right), [new(3)], hasHeader: true, expandColumns: true)));
        Assert.Equal(new[] { "c", "f", "g", "g", "g", "vv" }, Enumerable.Range(0, 6).Select(r => document.GetCellInput(0, r, 3)));
        Assert.Equal(new[] { "X", "a", "v", "c" }, Enumerable.Range(0, 4).Select(c => document.GetCellInput(0, 0, c)));
        document.Session.Undo();
        Assert.Equal(new[] { "c", "vv", "f", "g", "g", "g" }, Enumerable.Range(0, 6).Select(r => document.GetCellInput(0, r, 3)));
        Assert.Equal(range, document.Snapshot.Sheets[0].Data.Filter!.Range);
        Assert.Null(document.Snapshot.Sheets[0].Data.Sort);
    }

    [Fact]
    public void HeaderSelectionPreservesColumnsAndTopAndStopsAtACompletelyBlankRecord()
    {
        var sheet = Sheet("Bounds",
            (2, 1, T("above")), (3, 0, T("outside")), (3, 1, T("name")), (3, 2, T("value")), (3, 3, T("outside")),
            (4, 1, T("first")), (5, 2, T("second")),
            (6, 0, T("adjacent data is not a record")), (6, 3, T("adjacent data is not a record")),
            (7, 1, T("separate block")));
        Assert.Equal(new CellRectangle(3, 1, 5, 2), PersistentFiltering.ResolveRange(sheet, new(3, 1, 3, 2), new(3, 2)));
    }

    [Fact]
    public void BlankActiveHeaderDoesNotHideDataInOtherSelectedColumns()
    {
        var sheet = Sheet("Blank header", (0, 1, T("Name")), (1, 1, T("A")), (2, 0, T("B")));
        Assert.Equal(new CellRectangle(0, 0, 2, 1), PersistentFiltering.ResolveRange(sheet, new(0, 0, 0, 1), new(0, 0)));
    }

    [Fact]
    public void FormulaRecordsRemainConnectedButFormattingOnlyRowsStopExpansion()
    {
        var sheet = Sheet("Sparse", (0, 0, T("Name")), (0, 1, T("Value")),
            (1, 0, F("=\"\"")),
            (2, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.Blank, 1))),
            (3, 0, T("outside")));
        Assert.Equal(new CellRectangle(0, 0, 1, 1), PersistentFiltering.ResolveRange(sheet, new(0, 0, 0, 1), new(0, 1)));
    }

    [Fact]
    public void ExplicitMultiRowSelectionKeepsItsExactBounds()
    {
        var sheet = Sheet("Explicit", (0, 0, T("outside")), (0, 1, T("header")), (0, 2, T("outside")),
            (1, 1, T("selected")), (2, 1, T("outside")));
        var selection = new CellRectangle(0, 1, 1, 1);
        Assert.Equal(selection, PersistentFiltering.ResolveRange(sheet, selection, new(1, 1)));
    }

    [Fact]
    public void SingleCellAndWholeAxesRetainTheirConnectedRegionRules()
    {
        var sheet = Sheet("Offset", (4, 1, T("Name")), (4, 2, T("Amount")),
            (5, 1, T("A")), (5, 2, N(2)), (6, 1, T("B")), (6, 2, N(1)));
        var expected = new CellRectangle(4, 1, 6, 2);
        Assert.Equal(expected, PersistentFiltering.ResolveRange(sheet, new(5, 2, 5, 2), new(5, 2)));
        Assert.Equal(expected, PersistentFiltering.ResolveRange(sheet, new(0, 2, WorkbookLimits.MaxRows - 1, 2), new(0, 2)));
        Assert.Equal(expected, PersistentFiltering.ResolveRange(sheet, new(4, 0, 6, WorkbookLimits.MaxColumns - 1), new(4, 0)));
    }

    [Fact]
    public void EmptyAndBottomEdgeHeadersStayInsideTheirSelectedBounds()
    {
        var empty = Sheet("Empty");
        var blankSelection = new CellRectangle(4, 2, 4, 4);
        Assert.Equal(blankSelection, PersistentFiltering.ResolveRange(empty, blankSelection, new(4, 4)));
        var last = WorkbookLimits.MaxRows - 1;
        var edge = Sheet("Edge", (last - 1, 0, T("Name")), (last - 1, 1, T("Value")), (last, 1, N(1)));
        Assert.Equal(new CellRectangle(last - 1, 0, last, 1), PersistentFiltering.ResolveRange(edge, new(last - 1, 0, last - 1, 1), new(last - 1, 1)));
        Assert.Equal(new CellRectangle(last, 0, last, 1), PersistentFiltering.ResolveRange(edge, new(last, 0, last, 1), new(last, 1)));
    }

    [Fact]
    public void HeaderFilterExpansionDoesNotUseTheSmallerSortCellBudget()
    {
        var cells = Enumerable.Range(0, SortPlanning.MaxCells + 1)
            .Select(row => new KeyValuePair<CellPosition, CellSnapshot>(new(row, 0), row == 0 ? T("Header") : N(row)));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Large filter", cells);
        Assert.Equal(new CellRectangle(0, 0, SortPlanning.MaxCells, 1),
            PersistentFiltering.ResolveRange(sheet, new(0, 0, 0, 1), new(0, 0)));
    }
}
