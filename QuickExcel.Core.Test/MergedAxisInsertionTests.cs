using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class MergedAxisInsertionTests
{
    [Theory]
    [InlineData(SheetAxis.Row, "B2:D4", "B2:D5")]
    [InlineData(SheetAxis.Column, "B2:D4", "B2:E4")]
    [InlineData(SheetAxis.Row, "B2:D2", "B2:D3")]
    [InlineData(SheetAxis.Column, "B2:B4", "B2:C4")]
    [InlineData(SheetAxis.Row, "A1:C2", "A1:C3")]
    [InlineData(SheetAxis.Column, "A1:B3", "A1:C3")]
    public void InsertingIntoMergeAddsOneAxisRetainsAnchorAndReferencesAndUndoesAtomically(SheetAxis axis, string original, string expected)
    {
        var merge = CellRectangle.Parse(original); var row = axis == SheetAxis.Row;
        var anchor = new CellPosition(merge.Top, merge.Left);
        var boundary = row ? new CellPosition(merge.Top + 1, 6) : new CellPosition(6, merge.Left + 1);
        var moved = row ? boundary with { Row = boundary.Row + 1 } : boundary with { Column = boundary.Column + 1 };
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [
            new(anchor, new(CellContent.FromLiteral(CellValue.FromNumber(42), 1))),
            new(boundary, new(CellContent.FromLiteral(CellValue.FromNumber(7))))],
            rows: [new(merge.Top, new(Height: 32))], columns: [new(merge.Left, new(Width: 24))],
            mergedCells: [original, "H6:I7"]);
        var report = new WorksheetSnapshot(SheetId.New(), "Report", [new(new(0, 0),
            new(CellContent.FromFormula($"=SUM(Data!{new CellRectangle(anchor.Row, anchor.Column, anchor.Row, anchor.Column)},Data!{new CellRectangle(boundary.Row, boundary.Column, boundary.Row, boundary.Column)})")))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, report], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);

        session.Execute(ShiftAxisCommand.InsertIntoMergedCell(source.Id, axis, merge));
        var updated = session.Snapshot.GetSheet(source.Id);
        Assert.Equal(new[] { expected, row ? "H7:I8" : "I6:J7" }, updated.MergedCells.ToArray());
        Assert.Equal(source.GetCell(anchor.Row, anchor.Column)!.Content, updated.GetCell(anchor.Row, anchor.Column)!.Content);
        Assert.Equal(CellValue.FromNumber(7), updated.GetCell(moved.Row, moved.Column)!.Value);
        Assert.Equal(CellValue.Blank, updated.GetCell(boundary.Row, boundary.Column)?.Value ?? CellValue.Blank);
        var inserted = row ? anchor with { Row = anchor.Row + 1 } : anchor with { Column = anchor.Column + 1 };
        Assert.Equal(1, updated.GetStyleId(inserted.Row, inserted.Column));
        if (row) Assert.Equal(32, updated.Rows[inserted.Row].Height);
        else Assert.Equal(24, updated.Columns[inserted.Column].Width);
        var calculation = new WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromNumber(49), calculation.Cells[new(report.Id, 0, 0)].Value);
        Assert.Equal(1, session.Revision);

        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(source.MergedCells, session.Snapshot.GetSheet(source.Id).MergedCells);
        Assert.Equal(source.Cells, session.Snapshot.GetSheet(source.Id).Cells);
        Assert.True(session.Redo());
        Assert.Equal(updated.MergedCells, session.Snapshot.GetSheet(source.Id).MergedCells);
        Assert.Equal(updated.Cells, session.Snapshot.GetSheet(source.Id).Cells);
    }

    [Theory]
    [InlineData(SheetAxis.Row, "B2:D5")]
    [InlineData(SheetAxis.Column, "B2:E4")]
    public void ExpandedMergeAndContentSurviveXlsxRoundTrip(SheetAxis axis, string expected)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(1, 1), new(CellInput.Parse("合并内容")))], mergedCells: ["B2:D4"]);
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]));
        session.Execute(ShiftAxisCommand.InsertIntoMergedCell(source.Id, axis, new(1, 1, 3, 3)));
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "merged-insert-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "merged.xlsx"); var storage = new TypedWorkbookStorage();
        storage.Save(path, session.Snapshot);
        var reopened = storage.Open(path).Sheets[0];
        Assert.Equal(expected, Assert.Single(reopened.MergedCells));
        Assert.Equal(CellValue.FromText("合并内容"), reopened.GetCell(1, 1)!.Value);
    }

    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void MissingMergeAndOutOfBoundsInsertionFailWithoutHistory(SheetAxis axis)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data");
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]));
        var before = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(ShiftAxisCommand.InsertIntoMergedCell(source.Id, axis, new(1, 1, 3, 3))));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);

        var edge = axis == SheetAxis.Row ? new CellRectangle(WorkbookLimits.MaxRows - 2, 0, WorkbookLimits.MaxRows - 1, 1)
            : new CellRectangle(0, WorkbookLimits.MaxColumns - 2, 1, WorkbookLimits.MaxColumns - 1);
        source = new(source.Id, source.Name, mergedCells: [edge.ToString()]);
        using var edgeSession = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]));
        before = edgeSession.Snapshot;
        Assert.Throws<ArgumentException>(() => edgeSession.Execute(ShiftAxisCommand.InsertIntoMergedCell(source.Id, axis, edge)));
        Assert.Same(before, edgeSession.Snapshot); Assert.False(edgeSession.CanUndo);
    }

    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void ProtectedMergeCannotBypassExistingStructureGuards(SheetAxis axis)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data", mergedCells: ["B2:D4"], data: new() { Protection = new() { Sheet = true } });
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(ShiftAxisCommand.InsertIntoMergedCell(source.Id, axis, new(1, 1, 3, 3))));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }
}
