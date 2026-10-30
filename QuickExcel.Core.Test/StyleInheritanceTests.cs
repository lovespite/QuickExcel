using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class StyleInheritanceTests
{
    private static WorkbookSnapshot Model()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Styles", rows: [
            new(0, new(StyleId: 0, ApplyStyle: true)),
            new(1, new(StyleId: 2, ApplyStyle: false)),
            new(2, new(Height: 30)),
            new(3, new(StyleId: 2))
        ], columns: Enumerable.Range(0, 4).Select(c => new KeyValuePair<int, ColumnMetadata>(c, new(StyleId: 1))));
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }, new() { Bold = true }]);
    }

    [Fact]
    public void EmptyCellsAndNewInputUseTheSameExplicitRowThenColumnPrecedence()
    {
        using var doc = new WorkbookDocument(Model(), new(new TypedWorkbookStorage(), new WorkbookCalculator(new M3CalculationTests.NoScripts())));
        int[] expected = [0, 1, 1, 2, 1];
        for (var row = 0; row < expected.Length; row++)
        {
            Assert.Equal(expected[row], doc.Snapshot.Sheets[0].GetStyleId(row, 0));
            doc.SetCell(0, row, 0, "61");
            Assert.Equal(expected[row], doc.Snapshot.Sheets[0].GetCell(row, 0)!.Content.StyleId);
        }
        Assert.Equal("61", doc.GetCellValue(0, 0, 0));
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 1, 0));
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 2, 0));
        Assert.True(doc.Snapshot.Styles[doc.Snapshot.Sheets[0].GetStyleId(3, 0)].Bold);
    }

    [Fact]
    public void ClearingAbsentCellsDoesNotCreateDefaultOverridesOrUndoHistory()
    {
        using var session = new DocumentSession(Model()); var before = session.Snapshot; var id = before.Sheets[0].Id;
        session.Execute(new ClearCellsCommand(Enumerable.Range(0, 5).Select(row => new CellAddress(id, row, 1))));
        Assert.Same(before, session.Snapshot);
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(1, 1));
    }

    [Fact]
    public void ClearingStoredCellsRetainsNecessaryDefaultOverridesAndCanUndoAtomically()
    {
        using var session = new DocumentSession(Model()); var id = session.Snapshot.Sheets[0].Id;
        var addresses = new[] { new CellAddress(id, 0, 1), new(id, 1, 1), new(id, 2, 1) };
        session.Execute(new SetCellsCommand(addresses.Select(a => new CellEdit(a, CellContent.FromLiteral(CellValue.FromNumber(61))))));
        var populated = session.Snapshot;
        session.Execute(new ClearCellsCommand(addresses));
        Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 1)); // Active row 0 already supplies style 0.
        Assert.NotNull(session.Snapshot.Sheets[0].GetCell(1, 1)); // Disabled row must not mask the column style.
        Assert.NotNull(session.Snapshot.Sheets[0].GetCell(2, 1));
        Assert.All(addresses, a => Assert.Equal(0, session.Snapshot.GetSheet(id).GetStyleId(a.Row, a.Column)));
        session.Undo();
        Assert.All(addresses, a => Assert.Equal(populated.GetSheet(id).GetCell(a.Row, a.Column)!.Content,
            session.Snapshot.GetSheet(id).GetCell(a.Row, a.Column)!.Content));
    }

    [Fact]
    public void ResizeHideAndRowInsertionKeepActivationAndUndoRestoresIt()
    {
        using var session = new DocumentSession(Model()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 0, 4, 36, true));
        Assert.True(session.Snapshot.GetSheet(id).Rows[0].ApplyStyle);
        Assert.False(session.Snapshot.GetSheet(id).Rows[1].ApplyStyle);
        Assert.False(session.Snapshot.GetSheet(id).Rows[2].ApplyStyle);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1));
        Assert.True(session.Snapshot.GetSheet(id).Rows[1].ApplyStyle);
        Assert.False(session.Snapshot.GetSheet(id).Rows[2].ApplyStyle);
        Assert.Equal(0, session.Snapshot.GetSheet(id).GetStyleId(1, 1));
        session.Undo(); session.Undo();
        Assert.Equal(30, session.Snapshot.GetSheet(id).Rows[2].Height);
        Assert.False(session.Snapshot.GetSheet(id).Rows[0].Hidden);
        Assert.True(session.Snapshot.GetSheet(id).Rows[3].ApplyStyle);
    }

    [Fact]
    public void ClearingStillRejectsInvalidAddressesDuplicatesAndOversizedRequestsAtomically()
    {
        using var session = new DocumentSession(Model()); var snapshot = session.Snapshot; var id = snapshot.Sheets[0].Id;
        Assert.Throws<ArgumentException>(() => session.Execute(new ClearCellsCommand([new(id, 0, 0), new(id, 0, 0)])));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.Execute(new ClearCellsCommand([new(id, -1, 0)])));
        Assert.Throws<ArgumentException>(() => session.Execute(new ClearCellsCommand(Enumerable.Repeat(new CellAddress(id, 0, 0), WorkbookLimits.MaxCells + 1))));
        Assert.Same(snapshot, session.Snapshot);
    }
}
