using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DeleteCellsTests
{
    [Theory]
    [InlineData(CellDeleteShift.Left)]
    [InlineData(CellDeleteShift.Up)]
    public void DeleteClosesOnlyTheSelectedStripAndOneUndoRestoresEverything(CellDeleteShift direction)
    {
        CellPosition P(int row, int column) => direction == CellDeleteShift.Left ? new(row, column) : new(column, row);
        var source = new WorksheetSnapshot(SheetId.New(), "Data", new[]
        {
            new KeyValuePair<CellPosition, CellSnapshot>(P(1, 0), N(10)), new(P(1, 1), N(20)),
            new(P(1, 2), N(30)), new(P(1, 3), N(40)), new(P(1, 5), N(60)), new(P(2, 3), N(70)), new(P(3, 3), N(80))
        });
        var deleted = direction == CellDeleteShift.Left ? new CellRectangle(1, 1, 2, 2) : new(1, 1, 2, 2);
        using var session = new DocumentSession(Book(source));
        session.Execute(new DeleteCellsCommand(source.Id, deleted, direction));
        var result = session.Snapshot.GetSheet(source.Id);
        Assert.Equal(CellValue.FromNumber(10), result.GetCell(P(1, 0).Row, P(1, 0).Column)!.Value);
        Assert.Equal(CellValue.FromNumber(40), result.GetCell(P(1, 1).Row, P(1, 1).Column)!.Value);
        Assert.Null(result.GetCell(P(1, 2).Row, P(1, 2).Column));
        Assert.Equal(CellValue.FromNumber(60), result.GetCell(P(1, 3).Row, P(1, 3).Column)!.Value);
        Assert.Equal(CellValue.FromNumber(70), result.GetCell(P(2, 1).Row, P(2, 1).Column)!.Value);
        Assert.Equal(CellValue.FromNumber(80), result.GetCell(P(3, 3).Row, P(3, 3).Column)!.Value);
        Assert.Equal(1, session.Revision);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(source.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column), session.Snapshot.GetSheet(source.Id).Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column));
        Assert.True(session.Redo()); Assert.Equal(CellValue.FromNumber(40), session.Snapshot.GetSheet(source.Id).GetCell(P(1, 1).Row, P(1, 1).Column)!.Value);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ReferencesNamesAndTravellingFormulasKeepTargetsAndPersist(string extension)
    {
        var data = Sheet("Data", (0, 0, N(10)), (1, 1, N(20)), (1, 2, N(30)), (1, 3, N(40)), (1, 4, F("=$D$2+$A$1")), (2, 3, N(70)));
        var report = Sheet("Report", (0, 0, F("=Data!$D$2+Data!D$2+Data!$D2+Data!D2")),
            (1, 0, F("=Data!B2")), (2, 0, F("=SUM(MovedValues)")), (3, 0, F("=SUM(Data!B2:E2)")), (4, 0, F("=SUM(Data!B:B)")));
        var book = Book(data, report).WithDefinedNames([new("MovedValues", "=Data!$D$2:$D$3")]);
        using var session = new DocumentSession(book);
        session.Execute(new DeleteCellsCommand(data.Id, new(1, 1, 2, 2), CellDeleteShift.Left));
        var references = FormulaSyntax.Walk(FormulaSyntax.Parse(session.Snapshot.GetSheet(report.Id).GetCell(0, 0)!.Content.FormulaText)).OfType<ReferenceNode>().ToArray();
        Assert.All(references, r => { Assert.Equal(1, r.Row); Assert.Equal(1, r.Column); });
        Assert.True(references[0].ColumnAbsolute && references[0].RowAbsolute);
        Assert.False(references[1].ColumnAbsolute); Assert.True(references[1].RowAbsolute);
        Assert.True(references[2].ColumnAbsolute); Assert.False(references[2].RowAbsolute);
        Assert.Equal("=SUM(Data!B:B)", session.Snapshot.GetSheet(report.Id).GetCell(4, 0)!.Content.FormulaText);
        void Check(WorkbookSnapshot snapshot)
        {
            var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(snapshot));
            var s = snapshot.Sheets.Single(s => s.Name == "Report"); var d = snapshot.Sheets.Single(s => s.Name == "Data");
            Assert.Equal(CellValue.FromNumber(160), C(result, s).Value);
            Assert.Equal(CellValue.FromError(CellErrorCode.Reference), C(result, s, 1).Value);
            Assert.Equal(CellValue.FromNumber(110), C(result, s, 2).Value);
            Assert.Equal(CellValue.FromNumber(90), C(result, s, 3).Value);
            Assert.Equal(CellValue.FromNumber(50), C(result, d, 1, 2).Value);
            var name = Assert.IsType<ReferenceNode>(FormulaSyntax.Parse(snapshot.DefinedNames[0].Formula));
            Assert.Equal(new CellRectangle(1, 1, 2, 1), new(name.Row, name.Column, name.EndRow, name.EndColumn));
        }
        Check(session.Snapshot);
        var directory = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "delete-cells", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "deleted." + extension);
        var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); Check(storage.Open(path));
    }

    [Fact]
    public void UpwardDeletionShrinksVerticalReferencesAndNames()
    {
        var data = Sheet("Data", (0, 1, N(10)), (1, 1, N(20)), (2, 1, N(30)), (3, 1, N(40)));
        var report = Sheet("Report", (0, 0, F("=SUM(Values)")), (1, 0, F("=Data!$B$4")), (2, 0, F("=Data!B2")));
        using var session = new DocumentSession(Book(data, report).WithDefinedNames([new("Values", "=Data!$B$1:$B$4")]));
        session.Execute(new DeleteCellsCommand(data.Id, new(1, 1, 2, 1), CellDeleteShift.Up));
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromNumber(50), C(result, report).Value);
        Assert.Equal(CellValue.FromNumber(40), C(result, report, 1).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Reference), C(result, report, 2).Value);
    }

    [Theory]
    [InlineData(CellDeleteShift.Left)]
    [InlineData(CellDeleteShift.Up)]
    public void SparseDeletionAtTheLastRowOrColumnDoesNotAllocateTheEmptyStrip(CellDeleteShift direction)
    {
        var last = direction == CellDeleteShift.Left ? new CellPosition(1, WorkbookLimits.MaxColumns - 1) : new(WorkbookLimits.MaxRows - 1, 1);
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [new(last, N(9))]);
        using var session = new DocumentSession(Book(source));
        session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 1, 1), direction));
        var stored = Assert.Single(session.Snapshot.GetSheet(source.Id).Cells);
        Assert.Equal(direction == CellDeleteShift.Left ? last with { Column = last.Column - 1 } : last with { Row = last.Row - 1 }, stored.Key);
        Assert.Equal(CellValue.FromNumber(9), stored.Value.Value);
    }

    [Fact]
    public void CompleteMergesAreDeletedOrRelocatedAndRichRunsTravelUnchanged()
    {
        var rich = new RichCellText([new("Bold", new(Bold: true)), new(" normal")]);
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(1, 1), T("Deleted")), new(new(1, 3), new(CellContent.FromRichText(rich)))],
            mergedCells: ["B2:C3", "D2:E3", "A6:B6"]);
        using var session = new DocumentSession(Book(source));
        session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 2, 2), CellDeleteShift.Left));
        var sheet = session.Snapshot.GetSheet(source.Id);
        Assert.Equal(new[] { "B2:C3", "A6:B6" }, sheet.MergedCells);
        Assert.Equal(rich, sheet.GetCell(1, 1)!.Content.RichText);
        Assert.Null(sheet.GetCell(1, 3));
        Assert.True(session.Undo()); Assert.Equal("Deleted", session.Snapshot.GetSheet(source.Id).GetCell(1, 1)!.Value.TextValue);
    }

    [Theory]
    [InlineData(CellDeleteShift.Left)]
    [InlineData(CellDeleteShift.Up)]
    public void InheritedBlankStylesMoveWhileDimensionsRemainFixed(CellDeleteShift direction)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data", rows: [new(1, new(Height: 35)), new(3, new(Height: 45, StyleId: 1))],
            columns: [new(1, new(Width: 125)), new(3, new(Width: 155, StyleId: 2))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source], styles: [new(), new() { Bold = true }, new() { Italic = true }]);
        using var session = new DocumentSession(book);
        session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 2, 2), direction));
        var sheet = session.Snapshot.GetSheet(source.Id);
        Assert.Equal(direction == CellDeleteShift.Left ? 2 : 1, sheet.GetStyleId(1, 1));
        Assert.Equal(0, direction == CellDeleteShift.Left ? sheet.GetStyleId(1, 3) : sheet.GetStyleId(3, 1));
        Assert.Equal(source.Rows, sheet.Rows); Assert.Equal(source.Columns, sheet.Columns);
        Assert.True(session.Undo()); Assert.Empty(session.Snapshot.GetSheet(source.Id).Cells);
    }

    [Fact]
    public void IgnoredErrorPreferencesSplitAndMoveWithTheCellsAndRetainOtherMetadata()
    {
        var source = Sheet("Data", (0, 0, N(1))).WithData(new()
        {
            DefaultRowDescent = .25,
            IgnoredErrors = [new(new(0, 0, 2, 4), IgnoredCellErrors.NumberStoredAsText)]
        });
        using var session = new DocumentSession(Book(source));
        session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 1, 2), CellDeleteShift.Left));
        var sheet = session.Snapshot.GetSheet(source.Id);
        bool At(int row, int column) => sheet.Data.IgnoredErrors.Any(e => e.Range.Contains(row, column));
        Assert.True(At(0, 4)); Assert.True(At(2, 4)); Assert.True(At(1, 2)); Assert.False(At(1, 3));
        Assert.Equal(.25, sheet.Data.DefaultRowDescent);
        Assert.True(session.Undo()); Assert.Equal(source.Data, session.Snapshot.GetSheet(source.Id).Data);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FragmentedReferenceOrMergeFailsWithoutPublishingAnything(bool merge)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(0, 0), N(1))], mergedCells: merge ? ["A1:C3"] : null);
        var report = Sheet("Report", (0, 0, F("=SUM(Data!A1:C3)")));
        using var session = new DocumentSession(merge ? Book(source) : Book(source, report));
        var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 1, 1), CellDeleteShift.Left)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo); Assert.Equal(0, session.Revision);
    }

    [Fact]
    public void ProtectedHiddenAnnotatedAndValidatedWorkbooksRetainTheirSafetyGuards()
    {
        var sheets = new[]
        {
            new WorksheetSnapshot(SheetId.New(), "Data", data: new() { Protection = new() }),
            new WorksheetSnapshot(SheetId.New(), "Data", rows: [new(1, new(Hidden: true))]),
            new WorksheetSnapshot(SheetId.New(), "Data", annotations: new(notes: [new(new(1, 1), new("Author", new([new("Note")])))])),
            new WorksheetSnapshot(SheetId.New(), "Data", data: new() { Validations = [new([new(1, 1, 2, 2)], ValidationKind.List, "=\"A,B\"")] }),
            new WorksheetSnapshot(SheetId.New(), "Data", data: new() { Tables = [new(1, "Sales", new(0, 0, 3, 2), [new(1, "A"), new(2, "B"), new(3, "C")])] })
        };
        foreach (var sheet in sheets)
        {
            using var session = new DocumentSession(Book(sheet)); var before = session.Snapshot;
            Assert.ThrowsAny<Exception>(() => session.Execute(new DeleteCellsCommand(sheet.Id, new(1, 1, 1, 1), CellDeleteShift.Left)));
            Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        }
    }

    [Fact]
    public void PartialSpillDeletionFailsButDeletingTheEntireArrayIsAtomic()
    {
        var source = Sheet("Data", (1, 1, F("=SEQUENCE(3,2)")), (1, 3, N(99)));
        using var session = new DocumentSession(Book(source));
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new DeleteCellsCommand(source.Id, new(2, 1, 2, 1), CellDeleteShift.Left)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new DeleteCellsCommand(source.Id, new(1, 1, 3, 2), CellDeleteShift.Left));
        Assert.Equal(CellValue.FromNumber(99), session.Snapshot.GetSheet(source.Id).GetCell(1, 1)!.Value);
        Assert.DoesNotContain(session.Snapshot.GetSheet(source.Id).Cells.Values, c => c.SpillOwner is not null);
        Assert.True(session.Undo()); Assert.Equal("=SEQUENCE(3,2)", session.Snapshot.GetSheet(source.Id).GetCell(1, 1)!.Content.FormulaText);
    }
}
