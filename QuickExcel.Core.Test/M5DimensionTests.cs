using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M5DimensionTests
{
    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void ResizeCarriesFreshAndErrorResultsToNewRevisionWithoutCalculating(SheetAxis axis)
    {
        var sheet = Sheet("Data", (0, 0, F("=2")), (1, 0, F("=1")), (2, 0, F("=1/0")));
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate();
        var old = doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Calculation!;
        var result = doc.Session.Execute(new SetAxisMetadataCommand(sheet.Id, axis, 0, 1, 42));
        Assert.True(result.CalculationUnchanged);
        var fresh = doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Calculation!;
        Assert.Equal(CalculationState.Fresh, fresh.State);
        Assert.Equal(old.Value, fresh.Value);
        Assert.Equal(doc.Snapshot.Revision, fresh.InputRevision);
        Assert.Equal(old.CalculationGeneration, fresh.CalculationGeneration);
        var error = doc.Snapshot.GetSheet(sheet.Id).GetCell(2, 0)!.Calculation!;
        Assert.Equal(CalculationState.Error, error.State);
        Assert.Equal(doc.Snapshot.Revision, error.InputRevision);
        // Sorting requires current-revision keys; a pure resize must not make them stale.
        Assert.True(doc.Session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0)).Changed);
    }

    [Fact]
    public void ResizeKeepsStaleBlockedAndMissingCachesAndNeverGrantsScriptPermission()
    {
        var sheet = Sheet("Data", (0, 0, F("=1")), (1, 0, F("=UNKNOWN()")), (2, 0, F("=2")));
        var cells = sheet.Cells.ToBuilder();
        cells[new(0, 0)] = new(cells[new(0, 0)].Content, new(CellValue.FromNumber(9), CalculationState.Stale, 0, 1, 0));
        cells[new(1, 0)] = new(cells[new(1, 0)].Content, new(null, CalculationState.Blocked, 0, 1, 0, ["disabled"]));
        using var session = new DocumentSession(Book(sheet.WithCells(cells.ToImmutable())));
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 0, 1, 70));
        foreach (var position in cells.Keys) Assert.Same(cells[position].Calculation, session.Snapshot.GetSheet(sheet.Id).Cells[position].Calculation);
        Assert.False(session.ScriptsEnabled);
    }

    [Fact]
    public void ResizeCancelsInFlightCalculationAndRetainsUndoRedoAndSaveRevisions()
    {
        var sheet = Sheet("Data", (0, 0, F("=1")));
        using var session = new DocumentSession(Book(sheet));
        var request = session.BeginCalculation();
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Column, 0, 1, 25));
        Assert.True(request.CancellationToken.IsCancellationRequested);
        Assert.False(session.TryCommitCalculation(new(request.Stamp, new Dictionary<CellAddress, CellCalculation>())));
        Assert.True(session.IsDirty);
        Assert.True(session.Undo()); Assert.Empty(session.Snapshot.GetSheet(sheet.Id).Columns);
        Assert.True(session.Redo()); Assert.Equal(25, session.Snapshot.GetSheet(sheet.Id).Columns[0].Width);
        Assert.True(session.MarkSaved(session.SessionId, session.Revision)); Assert.False(session.IsDirty);
    }

    [Fact]
    public void HidingAndInputEditsStillInvalidateCalculations()
    {
        var sheet = Sheet("Data", (0, 0, F("=1")));
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate();
        Assert.False(doc.Session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 0, 1, 42, hidden: true)).CalculationUnchanged);
        Assert.Equal(CalculationState.Stale, doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Calculation!.State);
        doc.Calculate();
        doc.SetCell(0, 1, 0, "new input");
        Assert.Equal(CalculationState.Stale, doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Calculation!.State);
    }
}
