using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class PresentationCalculationTests
{
    [Fact]
    public void StyleAndItsHistoryPreserveFormulaResultsDiagnosticsAndCurrentRevision()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (0, 1, F("=A1*2")), (1, 1, F("=1/0")));
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate();
        var generation = doc.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.CalculationGeneration;
        var diagnostics = doc.Session.Diagnostics;
        DocumentChange? notification = null;
        doc.Session.Changed += (_, change) => notification = change;
        var result = doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 1), new(sheet.Id, 1, 1)], new(Bold: true, NumberFormatId: 2)));
        void Check(CommandResult change)
        {
            Assert.True(change.Changed && change.CalculationUnchanged);
            Assert.True(notification!.CalculationUnchanged);
            var fresh = doc.Snapshot.Sheets[0].GetCell(0, 1)!;
            Assert.Equal("=A1*2", fresh.Content.FormulaText);
            Assert.Equal(CellValue.FromNumber(4), fresh.Value);
            Assert.Equal(CalculationState.Fresh, fresh.Calculation!.State);
            Assert.Equal(doc.Session.Revision, fresh.Calculation.InputRevision);
            Assert.Equal(generation, fresh.Calculation.CalculationGeneration);
            var error = doc.Snapshot.Sheets[0].GetCell(1, 1)!.Calculation!;
            Assert.Equal(CalculationState.Error, error.State);
            Assert.Equal(doc.Session.Revision, error.InputRevision);
            Assert.Equal(diagnostics, doc.Session.Diagnostics);
        }
        Check(result);
        Assert.Equal("4.00", doc.GetCellValue(0, 0, 1));
        Check(doc.Session.UndoWithResult());
        Assert.Equal("4", doc.GetCellValue(0, 0, 1));
        Check(doc.Session.RedoWithResult());
        Assert.Equal("4.00", doc.GetCellValue(0, 0, 1));
        Assert.False(doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 1)], new(Bold: true))).Changed);
    }

    [Theory]
    [InlineData(CalculationState.Stale)]
    [InlineData(CalculationState.Blocked)]
    [InlineData(CalculationState.Dirty)]
    public void StylingNeverPromotesInvalidCaches(CalculationState state)
    {
        var cache = new CellCalculation(state == CalculationState.Stale ? CellValue.FromNumber(9) : null, state, 0, 1, 0, ["retained"]);
        var sheet = Sheet("Data", (0, 0, new(CellContent.FromFormula("=1"), cache)), (1, 0, F("=2")));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0)], new(FillColor: "FFFFFF00")));
        for (var i = 0; i < 3; i++)
        {
            Assert.Same(cache, session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation);
            Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 0)!.Calculation);
            Assert.False(session.ScriptsEnabled);
            if (i == 0) session.Undo(); else if (i == 1) session.Redo();
        }
    }

    [Fact]
    public void PresentationUndoUsesLatestCachesAndNeverRevivesResultsFromOldScripts()
    {
        var sheet = Sheet("Data", (0, 0, F("=1")));
        using var session = new DocumentSession(Book(sheet));
        void Commit(double value)
        {
            var request = session.BeginCalculation();
            Assert.True(session.TryCommitCalculation(new(request.Stamp, new Dictionary<CellAddress, CellCalculation>
            { [new(sheet.Id, 0, 0)] = new(CellValue.FromNumber(value), CalculationState.Fresh, request.Stamp.InputRevision,
                request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration) })));
        }
        Commit(1);
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Bold: true)));
        Commit(2);
        session.Undo();
        Assert.Equal(CellValue.FromNumber(2), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        session.Execute(new UpdateScriptSourceCommand("// new source"));
        var stale = session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation;
        session.Redo();
        Assert.Same(stale, session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation);
        Assert.False(session.ScriptsEnabled);
    }

    [Fact]
    public void StyleRejectsLateResultsAndContentHistoryStillInvalidates()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (0, 1, F("=A1*2")));
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate();
        var request = doc.Session.BeginCalculation();
        doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 1)], new(Bold: true)));
        Assert.True(request.CancellationToken.IsCancellationRequested);
        Assert.False(doc.Session.TryCommitCalculation(new(request.Stamp, new Dictionary<CellAddress, CellCalculation>())));
        doc.SetCell(0, 0, 0, "3");
        Assert.Equal(CalculationState.Stale, doc.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State);
        doc.Calculate();
        Assert.False(doc.Session.UndoWithResult().CalculationUnchanged);
        Assert.Equal(CalculationState.Stale, doc.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State);
        Assert.False(doc.Session.RedoWithResult().CalculationUnchanged);
    }
}
