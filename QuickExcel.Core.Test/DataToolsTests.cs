using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DataToolsTests
{
    [Fact]
    public void FlashFillExtractsTokensKeepsExamplesAndIsOneUndoableTextTransaction()
    {
        var sheet = Sheet("S", (0, 0, T("Ada Lovelace")), (0, 1, T("Ada")), (1, 0, T("Grace Hopper")), (2, 0, T("Alan Turing")));
        using var session = new DocumentSession(Book(sheet)); var original = session.Snapshot;
        var preview = FlashFill.Preview(original, sheet.Id, new(0, 0, 2, 1), 1);
        Assert.Same(original, session.Snapshot); Assert.Empty(preview.Overwrites);
        session.Execute(new ApplyCleaningCommand(preview));
        Assert.Equal("Grace", session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.TextValue);
        Assert.Equal("Alan", session.Snapshot.Sheets[0].GetCell(2, 1)!.Value.TextValue);
        Assert.Equal("Ada", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        session.Undo(); Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 1)); session.Redo();
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplyCleaningCommand(preview)));
    }

    [Fact]
    public void FlashFillCombinesColumnsAndChangesCaseWithoutExecutingFormulaText()
    {
        var sheet = Sheet("S", (0, 0, T("Ada")), (0, 1, T("Lovelace")), (0, 2, T("=lovelace, ADA")),
            (1, 0, T("Grace")), (1, 1, T("Hopper")), (1, 2, T("=hopper, GRACE")), (2, 0, T("Alan")), (2, 1, T("Turing")));
        var p = FlashFill.Preview(Book(sheet), sheet.Id, new(0, 0, 2, 2), 2);
        Assert.Equal("=turing, ALAN", p.Result.Sheets[0].GetCell(2, 2)!.Value.TextValue);
        Assert.Equal(CellContentKind.Literal, p.Result.Sheets[0].GetCell(2, 2)!.Content.Kind);
    }

    [Fact]
    public void FlashFillRejectsAmbiguityInconsistentExamplesAndHiddenRows()
    {
        var sheet = Sheet("S", (0, 0, T("Alex Alex")), (0, 1, T("Alex")), (1, 0, T("Grace Hopper")));
        Assert.Throws<InvalidOperationException>(() => FlashFill.Preview(Book(sheet), sheet.Id, new(0, 0, 1, 1), 1));
        sheet = Sheet("S", (0, 0, T("Ada Lovelace")), (0, 1, T("Ada")), (1, 0, T("Grace Hopper")), (1, 1, T("unrelated")), (2, 0, T("Alan Turing")));
        Assert.Throws<InvalidOperationException>(() => FlashFill.Preview(Book(sheet), sheet.Id, new(0, 0, 2, 1), 1));
        using var session = new DocumentSession(Book(sheet)); session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 1, 1, hidden: true));
        Assert.Throws<InvalidOperationException>(() => FlashFill.Preview(session.Snapshot, sheet.Id, new(0, 0, 2, 1), 1));
    }

    [Theory]
    [InlineData("=A1*3+2", 32, 10)]
    [InlineData("=A1^2", 25, 5)]
    public void GoalSeekUsesPrivateDraftsAndOnlyCommitsTheSolvedInput(string formula, double target, double expected)
    {
        var sheet = Sheet("S", (0, 0, N(1)), (0, 1, F(formula))); var book = Book(sheet);
        var calculator = new WorkbookCalculator(new NoScripts());
        var result = DataAnalysis.GoalSeek(book, calculator, new(sheet.Id, 0, 1), new(sheet.Id, 0, 0), target);
        Assert.Equal(expected, result.Input, 5); Assert.Equal(1, book.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal(formula, result.Preview.Result.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        using var session = new DocumentSession(book);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplyCleaningCommand(result.Preview)));
        session.Execute(new ApplyCleaningCommand(result.Preview, true)); session.Undo(); Assert.Equal(1, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Fact]
    public void AnalysisRejectsFailedFormulaNonConvergenceCancellationAndSourceOverlap()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (0, 1, F("=A1*0+4"))); var book = Book(sheet); var calculator = new WorkbookCalculator(new NoScripts());
        Assert.Throws<InvalidOperationException>(() => DataAnalysis.GoalSeek(book, calculator, new(sheet.Id, 0, 1), new(sheet.Id, 0, 0), 5));
        Assert.Throws<ArgumentException>(() => DataAnalysis.GoalSeek(book, calculator, new(sheet.Id, 0, 1), new(sheet.Id, 0, 1), 5));
        Assert.Throws<OperationCanceledException>(() => DataAnalysis.GoalSeek(book, calculator, new(sheet.Id, 0, 1), new(sheet.Id, 0, 0), 5, cancellation: new(true)));
        Assert.Throws<ArgumentException>(() => DataAnalysis.DataTable(book, calculator, new(sheet.Id, 0, 1), new(sheet.Id, 0, 0), [1, 2], null, [], sheet.Id, new(0, 0)));
        var errors = Sheet("S", (0, 0, N(1)), (0, 1, F("=1/0")));
        Assert.Throws<InvalidOperationException>(() => DataAnalysis.GoalSeek(Book(errors), calculator, new(errors.Id, 0, 1), new(errors.Id, 0, 0), 5));
    }

    [Fact]
    public void TwoVariableTableAndNamedScenariosHaveCorrectResultsAndKeepInputs()
    {
        var sheet = Sheet("S", (0, 0, N(5)), (0, 1, N(7)), (0, 2, F("=A1*B1"))); var book = Book(sheet); var calculator = new WorkbookCalculator(new NoScripts());
        var p = DataAnalysis.DataTable(book, calculator, new(sheet.Id, 0, 2), new(sheet.Id, 0, 0), [2, 3], new(sheet.Id, 0, 1), [10, 20], sheet.Id, new(2, 0));
        Assert.Equal(20, p.Result.Sheets[0].GetCell(3, 1)!.Value.NumberValue); Assert.Equal(60, p.Result.Sheets[0].GetCell(4, 2)!.Value.NumberValue);
        p = DataAnalysis.Scenarios(book, calculator, new(sheet.Id, 0, 2), [new(sheet.Id, 0, 0), new(sheet.Id, 0, 1)], [new("基准", [2, 4]), new("增长", [3, 6])], sheet.Id, new(2, 0));
        Assert.Equal(8, p.Result.Sheets[0].GetCell(3, 3)!.Value.NumberValue); Assert.Equal(18, p.Result.Sheets[0].GetCell(4, 3)!.Value.NumberValue);
        Assert.Equal(5, p.Result.Sheets[0].GetCell(0, 0)!.Value.NumberValue); Assert.Equal(7, p.Result.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
    }

    [Theory]
    [InlineData(ForecastMethod.Linear)]
    [InlineData(ForecastMethod.Holt)]
    public void ForecastExtendsKnownTrendAndRejectsUnevenTimeline(ForecastMethod method)
    {
        var sheet = Sheet("S", (0, 0, T("Time")), (0, 1, T("Value")), (1, 0, N(1)), (1, 1, N(3)), (2, 0, N(2)), (2, 1, N(5)), (3, 0, N(3)), (3, 1, N(7)));
        var p = DataAnalysis.Forecast(Book(sheet), sheet.Id, new(0, 0, 3, 1), true, 2, new(0, 3), method);
        Assert.Equal(4, p.Result.Sheets[0].GetCell(4, 3)!.Value.NumberValue); Assert.Equal(9, p.Result.Sheets[0].GetCell(4, 5)!.Value.NumberValue, 8);
        Assert.Equal(11, p.Result.Sheets[0].GetCell(5, 5)!.Value.NumberValue, 8);
        Assert.Throws<ArgumentException>(() => DataAnalysis.Forecast(Book(sheet), sheet.Id, new(0, 0, 3, 1), true, 2, new(0, 0)));
        using var session = new DocumentSession(Book(sheet)); session.Execute(new SetCellsCommand([new(new(sheet.Id, 3, 0), CellContent.FromLiteral(CellValue.FromNumber(4)))]));
        Assert.Throws<ArgumentException>(() => DataAnalysis.Forecast(session.Snapshot, sheet.Id, new(0, 0, 3, 1), true, 2, new(0, 3)));
    }

    private static WorksheetSnapshot Records() => Sheet("S", (0, 0, T("Cat.")), (0, 1, T("Count")), (1, 0, T("A")), (1, 1, N(1)),
        (2, 0, T("A")), (2, 1, N(2)), (3, 0, T("B")), (3, 1, N(3)), (4, 0, T("B")), (4, 1, N(4)), (5, 0, T("B")), (5, 1, N(1)),
        (6, 0, T("C")), (6, 1, N(6)), (7, 0, T("C")), (7, 1, N(7)), (8, 0, T("C")), (8, 1, N(8)));

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SubtotalsSupportTreeLevelsReplacementRemovalPageBreaksAndRoundTrip(bool below)
    {
        var sheet = Records(); using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        var preview = SubtotalCommands.Preview(session.Snapshot, sheet.Id, new(0, 0, 8, 1), 0, [1], options: new(PageBreaks: true, SummaryBelow: below));
        session.Execute(new ApplySubtotalsCommand(preview)); Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Assert.Equal(32, session.Snapshot.Sheets[0].GetCell(preview.GrandTotalRow, 1)!.Value.NumberValue);
        Assert.Equal(new double[] { 3, 8, 21 }, preview.Groups.Select(g => session.Snapshot.Sheets[0].GetCell(g.SummaryRow, 1)!.Value.NumberValue));
        Assert.Equal(2, session.Snapshot.Sheets[0].Printing.RowBreaks.Length);
        session.Execute(new ShowOutlineLevelCommand(sheet.Id, SheetAxis.Row, 1));
        Assert.All(Enumerable.Range(1, 12).Where(r => r != preview.GrandTotalRow), r => Assert.True(session.Snapshot.Sheets[0].Rows[r].Hidden));
        session.Execute(new ShowOutlineLevelCommand(sheet.Id, SheetAxis.Row, 2));
        Assert.All(preview.Groups, g => { Assert.False(session.Snapshot.Sheets[0].Rows[g.SummaryRow].Hidden); Assert.True(session.Snapshot.Sheets[0].Rows[g.FirstRow].Hidden); });
        session.Execute(new ShowOutlineLevelCommand(sheet.Id, SheetAxis.Row, 3)); Assert.All(session.Snapshot.Sheets[0].Rows.Values, r => Assert.False(r.Hidden));
        var path = Path.Combine(Path.GetTempPath(), "qe-data-tools-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
            var reopened = storage.Open(path); var id = reopened.Sheets[0].Id;
            var replacement = SubtotalCommands.Preview(reopened, id, preview.Range, 0, [1], 1, new(SummaryBelow: !below));
            Assert.Equal(4, replacement.RemovedRows); Assert.Equal(preview.Range, replacement.Range);
            var removal = SubtotalCommands.Remove(replacement.Result, id, replacement.Range);
            Assert.Equal(4, removal.RemovedRows); Assert.Equal(new CellRectangle(0, 0, 8, 1), removal.Range);
            Assert.False(WorksheetOutline.HasGroups(removal.Result.Sheets[0])); Assert.Empty(removal.Result.Sheets[0].Printing.RowBreaks);
            Assert.Equal("C", removal.Result.Sheets[0].GetCell(8, 0)!.Value.TextValue);
        }
        finally { File.Delete(path); }
        session.Undo(); session.Redo(); Assert.All(session.Snapshot.Sheets[0].Rows.Values, r => Assert.False(r.Hidden));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SubtotalsCanAddAnotherAggregateAndRemoveBothLevels(bool below)
    {
        var sheet = Records(); var first = SubtotalCommands.Preview(Book(sheet), sheet.Id, new(0, 0, 8, 1), 0, [1], options: new(SummaryBelow: below));
        var second = SubtotalCommands.Preview(first.Result, sheet.Id, first.Range, 0, [1], 1, new(false, SummaryBelow: below));
        using var session = new DocumentSession(second.Result); var calculator = new WorkbookCalculator(new NoScripts());
        Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Assert.Equal(4, session.Snapshot.Sheets[0].GetCell(second.GrandTotalRow, 1)!.Value.NumberValue);
        var removed = SubtotalCommands.Remove(session.Snapshot, sheet.Id, second.Range);
        Assert.Equal(8, removed.RemovedRows); Assert.Equal(18, removed.Result.Sheets[0].Cells.Count(p => p.Value.Value.Kind != CellValueKind.Blank));
    }

    [Fact]
    public void SubtotalRemovalRejectsExtraDataInSummaryRows()
    {
        var sheet = Records(); var first = SubtotalCommands.Preview(Book(sheet), sheet.Id, new(0, 0, 8, 1), 0, [1]); using var session = new DocumentSession(first.Result);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, first.Groups[0].SummaryRow, 3), CellContent.FromLiteral(CellValue.FromText("keep")))]));
        Assert.Throws<NotSupportedException>(() => SubtotalCommands.Remove(session.Snapshot, sheet.Id, first.Range));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SubtotalsPreserveConditionalFormatsAndCrossSheetReferences(bool below)
    {
        var sheet = Records();
        var rule = new ConditionalFormatRule([new(1, 1, 8, 1)], ConditionalFormatKind.Expression, 1, new(FillColor: "FFFFC7CE"), "=B2=1");
        sheet = sheet.WithData(new() { ConditionalFormats = [rule] });
        var other = Sheet("References", (0, 0, F("=S!B9")));
        var preview = SubtotalCommands.Preview(Book(sheet, other), sheet.Id, new(0, 0, 8, 1), 0, [1], options: new(SummaryBelow: below));
        using var session = new DocumentSession(preview.Result); var calc = new WorkbookCalculator(new NoScripts());
        Assert.True(session.TryCommitCalculation(calc.Calculate(session.BeginCalculation())));
        Assert.Equal(8, session.Snapshot.Sheets[1].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("FFFFC7CE", ConditionalPresentation.GetStyle(session.Snapshot, session.Snapshot.Sheets[0], preview.Groups[0].FirstRow, 1).FillColor);
        var removed = SubtotalCommands.Remove(session.Snapshot, sheet.Id, preview.Range);
        var restored = removed.Result.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal(rule.Ranges.ToArray(), restored.Ranges.ToArray()); Assert.Equal("FFFFC7CE", restored.Style.FillColor);
        Assert.Contains("B9", removed.Result.Sheets[1].GetCell(0, 0)!.Content.FormulaText);
    }

    [Fact]
    public void ForecastCopiesDateFormatAndCannotApplyAStalePreview()
    {
        var sheet = Sheet("S", (0, 0, N(45000)), (0, 1, N(4)), (1, 0, N(45001)), (1, 1, N(6)), (2, 0, N(45002)), (2, 1, N(8)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(NumberFormatId: 14)));
        var p = DataAnalysis.Forecast(session.Snapshot, sheet.Id, new(0, 0, 2, 1), false, 1, new(0, 3));
        Assert.Equal(14, p.Result.Styles[p.Result.Sheets[0].GetStyleId(4, 3)].NumberFormatId);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 1), CellContent.FromLiteral(CellValue.FromNumber(5)))]));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplyCleaningCommand(p)));
        Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 3));
    }
}
