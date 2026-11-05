using QuickExcel.Calculation;
using QuickExcel.Domain;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DataCleaningTests
{
    [Fact]
    public void GeneralSplitRecognizesSafeNumbersAndYearFirstDatesAndKeepsOtherFieldsLiteral()
    {
        var sheet = Sheet("Data", (0, 0, T("001;007;02/01/2026;2026-01-02;1.234,5-;=1+1;1234567890123456;omit")));
        var book = Book(sheet);
        var formats = Enumerable.Repeat(CleaningColumnType.General, 8).ToArray();
        formats[1] = CleaningColumnType.Text;
        formats[7] = CleaningColumnType.Skip;
        var preview = DataCleaning.SplitText(book, sheet.Id, new(0, 0, 0, 0), new(0, 1), new(";"), formats, new(',', '.'));
        var result = preview.Result.Sheets[0];
        Assert.Equal(CellValue.FromNumber(1), result.GetCell(0, 1)!.Value);
        Assert.Equal(CellValue.FromText("007"), result.GetCell(0, 2)!.Value);
        Assert.Equal(CellValue.FromText("02/01/2026"), result.GetCell(0, 3)!.Value);
        Assert.Equal(CellValue.FromNumber(46024), result.GetCell(0, 4)!.Value);
        Assert.Equal(CellValue.FromNumber(-1234.5), result.GetCell(0, 5)!.Value);
        Assert.Equal(CellValue.FromText("=1+1"), result.GetCell(0, 6)!.Content.LiteralValue);
        Assert.Equal(CellValue.FromText("1234567890123456"), result.GetCell(0, 7)!.Value);
        Assert.Null(result.GetCell(0, 8));
    }

    [Fact]
    public void NavigationDoesNotExpirePreviewOrGetRewoundByCommit()
    {
        var sheet = Sheet("Data", (0, 0, T("001,2"))); using var session = new DocumentSession(Book(sheet));
        var preview = DataCleaning.SplitText(session.Snapshot, sheet.Id, new(0, 0, 0, 0), new(0, 1), new());
        session.UpdateNavigation(sheet.Id, new(4, 3), new(2, 1));
        var view = session.Snapshot.Sheets[0].View;
        session.Execute(new ApplyCleaningCommand(preview));
        Assert.Equal(view, session.Snapshot.Sheets[0].View);
        Assert.Equal("001", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
    }
    [Theory]
    [InlineData(ExcelDateSystem.Windows1900, "02/01/2026", CleaningColumnType.DateDmy, 46024)]
    [InlineData(ExcelDateSystem.Windows1900, "02/01/2026", CleaningColumnType.DateMdy, 46054)]
    [InlineData(ExcelDateSystem.Windows1900, "1900-02-29", CleaningColumnType.DateYmd, 60)]
    [InlineData(ExcelDateSystem.Mac1904, "2026-10-31", CleaningColumnType.DateYmd, 44864)]
    public void DateOrderAndDateSystemAreExplicit(ExcelDateSystem system, string input, CleaningColumnType type, double serial)
    {
        var sheet = Sheet("Data", (0, 0, T(input)));
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, system, [sheet]);
        var preview = DataCleaning.SplitText(book, sheet.Id, new(0, 0, 0, 0), new(0, 1), new(), [type]);
        Assert.Equal(serial, preview.Result.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
    }
    [Fact]
    public void PreviewDoesNotWriteAndCommitIsOneUndoableRevisionBoundTransaction()
    {
        var sheet = Sheet("Data", (0, 0, T("001,2.50,2026-10-31")), (0, 3, T("existing")));
        using var session = new DocumentSession(Book(sheet));
        var before = session.Snapshot;
        var preview = DataCleaning.SplitText(before, sheet.Id, new(0, 0, 0, 0), new(0, 2), new(), [CleaningColumnType.Text, CleaningColumnType.Number, CleaningColumnType.DateYmd]);
        Assert.Same(before, session.Snapshot); Assert.Single(preview.Overwrites);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplyCleaningCommand(preview)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new ApplyCleaningCommand(preview, true));
        Assert.Equal(CellValue.FromText("001"), session.Snapshot.Sheets[0].GetCell(0, 2)!.Value);
        Assert.Equal(CellValue.FromNumber(2.5), session.Snapshot.Sheets[0].GetCell(0, 3)!.Value);
        Assert.Equal(CellValue.FromNumber(46326), session.Snapshot.Sheets[0].GetCell(0, 4)!.Value);
        session.Undo(); Assert.Equal("existing", session.Snapshot.Sheets[0].GetCell(0, 3)!.Value.TextValue); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 2));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ApplyCleaningCommand(preview, true)));
        session.Redo(); Assert.Equal(CellValue.FromNumber(2.5), session.Snapshot.Sheets[0].GetCell(0, 3)!.Value);
    }
    [Fact]
    public void TextIsDefaultAndQuotedFieldsCannotBecomeFormulasOrLoseLeadingZeros()
    {
        var sheet = Sheet("Data", (0, 0, T("\"A,B\",001,=1+1,2026-10-31")), (1, 0, T("a,,\"say \"\"Hi\"\"\",z")));
        var book = Book(sheet); var preview = DataCleaning.SplitText(book, sheet.Id, new(0, 0, 1, 0), new(0, 1), new());
        var result = preview.Result.Sheets[0];
        Assert.Equal("A,B", result.GetCell(0, 1)!.Value.TextValue); Assert.Equal("001", result.GetCell(0, 2)!.Value.TextValue);
        Assert.Equal(CellContentKind.Literal, result.GetCell(0, 3)!.Content.Kind); Assert.Equal("=1+1", result.GetCell(0, 3)!.Value.TextValue);
        Assert.Equal("2026-10-31", result.GetCell(0, 4)!.Value.TextValue); Assert.Equal("say \"Hi\"", result.GetCell(1, 3)!.Value.TextValue);
        var bad = Book(Sheet("Bad", (0, 0, T("a,b")), (1, 0, T("\"unclosed,b"))));
        Assert.Throws<ArgumentException>(() => DataCleaning.SplitText(bad, bad.Sheets[0].Id, new(0, 0, 1, 0), new(0, 1), new()));
    }
    [Fact]
    public void FixedWidthNeverSplitsSurrogatePairsAndMatchesChineseExcelWidthUnits()
    {
        var sheet = Sheet("Data", (0, 0, T("中甲乙中文汉字尾")), (1, 0, T("a😀bcdefg")));
        var preview = DataCleaning.SplitText(Book(sheet), sheet.Id, new(0, 0, 1, 0), new(0, 1), new(Breaks: [3, 7]));
        Assert.Equal("中甲", preview.Result.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        Assert.Equal("乙中", preview.Result.Sheets[0].GetCell(0, 2)!.Value.TextValue);
        Assert.Equal("a😀b", preview.Result.Sheets[0].GetCell(1, 1)!.Value.TextValue);
        Assert.Equal("cdef", preview.Result.Sheets[0].GetCell(1, 2)!.Value.TextValue);
    }
    [Fact]
    public void ExplicitNumericConversionRejectsPrecisionLossDatesAndMalformedGroupingAtomically()
    {
        foreach (var bad in new[] { "1234567890123456", "2026-10-31", "1,2", "TRUE", "Infinity", "1e999", "1e-999", "1e-309" })
        {
            var sheet = Sheet("Data", (0, 0, T("001")), (1, 0, T(bad))); var book = Book(sheet);
            Assert.Throws<ArgumentException>(() => DataCleaning.ConvertTextNumbers(book, sheet.Id, new(0, 0, 1, 0)));
            Assert.Equal("001", book.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        }
        var valid = Sheet("Data", (0, 0, T("1.234,50")), (1, 0, T("12-")), (2, 0, T("50%")), (3, 0, F("=1+1")));
        var preview = DataCleaning.ConvertTextNumbers(Book(valid), valid.Id, new(0, 0, 3, 0), new(',', '.'));
        Assert.Equal(CellValue.FromNumber(1234.5), preview.Result.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(-12), preview.Result.Sheets[0].GetCell(1, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(0.5), preview.Result.Sheets[0].GetCell(2, 0)!.Value);
        Assert.Equal("=1+1", preview.Result.Sheets[0].GetCell(3, 0)!.Content.FormulaText);
    }
    [Fact]
    public void DuplicateRemovalKeepsFirstTypedDisplayKeyAndCopiesFormulaReferencesWithoutMovingOutsideCells()
    {
        var sheet = Sheet("Data", (0, 0, T("Key")), (1, 0, T("a")), (2, 0, T("A")), (3, 0, N(1)), (4, 0, T("1")), (5, 0, N(1)),
            (1, 1, F("=A2")), (2, 1, F("=A3")), (3, 1, F("=A4")), (4, 1, F("=A5")), (5, 1, F("=A6")), (0, 3, F("=A6")));
        var book = Book(sheet); var preview = DataCleaning.RemoveDuplicates(book, sheet.Id, new(0, 0, 5, 1), [0]);
        Assert.Equal(new[] { 2, 5 }, preview.RemovedRows);
        var result = preview.Result.Sheets[0];
        Assert.Equal(CellValue.FromNumber(1), result.GetCell(2, 0)!.Value); Assert.Equal(CellValue.FromText("1"), result.GetCell(3, 0)!.Value);
        Assert.Equal("=A3", result.GetCell(2, 1)!.Content.FormulaText);
        Assert.Equal("=A6", result.GetCell(0, 3)!.Content.FormulaText); Assert.Null(result.GetCell(4, 0));
        var unchanged = DataCleaning.RemoveDuplicates(book, sheet.Id, new(3, 0, 4, 0), [0], false);
        Assert.Empty(unchanged.Changes);
    }
    [Fact]
    public void ExpandToWholeRowsMovesAdjacentColumnsWhileKeysStayInTheSelectedColumns()
    {
        var sheet = Sheet("Data", (0, 0, T("Key")), (1, 0, T("a")), (2, 0, T("a")), (3, 0, T("b")),
            (0, 1, T("Record")), (1, 1, N(1)), (2, 1, N(2)), (3, 1, N(3)),
            (0, 2, T("Extra")), (1, 2, T("first")), (2, 2, T("duplicate")), (3, 2, T("third")));
        var book = Book(sheet);
        var compact = DataCleaning.RemoveDuplicates(book, sheet.Id, new(0, 0, 3, 1), [0]);
        Assert.Equal("duplicate", compact.Result.Sheets[0].GetCell(2, 2)!.Value.TextValue);
        var expanded = DataCleaning.RemoveDuplicates(book, sheet.Id, new(0, 0, 3, 1), [0], expandToWholeRows: true);
        Assert.Equal(new[] { 2 }, expanded.RemovedRows);
        Assert.Contains(expanded.Notices, notice => notice.Contains("扩展列"));
        var result = expanded.Result.Sheets[0];
        Assert.Equal(CellValue.FromText("b"), result.GetCell(2, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(3), result.GetCell(2, 1)!.Value);
        Assert.Equal(CellValue.FromText("third"), result.GetCell(2, 2)!.Value);
        Assert.Null(result.GetCell(3, 0)); Assert.Null(result.GetCell(3, 1)); Assert.Null(result.GetCell(3, 2));
    }
    [Fact]
    public void ExpandToWholeRowsStopsAtBlankColumnsAndRejectsUnsafeAdjacentContent()
    {
        var sheet = Sheet("Data", (0, 0, T("Key")), (1, 0, T("a")), (2, 0, T("a")),
            (0, 1, T("Record")), (1, 1, N(1)), (2, 1, N(2)), (0, 3, T("Gap")), (1, 3, T("keep")), (2, 3, T("keep")));
        var book = Book(sheet);
        var preview = DataCleaning.RemoveDuplicates(book, sheet.Id, new(0, 0, 2, 1), [0], expandToWholeRows: true);
        Assert.Equal("keep", preview.Result.Sheets[0].GetCell(2, 3)!.Value.TextValue);
        var rich = sheet.WithCells(sheet.Cells.SetItem(new(1, 2), new(CellContent.FromRichText(new([new("x")])))));
        Assert.Throws<NotSupportedException>(() => DataCleaning.RemoveDuplicates(Book(rich), rich.Id, new(0, 0, 2, 1), [0], expandToWholeRows: true));
    }
    [Fact]
    public void DuplicateCompactionAdjustsOnlyFormulaReferencesThatMoveWithTheRecord()
    {
        var sheet = Sheet("Data", (0, 0, T("Key")), (1, 0, T("a")), (2, 0, T("a")), (3, 0, T("b")),
            (1, 1, F("=D2")), (3, 1, F("=A4+D4")), (1, 3, N(10)), (2, 3, N(20)), (3, 3, N(30)));
        var preview = DataCleaning.RemoveDuplicates(Book(sheet), sheet.Id, new(0, 0, 3, 1), [0], expandToWholeRows: true);
        var result = preview.Result.Sheets[0];
        Assert.Equal("=D2", result.GetCell(1, 1)!.Content.FormulaText);
        Assert.Equal("=(A3+D4)", result.GetCell(2, 1)!.Content.FormulaText);
        var bad = Sheet("Bad", (0, 0, T("Key")), (1, 0, T("a")), (2, 0, T("a")), (3, 0, T("b")), (3, 1, F("=A3")));
        Assert.Throws<NotSupportedException>(() => DataCleaning.RemoveDuplicates(Book(bad), bad.Id, new(0, 0, 3, 1), [0]));
    }
    [Fact]
    public void ValidationFailureRejectsTheWholeReviewedCandidateAndHiddenRowsCannotBeCompressed()
    {
        var sheet = Sheet("Data", (0, 0, T("1,9")));
        sheet = sheet.WithData(new() { Validations = [new([new(0, 2, 0, 2)], ValidationKind.Whole, "=1", "=5")] });
        var book = Book(sheet); using var session = new DocumentSession(book, validation: new WorkbookValidation());
        var preview = DataCleaning.SplitText(book, sheet.Id, new(0, 0, 0, 0), new(0, 1), new(), [CleaningColumnType.Number, CleaningColumnType.Number]);
        Assert.Throws<DataValidationException>(() => session.Execute(new ApplyCleaningCommand(preview)));
        Assert.Same(book, session.Snapshot); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 1));
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 0, 1, hidden: true));
        Assert.Throws<InvalidOperationException>(() => DataCleaning.RemoveDuplicates(session.Snapshot, sheet.Id, new(0, 0, 0, 0), [0], false));
    }
}
