using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class WorkbookDataTransferTests
{
    [Fact]
    public void ExtractedRangeRebasesValuesPreservesNumbersAndDropsExecutableSource()
    {
        var sheet = Sheet("Source", (4, 2, Text("00123")), (5, 3, Number(12.5, 1)),
            (5, 4, Text("=1+1")), (1, 1, Text("outside")));
        var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904, [sheet], "throw new Error('must not execute')",
            [new(), new() { NumberFormatCode = "0.00", NumberFormatId = 164, Bold = true }]);
        var result = WorkbookDataTransfer.ExtractValues(source, [new(sheet.Id, "Imported", new(4, 2, 6, 5))]);
        var actual = result.Workbook.Sheets.Single();
        Assert.Equal("Imported", actual.Name);
        Assert.Equal("00123", actual.GetCell(0, 0)!.Value.TextValue);
        Assert.Equal(12.5, actual.GetCell(1, 1)!.Value.NumberValue);
        Assert.Equal("=1+1", actual.GetCell(1, 2)!.Value.TextValue);
        Assert.Equal(CellContentKind.Literal, actual.GetCell(1, 2)!.Content.Kind);
        Assert.Equal((3, 4), (actual.UsedRowCount, actual.UsedColumnCount));
        Assert.Equal("0.00", result.Workbook.Styles[actual.GetStyleId(1, 1)].NumberFormatCode);
        Assert.False(result.Workbook.Styles[actual.GetStyleId(1, 1)].Bold);
        Assert.Equal(ExcelDateSystem.Mac1904, result.Workbook.DateSystem);
        Assert.Empty(result.Workbook.ScriptSource);
        Assert.Equal(4, result.Workbook.ScriptModules.ScriptApiVersion);
        Assert.Null(result.Workbook.FileState);
        Assert.NotEqual(source.DocumentId, result.Workbook.DocumentId);
        Assert.NotEqual(sheet.Id, actual.Id);
        Assert.Contains(result.Notices, n => n.Contains("公式转为结果"));
        Assert.Equal(4, source.Sheets[0].Cells.Count);
    }

    [Theory]
    [InlineData(CalculationState.Dirty, 0)]
    [InlineData(CalculationState.Stale, 0)]
    [InlineData(CalculationState.Blocked, 0)]
    [InlineData(CalculationState.Fresh, 1)]
    public void StaleFormulaCachesRequireExplicitOptIn(CalculationState state, long revision)
    {
        var sheet = Sheet("S", (0, 0, new(CellContent.FromFormula("=1+1"),
            new(CellValue.FromNumber(2), state, revision, 0, 0))));
        var source = Book(sheet);
        Assert.Throws<InvalidOperationException>(() => WorkbookDataTransfer.ExtractValues(source, [new(sheet.Id)]));
        var accepted = WorkbookDataTransfer.ExtractValues(source, [new(sheet.Id)], allowStaleCaches: true);
        Assert.Equal(2, accepted.Workbook.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Contains(accepted.Notices, n => n.Contains("可能过期"));
    }

    [Fact]
    public void MissingCacheNeverBecomesBlankButPresentBlankIsRetained()
    {
        var missing = Sheet("S", (0, 0, new(CellContent.FromFormula("=1+1"))));
        Assert.Throws<InvalidOperationException>(() => WorkbookDataTransfer.ExtractValues(Book(missing), [new(missing.Id)], true));
        var blank = Sheet("S", (0, 0, new(CellContent.FromFormula("=A2"), new(CellValue.Blank, CalculationState.Fresh, 0, 0, 0))));
        Assert.Equal(CellValue.Blank, WorkbookDataTransfer.ExtractValues(Book(blank), [new(blank.Id)]).Workbook.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void SpillChildrenAreCopiedAsIndependentValuesAndValidateTheirCaches()
    {
        var id = SheetId.New();
        var child = new CellSnapshot(CellContent.FromLiteral(CellValue.Blank), new(CellValue.FromNumber(7), CalculationState.Fresh, 0, 0, 0), new(id, 0, 0));
        var sheet = new WorksheetSnapshot(id, "Array", new Dictionary<CellPosition, CellSnapshot> { [new(1, 0)] = child });
        var result = WorkbookDataTransfer.ExtractValues(Book(sheet), [new(id, Range: new(1, 0, 1, 0))]);
        Assert.Equal(CellValue.FromNumber(7), result.Workbook.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Null(result.Workbook.Sheets[0].GetCell(0, 0)!.SpillOwner);
        Assert.Null(result.Workbook.Sheets[0].GetCell(0, 0)!.Calculation);
    }

    [Fact]
    public void ExplicitValuesProjectionFlattensRichTextWithVisibleNoticeAndDoesNotMutateSource()
    {
        var text = new RichCellText([new("A", new(Bold: true)), new("B")]);
        var sheet = Sheet("Rich", (0, 0, new(CellContent.FromRichText(text))));
        var source = Book(sheet);
        Assert.Throws<ArgumentException>(() => new ImportWorksheetDataCommand(source));
        var result = WorkbookDataTransfer.ExtractValues(source, [new(sheet.Id)]);
        Assert.Equal("AB", result.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Null(result.Workbook.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.Same(text, source.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.Contains(result.Notices, n => n.Contains("富文本转为纯文本"));
    }

    [Fact]
    public void BlankInheritedNumberFormatsAreRetainedWithoutCopyingRowColumnLayout()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S", columns: new Dictionary<int, ColumnMetadata>
        { [2] = new(Width: 30, Hidden: true, StyleId: 1) });
        var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { NumberFormatId = 14 }]);
        var extracted = WorkbookDataTransfer.ExtractValues(source, [new(sheet.Id, Range: new(2, 2, 3, 3))]).Workbook;
        Assert.Equal(14, extracted.Styles[extracted.Sheets[0].GetStyleId(0, 0)].NumberFormatId);
        Assert.Equal(14, extracted.Styles[extracted.Sheets[0].GetStyleId(1, 0)].NumberFormatId);
        Assert.Equal(0, extracted.Sheets[0].GetStyleId(0, 1));
        Assert.Empty(extracted.Sheets[0].Columns);
        Assert.Empty(extracted.Sheets[0].Rows);
        using var session = new DocumentSession(Book(Sheet("Target")));
        session.Execute(new ImportWorksheetDataCommand(extracted));
        Assert.Equal(14, session.Snapshot.Styles[session.Snapshot.Sheets[1].GetStyleId(0, 0)].NumberFormatId);
    }

    [Fact]
    public void NameConflictPlanIsCaseInsensitiveAndReservesEarlierImportedNames()
    {
        var target = Book(Sheet("Data"));
        var source = Book(Sheet("data"), Sheet("data (2)"));
        var rename = WorkbookDataTransfer.PlanImport(target, source, SheetNameConflictPolicy.Rename);
        Assert.Equal(new[] { "data (2)", "data (2) (2)" }, rename.Select(p => p.TargetName));
        var skip = WorkbookDataTransfer.PlanImport(target, source, SheetNameConflictPolicy.Skip);
        Assert.True(skip[0].Skipped);
        Assert.Null(skip[0].TargetName);
        Assert.False(skip[1].Skipped);
        Assert.Throws<InvalidOperationException>(() => WorkbookDataTransfer.PlanImport(target, source, SheetNameConflictPolicy.Fail));
    }

    [Fact]
    public void LongNamesStayWithinExcelLimitWhenRenamed()
    {
        var name = new string('中', 31);
        var plan = WorkbookDataTransfer.PlanImport(Book(Sheet(name)), Book(Sheet(name)), SheetNameConflictPolicy.Rename);
        Assert.Equal(31, plan[0].TargetName!.Length);
        Assert.EndsWith(" (2)", plan[0].TargetName);
        WorkbookLimits.ValidateSheetName(plan[0].TargetName!);
    }

    [Fact]
    public void MultiSheetImportIsOneUndoTransactionAndKeepsExistingIdentityContentAndMetadata()
    {
        var original = Sheet("Data", (0, 0, Text("existing")));
        var target = Book(original);
        var source = Book(Sheet("data", (0, 0, Text("imported"))), Sheet("Other", (1, 1, Number(5))));
        using var session = new DocumentSession(target);
        session.Execute(new ImportWorksheetDataCommand(source));
        Assert.Equal(new[] { "Data", "data (2)", "Other" }, session.Snapshot.Sheets.Select(s => s.Name));
        Assert.Same(original, session.Snapshot.Sheets[0]);
        Assert.Equal(target.DocumentId, session.Snapshot.DocumentId);
        Assert.True(session.HistoryBytes > 1024);
        Assert.Equal("imported", session.Snapshot.Sheets[1].GetCell(0, 0)!.Value.TextValue);
        var importedId = session.Snapshot.Sheets[1].Id;
        Assert.NotEqual(source.Sheets[0].Id, importedId);
        Assert.True(session.Undo());
        Assert.Single(session.Snapshot.Sheets);
        Assert.False(session.CanUndo);
        Assert.True(session.Redo());
        Assert.Equal(importedId, session.Snapshot.Sheets[1].Id);
    }

    [Fact]
    public void SkipAllIsNoOpAndReadOnlyTargetRejectsImport()
    {
        var book = Book(Sheet("Data", (0, 0, Text("old"))));
        using var session = new DocumentSession(book);
        Assert.False(session.Execute(new ImportWorksheetDataCommand(Book(Sheet("data")), SheetNameConflictPolicy.Skip)).Changed);
        Assert.Same(book, session.Snapshot);
        Assert.False(session.CanUndo);
        using var readOnly = new DocumentSession(book, readOnly: true);
        Assert.Throws<InvalidOperationException>(() => readOnly.Execute(new ImportWorksheetDataCommand(Book(Sheet("Other")))));
        Assert.Same(book, readOnly.Snapshot);
    }

    [Fact]
    public void DatesConvertBetweenEpochsButOrdinaryNumbersDoNot()
    {
        var sheet = Sheet("Dates", (0, 0, Number(0, 1)), (0, 1, Number(0)));
        var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904, [sheet], styles: [new(), new() { NumberFormatId = 14 }]);
        using var session = new DocumentSession(Book(Sheet("Target")));
        session.Execute(new ImportWorksheetDataCommand(source));
        Assert.Equal(1462, session.Snapshot.Sheets[1].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal(0, session.Snapshot.Sheets[1].GetCell(0, 1)!.Value.NumberValue);
        Assert.Equal(14, session.Snapshot.Styles[session.Snapshot.Sheets[1].GetStyleId(0, 0)].NumberFormatId);
    }

    [Theory]
    [InlineData("h:mm", 0.5)]
    [InlineData("[h]:mm", 2.5)]
    public void TimeOnlyAndDurationFormatsNeverShiftTheirEpoch(string format, double value)
    {
        var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904,
            [Sheet("Times", (0, 0, Number(value, 1)))], styles: [new(), new() { NumberFormatCode = format, NumberFormatId = 164 }]);
        using var session = new DocumentSession(Book(Sheet("Target")));
        session.Execute(new ImportWorksheetDataCommand(source));
        Assert.Equal(value, session.Snapshot.Sheets[1].GetCell(0, 0)!.Value.NumberValue);
    }

    [Fact]
    public void ImpossibleDateConversionRejectsEntireBatch()
    {
        var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [Sheet("Valid", (0, 0, Text("valid"))), Sheet("Dates", (0, 0, Number(60, 1)))],
            styles: [new(), new() { NumberFormatId = 14 }]);
        var target = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904, [Sheet("Target")]);
        using var session = new DocumentSession(target);
        Assert.Throws<NotSupportedException>(() => session.Execute(new ImportWorksheetDataCommand(source)));
        Assert.Same(target, session.Snapshot);
        Assert.False(session.CanUndo);
    }

    [Fact]
    public void InvalidSelectionsLimitsAndCancellationAreRejectedBeforeImport()
    {
        var sheet = Sheet("S");
        var book = Book(sheet);
        Assert.Throws<ArgumentException>(() => WorkbookDataTransfer.ExtractValues(book, []));
        Assert.Throws<ArgumentException>(() => WorkbookDataTransfer.ExtractValues(book, [new(sheet.Id), new(sheet.Id)]));
        Assert.Throws<ArgumentException>(() => WorkbookDataTransfer.ExtractValues(book, [new(sheet.Id, "bad/name")]));
        Assert.Throws<ArgumentException>(() => WorkbookDataTransfer.ExtractValues(book, [new(sheet.Id, Range: new(2, 0, 0, 0))]));
        Assert.Throws<InvalidOperationException>(() => WorkbookDataTransfer.ExtractValues(book, [new(sheet.Id, Range: new(0, 0, 1000, 1000))]));
        Assert.Throws<OperationCanceledException>(() => WorkbookDataTransfer.ExtractValues(book, [new(sheet.Id)], cancellation: new(true)));
        var full = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, Enumerable.Range(0, 64).Select(i => Sheet("S" + i)));
        using var session = new DocumentSession(full);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ImportWorksheetDataCommand(book)));
        Assert.Same(full, session.Snapshot);
    }

    [Fact]
    public void ExtractedSheetCanRoundTripThroughRealXlsx()
    {
        var sheet = Sheet("源数据", (0, 0, Text("00123")), (1, 0, Number(3.25)), (1, 1, Text("=1+1")));
        var extracted = WorkbookDataTransfer.ExtractValues(Book(sheet), [new(sheet.Id, "Selected")]).Workbook;
        var path = Path.Combine(Path.GetTempPath(), "quickexcel-transfer-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage();
            storage.Export(path, extracted);
            var opened = storage.Open(path);
            Assert.Equal("Selected", opened.Sheets.Single().Name);
            Assert.Equal("00123", opened.Sheets[0].GetCell(0, 0)!.Value.TextValue);
            Assert.Equal(3.25, opened.Sheets[0].GetCell(1, 0)!.Value.NumberValue);
            Assert.Equal("=1+1", opened.Sheets[0].GetCell(1, 1)!.Value.TextValue);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static WorkbookSnapshot Book(params WorksheetSnapshot[] sheets) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, sheets);
    private static WorksheetSnapshot Sheet(string name, params (int Row, int Column, CellSnapshot Cell)[] cells) =>
        new(SheetId.New(), name, cells.ToImmutableDictionary(c => new CellPosition(c.Row, c.Column), c => c.Cell));
    private static CellSnapshot Text(string text) => new(CellContent.FromLiteral(CellValue.FromText(text)));
    private static CellSnapshot Number(double number, int style = 0) => new(CellContent.FromLiteral(CellValue.FromNumber(number), style));
}
