using System.IO.Compression;
using System.Text;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DataValidationTests
{
    private static DocumentSession Session(WorkbookSnapshot book) => new(book, validation: new WorkbookValidation());
    private static CellEdit Edit(SheetId id, int row, int column, double number) => new(new(id, row, column), CellContent.FromLiteral(CellValue.FromNumber(number)));

    [Fact]
    public void CandidateTransactionAndDependentFormulasAreCheckedBeforeAtomicCommit()
    {
        var sheet = Sheet("Data", (0, 0, N(1)), (0, 1, N(2)), (0, 2, F("=B1*2")), (1, 0, N(1)), (1, 1, N(2)));
        var rule = new WorksheetValidation([new(0, 0, 1, 0)], ValidationKind.Custom, "=AND(ISNUMBER(A1),A1<=B1)");
        using var session = Session(Book(sheet.WithData(new() { Validations = [rule] })));
        session.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 7), Edit(sheet.Id, 0, 1, 8)]));
        var before = session.Snapshot;
        Assert.Throws<DataValidationException>(() => session.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 9), Edit(sheet.Id, 1, 0, 3)])));
        Assert.Same(before, session.Snapshot);
        session.Execute(new SetValidationCommand(sheet.Id, new([new(0, 0, 1, 0)], ValidationKind.Custom, "=A1<C1")));
        session.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 17), Edit(sheet.Id, 0, 1, 9)]));
        Assert.Equal(CellValue.FromNumber(17), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        session.Undo(); Assert.Equal(CellValue.FromNumber(7), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void StopCannotBeOverriddenAndWarningsRequireAnExplicitWholeTransactionDecision()
    {
        var sheet = Sheet("Data"); var warning = new WorksheetValidation([new(0, 0, 0, 0)], ValidationKind.Whole, "=1", "=5", alert: ValidationAlert.Warning);
        var stop = new WorksheetValidation([new(1, 0, 1, 0)], ValidationKind.Whole, "=1", "=5");
        using var session = Session(Book(sheet.WithData(new() { Validations = [warning, stop] })));
        var command = new SetCellsCommand([Edit(sheet.Id, 0, 0, 9)]);
        Assert.True(Assert.Throws<DataValidationException>(() => session.Execute(command)).CanOverride);
        Assert.False(session.IsDirty); session.Execute(command, acceptValidationWarnings: true); Assert.True(session.CanUndo);
        var before = session.Snapshot;
        Assert.False(Assert.Throws<DataValidationException>(() => session.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 10), Edit(sheet.Id, 1, 0, 9)]), acceptValidationWarnings: true)).CanOverride);
        Assert.Same(before, session.Snapshot); session.Undo(); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 0));
        session.Execute(new SetValidationCommand(sheet.Id, new([new(0, 0, 0, 0)], ValidationKind.Custom, "=1/0", alert: ValidationAlert.Warning)));
        Assert.False(Assert.Throws<DataValidationException>(() => session.Execute(command, acceptValidationWarnings: true)).CanOverride);
    }

    [Fact]
    public void ClearPasteAndAbsentValidationEngineCannotBypassStopRules()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (1, 0, N(3))); var rule = new WorksheetValidation([new(0, 0, 1, 0)], ValidationKind.Whole, "=1", "=5", allowBlank: false);
        var book = Book(sheet.WithData(new() { Validations = [rule] }));
        using var session = Session(book);
        Assert.Throws<DataValidationException>(() => session.Execute(new ClearCellsCommand([new(sheet.Id, 0, 0)])));
        using var unconfigured = new DocumentSession(book);
        Assert.Throws<NotSupportedException>(() => unconfigured.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 4)])));
        Assert.Same(book, session.Snapshot); Assert.Same(book, unconfigured.Snapshot);
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1)));
        var source = session.Snapshot.WithSheets([sheet.WithCells(sheet.Cells.SetItem(new(0, 2), N(9))).WithData(new() { Validations = [rule] })]);
        using var pasteSession = Session(source);
        var clipboard = WorkbookClipboard.Capture(source, sheet.Id, new(0, 2, 0, 2));
        Assert.Throws<DataValidationException>(() => pasteSession.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(0, 0, 1, 0), source.Revision, CellPasteMode.Values)));
        Assert.Same(source, pasteSession.Snapshot);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void NormalSavePreservesRulesAndRemovalIsUndoable(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-validation-" + Guid.NewGuid().ToString("N") + "." + extension);
        try
        {
            var sheet = Sheet("Data", (0, 0, T("001"))); var rule = new WorksheetValidation([new(0, 0, 3, 0), new(5, 0, 6, 0)], ValidationKind.List,
                "=\"001,002\"", inputTitle: "ID", inputMessage: "保留前导零", errorMessage: "从列表中选择");
            using var session = Session(Book(sheet.WithData(new() { Validations = [rule] })));
            new TypedWorkbookStorage().Save(path, session.Snapshot);
            using var reopened = Session(new TypedWorkbookStorage().Open(path)); Assert.Empty(reopened.Snapshot.FileState!.CompatibilityIssues);
            Assert.True(rule.ContentEquals(reopened.Snapshot.Sheets[0].Data.Validations.Single()));
            var id = reopened.Snapshot.Sheets[0].Id;
            Assert.Throws<DataValidationException>(() => reopened.Execute(new SetCellsCommand([Edit(id, 0, 0, 1)])));
            reopened.Execute(new RemoveValidationCommand(id, new(0, 0, 0, 0))); Assert.Empty(reopened.Snapshot.Sheets[0].Data.Validations);
            reopened.Undo(); Assert.True(rule.ContentEquals(reopened.Snapshot.Sheets[0].Data.Validations.Single()));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void NamesAndSheetCopyTransformValidationReferencesAndKeepTheOriginalIndependent()
    {
        var sheet = Sheet("Data", (0, 1, N(4))); var rule = new WorksheetValidation([new(0, 0, 0, 0)], ValidationKind.Whole, "=Limit", comparison: ValidationOperator.LessThanOrEqual);
        using var session = Session(Book(sheet.WithData(new() { Validations = [rule] })).WithDefinedNames([new("Limit", "=Data!$B$1")]));
        session.Execute(new RenameDefinedNameCommand("Limit", null, "Maximum"));
        Assert.Equal("=Maximum", session.Snapshot.Sheets[0].Data.Validations[0].Formula1);
        var copy = new CopySheetCommand(sheet.Id, "Copy", 1); session.Execute(copy);
        session.Execute(new SetCellsCommand([Edit(copy.Id, 0, 1, 8), Edit(copy.Id, 0, 0, 7)]));
        Assert.Throws<DataValidationException>(() => session.Execute(new SetCellsCommand([Edit(sheet.Id, 0, 0, 7)])));
    }

    [Fact]
    public void ExplicitDateRuleControlsIsoInputWithoutReinterpretingUnconstrainedOrQuotedText()
    {
        var sheet = Sheet("Data").WithData(new() { Validations = [new([new(0, 0, 0, 0)], ValidationKind.Date, "=DATE(2026,10,1)", "=DATE(2026,10,31)")] });
        foreach (var system in new[] { ExcelDateSystem.Windows1900, ExcelDateSystem.Mac1904 })
        {
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, system, [sheet]);
            Assert.Equal(CellValue.FromNumber(system == ExcelDateSystem.Windows1900 ? 46310 : 44848), CellInput.ParseForCell("2026-10-15", book, sheet, 0, 0).LiteralValue);
            Assert.Equal(CellValue.FromText("2026-10-15"), CellInput.ParseForCell("'2026-10-15", book, sheet, 0, 0).LiteralValue);
            Assert.Equal(CellValue.FromText("2026-10-15"), CellInput.ParseForCell("2026-10-15", book, sheet, 0, 1).LiteralValue);
        }
    }

    [Fact]
    public void UnsupportedValidationIsPreservedAndBlocksInputInsteadOfDroppingTheRestriction()
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-validation-unsupported-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var saved = path + ".copy.xlsx";
        try
        {
            var sheet = Sheet("Data"); var rule = new WorksheetValidation([new(0, 0, 0, 0)], ValidationKind.Decimal, "=0", "=1");
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(sheet.WithData(new() { Validations = [rule] })));
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; string xml;
                using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
                entry.Delete(); using var writer = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open(), new UTF8Encoding(false));
                writer.Write(xml.Replace("type=\"decimal\"", "type=\"time\""));
            }
            var original = File.ReadAllBytes(path); var book = storage.Open(path);
            Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/dataValidation");
            using var session = Session(book);
            Assert.Throws<NotSupportedException>(() => session.Execute(new SetCellsCommand([Edit(book.Sheets[0].Id, 0, 0, 0.5)])));
            storage.Save(saved, book); Assert.Equal(original, File.ReadAllBytes(saved));
        }
        finally { File.Delete(path); File.Delete(saved); }
    }

    [Theory]
    [InlineData("excel-data-validation")]
    [InlineData("excel-data-validation-1904")]
    public void NativeValidationFixtureMatchesEveryExcelDecision(string fixtureName)
    {
        using var fixture = ScalarFormulaFixture.Load(fixtureName); var root = fixture.RootElement;
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-validation-fixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in root.GetProperty("Parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
            var book = new TypedWorkbookStorage().Open(path); var validation = new WorkbookValidation();
            foreach (var record in root.GetProperty("Cases").EnumerateArray())
            {
                var sheet = book.Sheets.Single(s => s.Name == record.GetProperty("Sheet").GetString()); var range = CellRectangle.Parse(record.GetProperty("Address").GetString()!);
                var valid = validation.Check(book, [new(sheet.Id, range.Top, range.Left)]).Count == 0;
                Assert.True(valid == record.GetProperty("Valid").GetBoolean(), record.GetProperty("Case").GetString());
            }
        }
        finally { File.Delete(path); }
    }
}
