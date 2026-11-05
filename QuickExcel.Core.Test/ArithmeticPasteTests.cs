using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ArithmeticPasteTests
{
    private static ExternalClipboardRegion Source(params CellValue[] values) => new(values.Length, 1, values.Select(v => new ExternalClipboardCell(v)));
    [Fact]
    public void DestinationFormattingLinksAndNotesRemainAttached()
    {
        var source = Sheet("S", (0, 0, new(CellContent.FromLiteral(CellValue.FromNumber(8), 1))));
        var annotations = new WorksheetAnnotations(links: [new(new(0, 0), new("https://example.com"))], notes: [new(new(0, 0), new("Author", new([new("Keep")])))]);
        var sheet = source.WithAnnotations(annotations);
        using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]));
        session.Execute(new ArithmeticPasteCommand(Source(CellValue.FromNumber(2)), sheet.Id, new(0, 0, 0, 0), 0, PasteArithmetic.Add));
        Assert.True(session.Snapshot.Sheets[0].Annotations.ContentEquals(annotations)); Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(0, 0));
        Assert.Equal(10, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }
    [Fact]
    public void StaleSpillResultsRemainMarkedStaleAfterValueOnlyCapture()
    {
        var id = SheetId.New();
        var sheet = new WorksheetSnapshot(id, "S", [new(new(1, 0), new(CellContent.FromLiteral(CellValue.Blank),
            new(CellValue.FromNumber(2), CalculationState.Stale, 0, 0, 0), new(id, 0, 0)))]);
        Assert.True(WorkbookClipboard.Capture(Book(sheet), id, new(1, 0, 1, 0)).HasOldResults);
    }
    [Fact]
    public void IndependentExcelNumericBlankAndFormulaCasesMatchExceptExplicitZeroDivisionRejection()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-arithmetic-paste.json")));
        foreach (var sample in fixture.RootElement.GetProperty("cases").EnumerateArray())
        {
            var kind = sample.GetProperty("kind").GetString(); var operation = Enum.Parse<PasteArithmetic>(sample.GetProperty("operation").GetString()!, true);
            var sheet = Sheet("S", (0, 0, kind == "formula-target" ? F("=4+4") : kind == "blank-target" ? new(CellContent.FromLiteral(CellValue.Blank)) : N(8)));
            using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
            Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
            var value = kind is "blank-source" or "skip-blank" ? CellValue.Blank : CellValue.FromNumber(2);
            var source = new ExternalClipboardRegion(1, 1, [new(value, Formula: kind == "formula-source" ? "=1+1" : null)]);
            var command = new ArithmeticPasteCommand(source, sheet.Id, new(0, 0, 0, 0), session.Snapshot.Revision, operation, kind == "skip-blank");
            if (sample.GetProperty("display").GetString()!.StartsWith('#'))
            {
                var before = session.Snapshot; Assert.Throws<InvalidOperationException>(() => session.Execute(command)); Assert.Same(before, session.Snapshot); continue;
            }
            session.Execute(command); Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
            var result = session.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!;
            Assert.Equal(sample.GetProperty("value").GetDouble(), result.Value.NumberValue);
            Assert.Equal(sample.GetProperty("hasFormula").GetBoolean(), result.Content.Kind == CellContentKind.Formula);
        }
    }
    [Fact]
    public void FormulaReferencesRemainLiveAndBatchIsOneUndoablePersistedEdit()
    {
        var sheet = Sheet("S", (0, 0, N(3)), (0, 1, F("=A1*2")), (1, 1, N(8)));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Calculate(); session.Execute(new ArithmeticPasteCommand(Source(CellValue.FromNumber(2)), sheet.Id, new(0, 1, 1, 1), 0, PasteArithmetic.Multiply)); Calculate();
        Assert.Equal(12, session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue); Assert.Equal(16, session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue);
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "arithmetic-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        foreach (var ext in new[] { "xlsx", "xlsq" })
        { var path = Path.Combine(folder, "book." + ext); var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); Assert.Equal(12, storage.Open(path).Sheets[0].GetCell(0, 1)!.Value.NumberValue); }
        Assert.True(session.Undo()); Assert.Equal("=A1*2", session.Snapshot.Sheets[0].GetCell(0, 1)!.Content.FormulaText); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(4)))])); Calculate();
        Assert.Equal(16, session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
    }
    [Theory]
    [InlineData("text")] [InlineData("boolean")] [InlineData("error")] [InlineData("zero")] [InlineData("overflow")]
    [InlineData("protected")] [InlineData("validation")] [InlineData("stale")] [InlineData("revision")] [InlineData("array")] [InlineData("hidden")]
    public void FailuresLeaveEveryTargetAndHistoryUntouched(string kind)
    {
        var second = kind switch { "text" => T("2"), "boolean" => new(CellContent.FromLiteral(CellValue.FromBoolean(true))),
            "error" => new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Value))), "stale" => F("=2"),
            "array" => new(CellContent.FromFormula("=SEQUENCE(1)", dynamicArray: true)), "overflow" => N(double.MaxValue), _ => N(8) };
        var sourceSheet = Sheet("S", (0, 0, N(8)), (1, 0, second));
        var sheet = new WorksheetSnapshot(sourceSheet.Id, sourceSheet.Name, sourceSheet.Cells, rows: kind == "hidden" ? [new(1, new(Hidden: true))] : null,
            data: kind == "protected" ? new() { Protection = new() } : kind == "validation" ? new() { Validations = [new([new(1, 0, 1, 0)], ValidationKind.Whole, "=1", "=5")] } : null);
        var book = Book(sheet); using var session = new DocumentSession(book, validation: new WorkbookValidation());
        var source = Source(CellValue.FromNumber(2), CellValue.FromNumber(kind == "zero" ? 0 : 2));
        Assert.ThrowsAny<Exception>(() => session.Execute(new ArithmeticPasteCommand(source, sheet.Id, new(0, 0, 1, 0), kind == "revision" ? 1 : 0,
            kind == "zero" ? PasteArithmetic.Divide : PasteArithmetic.Multiply)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
    [Fact]
    public void SkipBlanksDoesNotTouchTextTargetsAndMissingSourceResultsNeverBecomeZero()
    {
        var sheet = Sheet("S", (0, 0, N(5)), (1, 0, T("keep"))); var book = Book(sheet); using var session = new DocumentSession(book);
        session.Execute(new ArithmeticPasteCommand(Source(CellValue.FromNumber(2), CellValue.Blank), sheet.Id, new(0, 0, 1, 0), 0, PasteArithmetic.Add));
        Assert.Equal(7, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue); Assert.Equal("keep", session.Snapshot.Sheets[0].GetCell(1, 0)!.Value.TextValue);
        var before = session.Snapshot; var missing = new ExternalClipboardRegion(1, 1, [new(CellValue.Blank, Formula: "=A1", MissingResult: true)]);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ArithmeticPasteCommand(missing, sheet.Id, new(0, 0, 0, 0), before.Revision, PasteArithmetic.Add)));
        Assert.Same(before, session.Snapshot);
    }
}
