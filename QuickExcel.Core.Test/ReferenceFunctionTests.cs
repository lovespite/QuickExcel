using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ReferenceFunctionTests
{
    [Fact]
    public void StructureChangesUndoAndSpillConflictsUseNormalCalculationHistory()
    {
        var sheet = Sheet("Data", (5, 3, F("=ROW()+ROWS(A1:A3)")), (0, 7, F("=ROW(B2:B4)")), (1, 7, T("occupied")));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Calculate(); Assert.Equal(CellValue.FromNumber(9), session.Snapshot.Sheets[0].GetCell(5, 3)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill), session.Snapshot.Sheets[0].GetCell(0, 7)!.Value);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1)); Calculate();
        Assert.Equal(CellValue.FromNumber(11), session.Snapshot.Sheets[0].GetCell(6, 3)!.Value);
        Assert.True(session.Undo()); Calculate(); Assert.Equal(CellValue.FromNumber(9), session.Snapshot.Sheets[0].GetCell(5, 3)!.Value);
    }
    [Fact]
    public void MatchesIndependentExcelResults()
    {
        using var fixture = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-common-functions.json")));
        foreach (var item in fixture.RootElement.GetProperty("cases").EnumerateArray())
        {
            var sheet = Sheet("Data", (5, 3, F(item.GetProperty("formula").GetString()!)));
            var result = C(new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet))), sheet, 5, 3);
            var actual = result.Array?.Values.ToArray() ?? [result.Value!.Value];
            var expected = item.GetProperty("values").EnumerateArray().ToArray();
            Assert.Equal(expected.Length, actual.Length);
            for (var i = 0; i < actual.Length; i++)
            {
                Assert.True(expected[i].GetProperty("kind").GetString() == actual[i].Kind.ToString(), item.GetProperty("formula").GetString() + " => " + actual[i]);
                if (actual[i].Kind == CellValueKind.Number) Assert.Equal(expected[i].GetProperty("value").GetDouble(), actual[i].NumberValue, 10);
                else Assert.Equal(expected[i].GetProperty("value").GetString(), CellInput.DisplayText(actual[i]));
            }
        }
    }
    [Theory]
    [InlineData("=ROW()", 6)]
    [InlineData("=COLUMN()", 4)]
    [InlineData("=ROW(D6)", 6)]
    [InlineData("=COLUMN(D6)", 4)]
    [InlineData("=ROWS(D6:E8)", 3)]
    [InlineData("=COLUMNS(D6:E8)", 2)]
    [InlineData("=ROWS(SEQUENCE(3,2))", 3)]
    [InlineData("=COLUMNS(SEQUENCE(3,2))", 2)]
    [InlineData("=ROWS(1)", 1)]
    [InlineData("=COLUMNS(\"text\")", 1)]
    public void CoordinatesDimensionsAndSelfReferencesDoNotReadCellValues(string formula, double expected)
    {
        var sheet = Sheet("Data", (5, 3, F(formula)));
        var book = Book(sheet); var index = new WorkbookDependencyIndex(book);
        Assert.Empty(index.CycleMembers);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(book));
        Assert.Equal(CellValue.FromNumber(expected), C(result, sheet, 5, 3).Value);
    }

    [Fact]
    public void ReferenceArraysNamesAndCrossSheetReferencesSpillWithOriginalCoordinates()
    {
        var data = Sheet("Data"); var sheet = Sheet("Results", (0, 0, F("=ROW(Named)")), (0, 3, F("=COLUMN(Data!B2:D3)")));
        var book = Book(data, sheet).WithDefinedNames([new("Named", "=Data!$B$2:$D$4")]);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(book));
        Assert.Equal(new double[] { 2, 3, 4 }, C(result, sheet).Array!.Values.Select(v => v.NumberValue));
        Assert.Equal((1, 3), (C(result, sheet, 0, 3).Array!.Rows, C(result, sheet, 0, 3).Array!.Columns));
        Assert.Equal(new double[] { 2, 3, 4 }, C(result, sheet, 0, 3).Array!.Values.Select(v => v.NumberValue));
    }

    [Theory]
    [InlineData("=ROW(5)", CellErrorCode.Value)]
    [InlineData("=ROWS()", CellErrorCode.Value)]
    [InlineData("=COLUMN(A1,B2)", CellErrorCode.Value)]
    [InlineData("=ROWS(#N/A)", CellErrorCode.NotAvailable)]
    [InlineData("=ROW(Missing!A1)", CellErrorCode.Reference)]
    public void InvalidArgumentsAndMissingTargetsPropagateErrors(string formula, CellErrorCode error)
    {
        var sheet = Sheet("Data", (0, 0, F(formula)));
        Assert.Equal(CellValue.FromError(error), C(new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet))), sheet).Value);
    }

    [Fact]
    public void SourceFunctionsSurviveStorageAndCompletion()
    {
        var sheet = Sheet("Data", (4, 3, F("=ROW()+COLUMN()+ROWS(A1:A3)")));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(sheet)); var book = storage.Open(path);
            Assert.False(book.FileState?.FormulaReadOnly == true);
            Assert.Equal("=ROW()+COLUMN()+ROWS(A1:A3)", book.Sheets[0].GetCell(4, 3)!.Content.FormulaText);
            Assert.Equal(CellValue.FromNumber(12), C(new WorkbookCalculator(new NoScripts()).Calculate(Request(book)), book.Sheets[0], 4, 3).Value);
            Assert.Contains("ROWS", FormulaEditing.Complete("=ROW", 4)!.Names);
        }
        finally { File.Delete(path); }
    }
}
