using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DateMathFunctionTests
{
    [Theory]
    [InlineData("=YEAR(A1:A2)")]
    [InlineData("=ROUNDUP(A1:A2,0)")]
    [InlineData("=WORKDAY(45351,10,IF(TRUE,A1:A2,B1:B2))")]
    public void UnimplementedArraysKeepOriginalFileAndBlockedCache(string formula)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"); var copy = path + ".copy.xlsx";
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(Sheet("S", (0, 5, F(formula, 123)))));
            using var session = new DocumentSession(storage.Open(path)); Assert.True(session.Snapshot.FileState!.FormulaReadOnly);
            ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
            var cell = session.Snapshot.Sheets[0].GetCell(0, 5)!; Assert.Equal(123, cell.Value.NumberValue); Assert.Equal(CalculationState.Blocked, cell.Calculation!.State);
            storage.Save(copy, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
        }
        finally { File.Delete(path); File.Delete(copy); }
    }
    [Theory]
    [InlineData(1900, "xlsx")]
    [InlineData(1904, "xlsx")]
    [InlineData(1900, "xlsq")]
    [InlineData(1904, "xlsq")]
    public void IndependentExcelDateAndMathMatrixMatchesEditUndoAndRoundtrip(int system, string extension)
    {
        using var specJson = ScalarFormulaFixture.Load($"date-math-{system}-spec"); using var expectedJson = ScalarFormulaFixture.Load($"excel-date-math-{system}");
        var spec = specJson.RootElement; var expected = expectedJson.RootElement;
        using var session = new DocumentSession(ScalarFormulaFixture.Book(spec)); var calculator = new WorkbookCalculator(new NoScripts());
        void Check(string state) { ScalarFormulaFixture.Calculate(session, calculator); ScalarFormulaFixture.Compare(session.Snapshot, spec, expected.GetProperty(state)); }
        Check("Initial"); ScalarFormulaFixture.Edit(session, spec); Check("Edited");
        Assert.True(session.Undo()); Assert.False(session.CanUndo); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); using var reopened = new DocumentSession(storage.Open(path));
            Assert.False(reopened.Snapshot.FileState?.FormulaReadOnly); ScalarFormulaFixture.Calculate(reopened, new WorkbookCalculator(new NoScripts()));
            ScalarFormulaFixture.Compare(reopened.Snapshot, spec, expected.GetProperty("Edited"));
        }
        finally { File.Delete(path); }
    }
}
