using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class InsertedFormattingTests
{
    [Theory]
    [InlineData(SheetAxis.Row, 0)]
    [InlineData(SheetAxis.Row, 1)]
    [InlineData(SheetAxis.Column, 0)]
    [InlineData(SheetAxis.Column, 1)]
    public void InsertCopiesOnlyNeighborFormattingAndDimensionsAtomically(SheetAxis axis, int index)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "S", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1)))],
            rows: [new(0, new(Height: 32, Hidden: true))], columns: [new(0, new(Width: 26, Hidden: true))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source], styles: [new(), new() { Bold = true, NumberFormatId = 49 }]);
        using var session = new DocumentSession(book);
        session.Execute(new ShiftAxisCommand(source.Id, axis, index, 2));
        var result = session.Snapshot.Sheets[0];
        for (var n = index; n < index + 2; n++)
        {
            var p = axis == SheetAxis.Row ? new CellPosition(n, 0) : new CellPosition(0, n);
            Assert.Equal(CellValue.Blank, result.GetCell(p.Row, p.Column)?.Value ?? CellValue.Blank);
            Assert.Equal(index == 0 ? 0 : 1, result.GetStyleId(p.Row, p.Column));
            if (axis == SheetAxis.Row)
            {
                Assert.Equal(index == 0 ? null : (double?)32, result.Rows.GetValueOrDefault(n)?.Height);
                Assert.False(result.Rows.GetValueOrDefault(n)?.Hidden ?? false);
            }
            else
            {
                Assert.Equal(index == 0 ? null : (double?)26, result.Columns.GetValueOrDefault(n)?.Width);
                Assert.False(result.Columns.GetValueOrDefault(n)?.Hidden ?? false);
            }
        }
        Assert.Single(result.Cells.Values, c => c.Value.Kind == CellValueKind.Text);
        Assert.True(session.Undo()); Assert.Equal(source.Cells, session.Snapshot.Sheets[0].Cells); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.Equal(result.Cells, session.Snapshot.Sheets[0].Cells);
    }

    [Theory]
    [InlineData("=数据!#REF!*2", "=(#REF!*2)")]
    [InlineData("=SUM('O''Brien'!#REF!)", "=SUM(#REF!)")]
    public void ExcelDeletedQualifiedReferenceParsesAndSavesWithoutLosingFormula(string formula, string canonical)
    {
        Assert.Equal(canonical, "=" + FormulaSyntax.Print(FormulaSyntax.Parse(formula)));
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "qualified-ref-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "S", [new(new(0, 0), new(CellContent.FromFormula(formula), new(CellValue.FromError(CellErrorCode.Reference), CalculationState.Stale, 0, 0, 0)))])]);
        var path = Path.Combine(folder, "error.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, book);
        var reopened = storage.Open(path);
        Assert.False(reopened.FileState!.FormulaReadOnly); Assert.Equal(formula, reopened.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        using var session = new DocumentSession(reopened);
        var result = new QuickExcel.Calculation.WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation());
        Assert.Equal(CellValue.FromError(CellErrorCode.Reference), result.Cells[new(reopened.Sheets[0].Id, 0, 0)].Value);
    }
}
