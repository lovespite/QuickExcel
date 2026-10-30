using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class StructuredReferenceTests
{
    internal static WorkbookSnapshot Fixture()
    {
        var data = Sheet("Data", (0, 0, T("Qty")), (0, 1, T("Unit Price")), (0, 2, T("Amount")),
            (1, 0, N(2)), (1, 1, N(10)), (1, 2, F("=[@Qty]*[@[Unit Price]]")),
            (2, 0, N(3)), (2, 1, N(20)), (2, 2, F("=Sales[[#This Row],[Qty]]*Sales[@[Unit Price]]")),
            (3, 0, T("Total")), (3, 2, F("=SUM(Sales[Amount])")));
        var table = new WorksheetTable(1, "Sales", new(0, 0, 3, 2), [new(1, "Qty"), new(2, "Unit Price"), new(3, "Amount", CalculatedFormula: "=[@Qty]*[@[Unit Price]]")], totalsRow: true);
        return Book(data.WithData(new() { Tables = [table] }), Sheet("Summary", (0, 0, F("=SUM(Sales[Amount])")), (1, 0, F("=SUM(Sales[[Qty]:[Unit Price]])")),
            (2, 0, F("=Sales[[#Totals],[Amount]]")), (3, 0, F("=Sales[[#Headers],[Unit Price]]"))));
    }

    [Theory]
    [InlineData("=Sales[Qty]")]
    [InlineData("=[@Qty]")]
    [InlineData("=Sales[@[Qty]:[Unit Price]]")]
    [InlineData("=Sales[[#Headers], [#Data], [Qty]:[Amount]]")]
    [InlineData("='Data'!Sales[['#Has'@]]")]
    [InlineData("=Sales[@]")]
    public void StructuredSyntaxRoundTripsThroughAst(string formula)
    {
        var syntax = FormulaSyntax.Parse(formula); Assert.Equal(syntax, FormulaSyntax.Parse("=" + FormulaSyntax.Print(syntax)));
        Assert.Equal(formula, FormulaCopy.Translate(formula, new(1, 1), new(4, 5)));
    }

    [Theory]
    [InlineData("=Sales[[Qty],[Amount]]")]
    [InlineData("=Sales[[#Headers],[#Totals]]")]
    [InlineData("=Sales[[#This Row],[#All]]")]
    [InlineData("=Sales[[Unknown]")]
    [InlineData("=Sales[#Unknown]")]
    public void UnsupportedStructuredSyntaxRemainsExplicit(string formula) => Assert.Throws<FormatException>(() => FormulaSyntax.Parse(formula));

    [Fact]
    public void CurrentRowsCrossSheetAggregatesAndSpecialRowsBindIntoSparseDependencies()
    {
        var book = Fixture(); var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(book));
        var data = book.Sheets[0]; var summary = book.Sheets[1];
        Assert.Equal(CellValue.FromNumber(20), C(result, data, 1, 2).Value); Assert.Equal(CellValue.FromNumber(60), C(result, data, 2, 2).Value);
        Assert.Equal(CellValue.FromNumber(80), C(result, summary).Value); Assert.Equal(CellValue.FromNumber(35), C(result, summary, 1).Value);
        Assert.Equal(CellValue.FromNumber(80), C(result, summary, 2).Value); Assert.Equal(CellValue.FromText("Unit Price"), C(result, summary, 3).Value);
        var index = new WorkbookDependencyIndex(book); Assert.Contains(new CellAddress(data.Id, 1, 2), index.AffectedBy([new(data.Id, 1, 0)]));
        Assert.Empty(index.CycleMembers);
    }

    private static void Recalculate(DocumentSession session) => Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));

    [Fact]
    public void TableAndColumnRenameRewriteBoundNodesAndMetadataWithOneUndoEach()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new RenameTableColumnCommand(id, 1, 2, "Price # @ [net]")); Recalculate(session);
        Assert.Equal("Price # @ [net]", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        Assert.Contains("Price '# '@ '[net']", session.Snapshot.Sheets[0].Data.Tables[0].Columns[2].CalculatedFormula);
        session.Execute(new RenameTableCommand(id, 1, "Revenue")); Recalculate(session);
        Assert.Contains("Revenue", session.Snapshot.Sheets[1].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        session.Undo(); session.Undo(); Assert.Equal("Sales", session.Snapshot.Sheets[0].Data.Tables[0].Name); Assert.Equal("Unit Price", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
    }

    [Fact]
    public void InsertDeleteRowsColumnsAndSheetCopyKeepCalculatedColumnIdentity()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 2, 1)); Recalculate(session);
        Assert.Equal(4, session.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom); Assert.Equal(CellContentKind.Formula, session.Snapshot.Sheets[0].GetCell(2, 2)!.Content.Kind);
        Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 1, 1)); Recalculate(session);
        Assert.Equal("列1", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue); Assert.Equal(3, session.Snapshot.Sheets[0].Data.Tables[0].Columns[3].Id);
        Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 1, 1, true)); session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 2, 1, true));
        var copy = new CopySheetCommand(id, "Copy", 2); session.Execute(copy); Recalculate(session);
        var cloned = session.Snapshot.GetSheet(copy.Id); Assert.Equal("Sales_1", cloned.Data.Tables[0].Name); Assert.Equal(2, cloned.Data.Tables[0].Id);
        Assert.Equal(CellValue.FromNumber(80), cloned.GetCell(3, 2)!.Value);
        session.Execute(new SetCellsCommand([new(new(copy.Id, 1, 0), CellContent.FromLiteral(CellValue.FromNumber(4)))])); Recalculate(session);
        Assert.Equal(CellValue.FromNumber(100), session.Snapshot.GetSheet(copy.Id).GetCell(3, 2)!.Value); Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void DeletedTableColumnProducesReferenceErrorsAndUndoRestoresBindings()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 1, 1, true)); Recalculate(session);
        Assert.Equal(CellValue.FromError(CellErrorCode.Reference), session.Snapshot.Sheets[0].GetCell(1, 1)!.Value);
        Assert.Contains("#REF!", session.Snapshot.Sheets[0].Data.Tables[0].Columns[1].CalculatedFormula);
        session.Undo(); Recalculate(session); Assert.Equal(CellValue.FromNumber(80), session.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        session.Execute(new DeleteSheetCommand(id)); Recalculate(session); Assert.Equal(CellValue.FromError(CellErrorCode.Reference), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void ExplicitCalculatedColumnAndExpansionPreserveExistingExceptions()
    {
        var original = Fixture().Sheets[0]; var table = original.Data.Tables[0];
        var flat = new WorksheetTable(table.Id, table.Name, table.Range with { Bottom = 2 }, table.Columns);
        using var session = new DocumentSession(Book(original.WithCells(original.Cells.Remove(new(3, 0)).Remove(new(3, 2))).WithData(new() { Tables = [flat] })));
        session.Execute(new SetTableCalculatedColumnCommand(original.Id, 1, 3, "=[@Qty]+[@[Unit Price]]")); Recalculate(session);
        Assert.Equal(CellValue.FromNumber(12), session.Snapshot.Sheets[0].GetCell(1, 2)!.Value);
        session.Execute(new SetCellsCommand([new(new(original.Id, 2, 2), CellContent.FromLiteral(CellValue.Blank))]));
        session.Execute(new ResizeTableCommand(original.Id, 1, 4)); Recalculate(session);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(2, 2)); Assert.Equal(CellContentKind.Formula, session.Snapshot.Sheets[0].GetCell(4, 2)!.Content.Kind);
        session.Undo(); Assert.Equal(2, session.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom); Assert.Null(session.Snapshot.Sheets[0].GetCell(4, 2));
        session.Redo(); session.Execute(new ResizeTableCommand(original.Id, 1, 2)); Recalculate(session);
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), session.Snapshot.Sheets[0].GetCell(4, 2)!.Value);
        Assert.Contains("Sales", session.Snapshot.Sheets[0].GetCell(4, 2)!.Content.FormulaText);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ModeledStructuredFormulasReopenEditableWithFreshRecalculation(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-structured-" + Guid.NewGuid().ToString("N") + "." + extension);
        try
        {
            using var session = new DocumentSession(Fixture()); Recalculate(session); new TypedWorkbookStorage().Save(path, session.Snapshot);
            var read = new TypedWorkbookStorage().Open(path); Assert.False(read.FileState!.FormulaReadOnly); Assert.Empty(read.FileState.CompatibilityIssues);
            using var reopened = new DocumentSession(read); reopened.Execute(new RenameTableCommand(read.Sheets[0].Id, 1, "Changed")); Recalculate(reopened);
            Assert.Equal(CellValue.FromNumber(80), reopened.Snapshot.Sheets[1].GetCell(0, 0)!.Value);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DeletingAllRecordsRetainsExcelInsertionRowAndReusableCalculatedTemplate()
    {
        var sheet = Sheet("Data", (0, 0, T("Qty")), (0, 1, T("Price")), (0, 2, T("Amount")), (1, 0, N(2)), (1, 1, N(4)),
            (1, 2, F("=A2*B2")), (2, 0, N(3)), (2, 1, N(5)), (2, 2, F("=A3*B3")));
        var table = new WorksheetTable(1, "Orders", new(0, 0, 2, 2), [new(1, "Qty"), new(2, "Price"), new(3, "Amount", CalculatedFormula: "=A2*B2")], filter: new(new(0, 0, 2, 2)));
        using var session = new DocumentSession(Book(sheet.WithData(new() { Tables = [table] })));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 2, true));
        var empty = session.Snapshot.Sheets[0].Data.Tables[0]; Assert.Equal(new(0, 0, 1, 2), empty.Range); Assert.True(empty.InsertRow);
        Assert.Equal("=A2*B2", empty.Columns[2].CalculatedFormula); Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 2));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 0), CellContent.FromLiteral(CellValue.FromNumber(4))), new(new(sheet.Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(5)))]));
        Recalculate(session); Assert.Equal(CellValue.FromNumber(20), session.Snapshot.Sheets[0].GetCell(1, 2)!.Value); Assert.False(session.Snapshot.Sheets[0].Data.Tables[0].InsertRow);
        session.Undo(); Assert.True(session.Snapshot.Sheets[0].Data.Tables[0].InsertRow); Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 2));
    }

    [Fact]
    public void EmptyCalculatedTemplateFollowsDefinedNameRenameAndColumnFunctionLikeTextIsLiteral()
    {
        var sheet = Sheet("Data", (0, 0, T("CONCAT(a,b)")), (0, 1, T("Amount")));
        var table = new WorksheetTable(1, "Orders", new(0, 0, 1, 1),
            [new(1, "CONCAT(a,b)"), new(2, "Amount", CalculatedFormula: "=Rate*[@[CONCAT(a,b)]]")], insertRow: true);
        using var session = new DocumentSession(Book(sheet.WithData(new() { Tables = [table] })).WithDefinedNames([new("Rate", "=2")]));
        session.Execute(new RenameDefinedNameCommand("Rate", null, "Factor"));
        Assert.Contains("Factor", session.Snapshot.Sheets[0].Data.Tables[0].Columns[1].CalculatedFormula);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 0), CellContent.FromLiteral(CellValue.FromNumber(7)))]));
        Recalculate(session); Assert.Equal(CellValue.FromNumber(14), session.Snapshot.Sheets[0].GetCell(1, 1)!.Value);
        Assert.DoesNotContain("_xlfn.", ExcelFormulaNames.Encode("=SUM(Orders[[CONCAT(a,b)]])"));
        session.Undo(); session.Undo(); Assert.Contains("Rate", session.Snapshot.Sheets[0].Data.Tables[0].Columns[1].CalculatedFormula);
    }

    [Fact]
    public void RemovingAllDataFromTableWithTotalsIsRejectedAtomically()
    {
        using var session = new DocumentSession(Fixture()); var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(before.Sheets[0].Id, SheetAxis.Row, 1, 2, true)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void NativeStructuredFormulaFixtureRecalculatesIncludingLiteralAtSignAndEmptyTable()
    {
        using var fixture = ScalarFormulaFixture.Load("excel-structured-tables");
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-structured-fixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create))
                foreach (var part in fixture.RootElement.GetProperty("Parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new System.Text.UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
            using var session = new DocumentSession(new TypedWorkbookStorage().Open(path)); Assert.False(session.Snapshot.FileState!.FormulaReadOnly); Recalculate(session);
            foreach (var record in fixture.RootElement.GetProperty("Stages").GetProperty("initial").GetProperty("After").EnumerateArray())
            {
                var sheet = session.Snapshot.Sheets.Single(s => s.Name == record.GetProperty("Sheet").GetString());
                foreach (var expected in record.GetProperty("Cells").EnumerateArray())
                {
                    var position = CellRectangle.Parse(expected.GetProperty("Address").GetString()!); var value = sheet.GetCell(position.Top, position.Left)!.Value;
                    Assert.Equal(expected.GetProperty("Kind").GetString(), value.Kind.ToString());
                    var raw = expected.GetProperty("Value");
                    if (value.Kind == CellValueKind.Number) Assert.Equal(raw.GetDouble(), value.NumberValue, 10);
                    else Assert.Equal(raw.GetString(), CellInput.DisplayText(value));
                }
            }
        }
        finally { File.Delete(path); }
    }
}
