using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using QuickExcel.Presentation;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class OrderStatisticTests
{
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void IndependentExcelMatrixMatchesIncrementalHistoryAndBothFormats(string extension)
    {
        using var specJson=ScalarFormulaFixture.Load("order-statistic-spec");
        using var oracle=ScalarFormulaFixture.Load("excel-order-statistics"); var spec=specJson.RootElement;
        using var session=new DocumentSession(ScalarFormulaFixture.Book(spec)); var calculator=new WorkbookCalculator(new NoScripts());
        void Check(string state)
        {
            ScalarFormulaFixture.Calculate(session,calculator);
            ScalarFormulaFixture.Compare(session.Snapshot,spec,oracle.RootElement.GetProperty(state));
        }
        Check("Initial"); ScalarFormulaFixture.Edit(session,spec); Check("Edited");
        Assert.True(session.Undo()); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var directory=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"../../../../e5/order-statistics",Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory); var path=Path.Combine(directory,"statistics."+extension); var storage=new TypedWorkbookStorage();
        storage.Save(path,session.Snapshot);
        using var opened=new DocumentSession(storage.Open(path)); Assert.False(opened.Snapshot.FileState!.FormulaReadOnly);
        Assert.Empty(opened.Snapshot.FileState.CompatibilityIssues);
        ScalarFormulaFixture.Calculate(opened,new WorkbookCalculator(new NoScripts()));
        ScalarFormulaFixture.Compare(opened.Snapshot,spec,oracle.RootElement.GetProperty("Edited"));
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(directory)!,"latest-"+extension+".txt"),path);
    }

    [Theory]
    [InlineData("=LARGE(A1:A3,B1:B2)")]
    [InlineData("=SMALL(A1:A3,SEQUENCE(2))")]
    [InlineData("=RANK.EQ(A1:A3,B1:B3)")]
    [InlineData("=RANK.EQ(1,A1:A3,B1#)")]
    [InlineData("=LET(k,SEQUENCE(2),LARGE(A1:A3,k))")]
    public void UnimplementedArrayOptionsPreserveOriginalFormulaAndOldCache(string formula)
    {
        var directory=DynamicArrayStorageTests.EvidenceDirectory(); var path=Path.Combine(directory,"guard.xlsx");
        var cell=new System.Xml.Linq.XElement("row",new System.Xml.Linq.XAttribute("r",1),
            new System.Xml.Linq.XElement("c",new System.Xml.Linq.XAttribute("r","D1"),
                new System.Xml.Linq.XElement("f",ExcelFormulaNames.Encode(formula)[1..]),new System.Xml.Linq.XElement("v",123)));
        var bytes=Fixtures.WorkbookPackageFixture.Xlsx(cell.ToString()); File.WriteAllBytes(path,bytes);
        var storage=new TypedWorkbookStorage(); using var session=new DocumentSession(storage.Open(path));
        Assert.True(session.Snapshot.FileState!.FormulaReadOnly);
        ScalarFormulaFixture.Calculate(session,new WorkbookCalculator(new NoScripts()));
        var output=session.Snapshot.Sheets[0].GetCell(0,3)!;
        Assert.Equal(CalculationState.Blocked,output.Calculation!.State); Assert.Equal(CellValue.FromNumber(123),output.Value);
        var copy=Path.Combine(directory,"copy.xlsx"); storage.Save(copy,session.Snapshot); Assert.Equal(bytes,File.ReadAllBytes(copy));
    }

    [Fact]
    public void NumericBudgetAndFormulaCompletionRemainExplicit()
    {
        var values=Enumerable.Range(0,100001).Select(row=>KeyValuePair.Create(new CellPosition(row,0),N(row)));
        var sheet=new WorksheetSnapshot(SheetId.New(),"S",values.Append(KeyValuePair.Create(new CellPosition(0,1),F("=LARGE(A1:A100001,1)"))));
        var result=new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet)));
        Assert.Equal(CellValue.FromError(CellErrorCode.Number),result.Cells[new(sheet.Id,0,1)].Value);
        Assert.Contains("RANK.EQ",FormulaEditing.Complete("=RANK",5)!.Names);
        Assert.Equal("LARGE(array, k)",FormulaEditing.Signature("LARGE"));
        Assert.Equal("RANK.EQ(number, ref, [order])",FormulaEditing.Signature("RANK.EQ"));
        Assert.Equal("=_xlfn.RANK.EQ(A1,A1:A3)",ExcelFormulaNames.Encode("=RANK.EQ(A1,A1:A3)"));
        var chain="=LET("+string.Join(',',Enumerable.Range(0,60).Select(i=>$"local_{i},"+(i==0 ? "1" : $"local_{i-1}+local_{i-1}")))+",LARGE(A1:A3,local_59))";
        Assert.NotNull(FormulaBinding.Bind(FormulaSyntax.Parse(chain),Book(sheet),sheet.Id));
    }
}
