using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class PivotAggregationTests
{
    internal static void Compare(WorkbookSnapshot book, JsonElement reports)
    {
        foreach (var report in reports.EnumerateArray())
        {
            var sheet=book.Sheets.Single(s=>s.Name==report.GetProperty("Sheet").GetString());
            var pivot=sheet.Data.Pivots.Single(p=>p.Name==report.GetProperty("Name").GetString());
            Assert.Equal(CellRectangle.Parse(report.GetProperty("Range").GetString()!),pivot.Range);
            foreach (var item in report.GetProperty("Values").EnumerateArray())
            {
                var row=pivot.Range.Top+item.GetProperty("Row").GetInt32(); var column=pivot.Range.Left+item.GetProperty("Column").GetInt32();
                var expected=ScalarFormulaFixture.Value(item); var actual=sheet.GetCell(row,column)?.Value ?? CellValue.Blank;
                Assert.Equal(expected.Kind,actual.Kind);
                if (expected.Kind==CellValueKind.Number) Assert.Equal(expected.NumberValue,actual.NumberValue,12);
                else Assert.Equal(expected,actual);
            }
        }
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void AllAggregationsMatchExcelThroughEditRefreshHistoryAndStorage(string extension)
    {
        using var oracle=ScalarFormulaFixture.Load("excel-pivot-aggregations");
        using var session=new DocumentSession(PivotReportTests.SourceBook(extended:true)); var target=session.Snapshot.Sheets[1].Id;
        var source=new PivotSource("Data",new(0,0,11,2),"SalesData");
        foreach (var (aggregation,row,caption) in new[] {(PivotAggregation.Sum,2,"Total"),(PivotAggregation.Count,11,"Count"),
            (PivotAggregation.Average,20,"Average"),(PivotAggregation.Min,29,"Min"),(PivotAggregation.Max,38,"Max")})
            session.Execute(new CreatePivotCommand(target,new(row,0),source,"Sales"+aggregation,0,1,2,aggregation,caption));
        Compare(session.Snapshot,oracle.RootElement.GetProperty("Initial"));
        var directory=DynamicArrayStorageTests.EvidenceDirectory(); var storage=new TypedWorkbookStorage();
        var initial=Path.Combine(directory,"aggregations."+extension); storage.Save(initial,session.Snapshot);
        using var opened=new DocumentSession(storage.Open(initial));
        Assert.Empty(opened.Snapshot.FileState!.CompatibilityIssues);
        Compare(opened.Snapshot,oracle.RootElement.GetProperty("Initial"));
        target=opened.Snapshot.Sheets[1].Id;
        opened.Execute(new SetCellsCommand([new(new(opened.Snapshot.Sheets[0].Id,1,2),CellContent.FromLiteral(CellValue.FromNumber(12)))]));
        Assert.All(opened.Snapshot.GetSheet(target).Data.Pivots,p=>Assert.True(p.Stale));
        Compare(opened.Snapshot,oracle.RootElement.GetProperty("Initial"));
        var stale=Path.Combine(directory,"stale."+extension); storage.Save(stale,opened.Snapshot);
        Assert.All(storage.Open(stale).Sheets[1].Data.Pivots,p=>Assert.True(p.Stale));
        foreach (var pivot in opened.Snapshot.GetSheet(target).Data.Pivots)
        {
            var before=opened.Snapshot.GetSheet(target);
            opened.Execute(new RefreshPivotCommand(target,pivot.Name));
            Assert.False(opened.Snapshot.GetSheet(target).Data.Pivots.Single(p=>p.Name==pivot.Name).Stale);
            Assert.True(opened.Undo()); Assert.Equal(before.Cells,opened.Snapshot.GetSheet(target).Cells);
            Assert.True(opened.Snapshot.GetSheet(target).Data.Pivots.Single(p=>p.Name==pivot.Name).Stale);
            Assert.True(opened.Redo());
        }
        Compare(opened.Snapshot,oracle.RootElement.GetProperty("Edited"));
        var edited=Path.Combine(directory,"edited."+extension); storage.Save(edited,opened.Snapshot);
        Compare(storage.Open(edited),oracle.RootElement.GetProperty("Edited"));
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(directory)!,"latest-pivot-aggregations-"+extension+".txt"),initial);
    }

    [Theory]
    [InlineData(PivotAggregation.Count,2)]
    [InlineData(PivotAggregation.Min,1e308)]
    [InlineData(PivotAggregation.Max,1e308)]
    public void NonSummingAggregationsDoNotOverflowAnUnusedSum(PivotAggregation aggregation,double expected)
    {
        var cache=new PivotCache(["Group","Value"],[[CellValue.FromText("A"),CellValue.FromNumber(1e308)],
            [CellValue.FromText("A"),CellValue.FromNumber(1e308)]]);
        var result=PivotReport.Build(cache,new(0,0),0,null,1,aggregation,"Value");
        Assert.Equal(CellValue.FromNumber(expected),result.Cells[new(1,1)]);
        Assert.Equal(CellValue.FromNumber(expected),result.Cells[new(2,1)]);
    }
}
