using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class PivotReportTests
{
    internal static WorkbookSnapshot SourceBook(bool extended = false)
    {
        object?[][] values=[["Region","Product","Amount"],["North","A",10d],["South","A",20d],["North","B",5d],
            ["South","B",0d],["North","A",7d],["South","A","text"],["East","B",-3d],["East","A",null]];
        if (extended) values=[..values,["South","C","123"],["North","C",true],["East","C",null]];
        var cells=values.SelectMany((row,r)=>row.Select((v,c)=>KeyValuePair.Create(new CellPosition(r,c),new CellSnapshot(CellContent.FromLiteral(
            v is double number ? CellValue.FromNumber(number) : v is string text ? CellValue.FromText(text) : v is bool boolean ? CellValue.FromBoolean(boolean) : CellValue.Blank)))));
        var data=new WorksheetSnapshot(SheetId.New(),"Data",cells,data:new() { Tables=[new(1,"SalesData",new(0,0,values.Length-1,2),[new(1,"Region"),new(2,"Product"),new(3,"Amount")])] });
        return new(Guid.NewGuid(),0,ExcelDateSystem.Windows1900,[data,new(SheetId.New(),"Summary")]);
    }
    private static PivotSource Source => new("Data",new(0,0,8,2),"SalesData");

    [Fact]
    public void SumAndCountResultsMatchIndependentExcelIncludingBlankAndTextMeasures()
    {
        using var evidence=ScalarFormulaFixture.Load("excel-pivots");
        using var session=new DocumentSession(SourceBook()); var target=session.Snapshot.Sheets[1].Id;
        session.Execute(new CreatePivotCommand(target,new(2,0),Source,"SalesSum",0,1,2,PivotAggregation.Sum,"Total"));
        session.Execute(new CreatePivotCommand(target,new(11,0),Source,"SalesCount",0,1,2,PivotAggregation.Count,"Count"));
        void Check(string state)
        {
            foreach (var report in evidence.RootElement.GetProperty(state).EnumerateArray())
            {
                var pivot=session.Snapshot.GetSheet(target).Data.Pivots.Single(p=>p.Name==report.GetProperty("Name").GetString());
                Assert.Equal(CellRectangle.Parse(report.GetProperty("Range").GetString()!),pivot.Range);
                foreach (var item in report.GetProperty("Values").EnumerateArray())
                {
                    var position=new CellPosition(pivot.Range.Top+item.GetProperty("Row").GetInt32(),pivot.Range.Left+item.GetProperty("Column").GetInt32());
                    Assert.Equal(ScalarFormulaFixture.Value(item),session.Snapshot.GetSheet(target).GetCell(position.Row,position.Column)?.Value ?? CellValue.Blank);
                }
            }
        }
        Check("Initial");
        session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id,1,2),CellContent.FromLiteral(CellValue.FromNumber(12)))]));
        Assert.All(session.Snapshot.GetSheet(target).Data.Pivots,p=>Assert.True(p.Stale)); Check("Initial");
        session.Execute(new RefreshPivotCommand(target,"SalesSum")); session.Execute(new RefreshPivotCommand(target,"SalesCount"));
        Assert.All(session.Snapshot.GetSheet(target).Data.Pivots,p=>Assert.False(p.Stale)); Check("Edited");
        Assert.True(session.Undo()); Assert.True(session.Snapshot.GetSheet(target).Data.Pivots.Single(p=>p.Name=="SalesCount").Stale);
    }

    [Fact]
    public void OutputOwnershipAndRefreshCollisionAreAtomicAndConversionLeavesValues()
    {
        using var session=new DocumentSession(SourceBook()); var target=session.Snapshot.Sheets[1].Id; var data=session.Snapshot.Sheets[0].Id;
        session.Execute(new CreatePivotCommand(target,new(0,0),Source,"Pivot",0,null,2,PivotAggregation.Sum,"Total"));
        var before=session.Snapshot;
        Assert.Throws<InvalidOperationException>(()=>session.Execute(new SetCellsCommand([new(new(target,1,1),CellContent.FromLiteral(CellValue.FromNumber(3)))])));
        Assert.Same(before,session.Snapshot);
        Assert.Throws<NotSupportedException>(()=>session.Execute(new ShiftAxisCommand(data,SheetAxis.Row,0,1)));
        session.Execute(new SetCellStyleCommand([new(target,1,1)],new(Bold:true)));
        session.Execute(new SetSheetViewCommand(target,session.Snapshot.GetSheet(target).View with { Zoom=125 }));
        session.Execute(new SetAxisMetadataCommand(target,SheetAxis.Column,0,1,size:18));
        Assert.False(session.Snapshot.GetSheet(target).Data.Pivots.Single().Stale);
        session.Execute(new SetCellsCommand([new(new(target,5,0),CellContent.FromLiteral(CellValue.FromText("occupied"))),
            new(new(data,8,0),CellContent.FromLiteral(CellValue.FromText("West")))]));
        before=session.Snapshot;
        Assert.Throws<InvalidOperationException>(()=>session.Execute(new RefreshPivotCommand(target,"Pivot")));
        Assert.Same(before,session.Snapshot);
        session.Execute(new RemovePivotCommand(target,"Pivot"));
        Assert.Empty(session.Snapshot.GetSheet(target).Data.Pivots);
        Assert.Equal(CellValue.FromNumber(-3),session.Snapshot.GetSheet(target).GetCell(1,1)!.Value);
        Assert.True(session.Undo()); Assert.Single(session.Snapshot.GetSheet(target).Data.Pivots);
    }

    [Fact]
    public void CacheOwnsValuesAndRefreshRequiresFreshFormulaSourcesAndHonorsCancellation()
    {
        var records=new[] {new[] {CellValue.FromNumber(1)}}; var cache=new PivotCache(["Amount"],records);
        records[0][0]=CellValue.FromNumber(9); Assert.Equal(CellValue.FromNumber(1),cache.Records[0][0]);
        Assert.Throws<ArgumentException>(()=>new PivotCache(["A","a"],[]));
        Assert.Throws<ArgumentException>(()=>new PivotCache(["A"],[new[] {CellValue.Blank,CellValue.Blank}]));
        using var session=new DocumentSession(SourceBook()); var data=session.Snapshot.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(data,1,2),CellContent.FromFormula("=20+1"))]));
        Assert.Throws<InvalidOperationException>(()=>PivotReport.Capture(session.Snapshot,Source));
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal(CellValue.FromNumber(21),PivotReport.Capture(session.Snapshot,Source).Records[0][2]);
        using var cancellation=new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>PivotReport.Capture(session.Snapshot,Source,cancellation.Token));
    }
}
