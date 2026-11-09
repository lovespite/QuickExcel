using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptP5Tests
{
    private static WorkbookDocument Document(string source) => new(Book(Sheet("Data", (0, 0, N(42))))
        .WithScriptModules(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source, ScriptLanguage.JavaScript) },
            ["main"], bindings: [new("main", "actions")])),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static void Run(WorkbookDocument doc) { doc.EnableScriptsAndCalculate(); doc.RunAction("main"); }

    [Fact]
    public void CopyFromAndHyperlinkClearModesAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('B1:C2').setValues([[1,2],[3,4]]);
              s.getRange('B1').getFormat().getFont().setBold(true);
              s.getRange('C2').getFormat().getFont().setBold(true);
              s.getRange('B3').setFormula('=1+1');
              s.getRange('E1').copyFrom(s.getRange('B1:C2'));
              if(s.getRange('E1:F2').getValues().map(r=>r.join(',')).join(';')!=='1,2;3,4')throw Error('copy');
              s.getRange('E3').copyFrom('B3',QuickExcel.RangeCopyType.values);
              if(s.getRange('E3').getValue()!==2||s.getRange('E3').getFormula()!=='')throw Error('copy values');
              s.getRange('E5').copyFrom(s.getRange('B1:C2'),QuickExcel.RangeCopyType.all,false,true);
              if(s.getRange('E5:F6').getValues().map(r=>r.join(',')).join(';')!=='1,3;2,4')throw Error('copy transpose');
              s.getRange('E8:F9').copyFrom('B1:C2');
              if(s.getRange('E8:F9').getValues().map(r=>r.join(',')).join(';')!=='1,2;3,4')throw Error('copy exact');
              s.getRange('G1').copyFrom(s.getRange('B1:C2'),QuickExcel.RangeCopyType.formats);
              if(!s.getRange('G1').getFormat().getFont().getBold()||!s.getRange('H2').getFormat().getFont().getBold())throw Error('copy formats');
              s.getRange('A1').setHyperlink({address:'https://example.com'});
              s.getRange('A1').clear(QuickExcel.ClearApplyTo.hyperlinks);
              if(s.getRange('A1').getHyperlink()!==undefined||s.getRange('A1').getValue()!==42)throw Error('clear hyperlinks');
              s.getRange('A1').setHyperlink({address:'https://example.com'});
              s.getRange('A1').getFormat().getFont().setBold(true);
              s.getRange('A1').clear(QuickExcel.ClearApplyTo.removeHyperlinks);
              if(s.getRange('A1').getHyperlink()!==undefined||s.getRange('A1').getValue()!==42||s.getRange('A1').getFormat().getFont().getBold())throw Error('remove hyperlinks');
              s.getRange('A2').setValue(7);
              s.getRange('A2').getFormat().getFont().setBold(true);
              s.getRange('A2').clearOrResetContents();
              if(s.getRange('A2').getValue()!==''||!s.getRange('A2').getFormat().getFont().getBold())throw Error('reset contents');
            }
            """);
        Run(doc);
        var sheet = doc.Snapshot.Sheets[0];
        Assert.Equal("2", doc.GetCellValue(0, 2, 4));
        Assert.True(doc.Snapshot.Styles[sheet.GetStyleId(0, 6)].Bold);
        Assert.True(doc.Snapshot.Styles[sheet.GetStyleId(1, 7)].Bold);
        Assert.True(doc.Snapshot.Styles[sheet.GetStyleId(1, 0)].Bold);
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
    }

    [Fact]
    public void SurroundingRegionRangeEdgeAndGroupingAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('A1:B3').setValues([[1,2],[3,4],[5,6]]);
              s.getRange('D1:D2').setValues([[9],[8]]);
              s.getRange('B6').setValue(7);
              if(s.getRange('A2').getSurroundingRegion().getAddress()!=='A1:B3')throw Error('region');
              if(s.getRange('C2').getSurroundingRegion().getAddress()!=='C2')throw Error('blank region');
              if(s.getRange('A1').getRangeEdge(QuickExcel.KeyboardDirection.down).getAddress()!=='A3')throw Error('edge down');
              if(s.getRange('A1').getRangeEdge(QuickExcel.KeyboardDirection.right).getAddress()!=='B1')throw Error('edge right');
              if(s.getRange('B3').getRangeEdge(QuickExcel.KeyboardDirection.down).getAddress()!=='B6')throw Error('edge gap');
              if(s.getRange('A5').getRangeEdge(QuickExcel.KeyboardDirection.right).getAddress()!=='BZT5')throw Error('edge boundary');
              const rows=s.getRange('A10:B11').getEntireRow();
              rows.group(QuickExcel.GroupOption.byRows);
              rows.hideGroupDetails(QuickExcel.GroupOption.byRows);
              if(!rows.getRowHidden())throw Error('hide details');
              rows.showGroupDetails(QuickExcel.GroupOption.byRows);
              if(rows.getRowHidden())throw Error('show details');
              rows.ungroup(QuickExcel.GroupOption.byRows);
            }
            """);
        Run(doc);
        var sheet = doc.Snapshot.Sheets[0]; Assert.Equal("7", doc.GetCellValue(0, 5, 1));
        Assert.False(sheet.Rows.GetValueOrDefault(9)?.OutlineLevel > 0);
        Assert.True(doc.Undo()); Assert.Equal("42", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void AutoFillExtendedTypesCoverValuesFormatsTrendsAndDates()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('H1:H2').setValues([[1],[2]]);
              s.getRange('H1:H2').autoFill('H1:H5',QuickExcel.AutoFillType.fillValues);
              if(s.getRange('H1:H5').getValues().map(r=>r[0]).join(',')!=='1,2,1,2,1')throw Error('fill values');
              s.getRange('I1:I2').getFormat().getFont().setBold(true);
              s.getRange('I1:I2').autoFill('I1:I4',QuickExcel.AutoFillType.fillFormats);
              if(!s.getRange('I3').getFormat().getFont().getBold()||!s.getRange('I4').getFormat().getFont().getBold())throw Error('fill formats');
              s.getRange('J1:J2').setValues([[1],[3]]);
              s.getRange('J1:J2').autoFill('J1:J5',QuickExcel.AutoFillType.linearTrend);
              const trend=s.getRange('J1:J5').getValues().map(r=>r[0]).join(',');
              if(trend!=='1,3,5,7,9')throw Error('linear trend:'+trend);
              s.getRange('K1:K2').setValues([[1],[2]]);
              s.getRange('K1:K2').autoFill('K1:K5',QuickExcel.AutoFillType.growthTrend);
              if(s.getRange('K1:K5').getValues().map(r=>r[0]).join(',')!=='1,2,4,8,16')throw Error('growth trend');
              s.getRange('L1:L2').setValues([[1],[2]]);
              s.getRange('L1:L2').autoFill('L1:L4',QuickExcel.AutoFillType.fillDays);
              if(s.getRange('L1:L4').getValues().map(r=>r[0]).join(',')!=='1,2,3,4')throw Error('fill days');
              s.getRange('M1').setValue(45292);
              s.getRange('M1').autoFill('M1:M6',QuickExcel.AutoFillType.fillWeekdays);
              if(s.getRange('M2').getValue()!==45293||s.getRange('M6').getValue()!==45299)throw Error('weekdays');
              s.getRange('N1:N2').setValues([[45292],[45323]]);
              s.getRange('N1:N2').autoFill('N1:N4',QuickExcel.AutoFillType.fillMonths);
              if(s.getRange('N3').getValue()!==45352)throw Error('months');
            }
            """);
        Run(doc);
        Assert.Equal("16", doc.GetCellValue(0, 4, 10));
        Assert.Equal("45299", doc.GetCellValue(0, 5, 12));
        Assert.Equal("45352", doc.GetCellValue(0, 2, 13));
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 7));
    }
}
