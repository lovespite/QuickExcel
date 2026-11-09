using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptP4Tests
{
    private static WorkbookDocument Document(string source) => new(Book(Sheet("Data", (0, 0, N(42))))
        .WithScriptModules(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source, ScriptLanguage.JavaScript) },
            ["main"], bindings: [new("main", "actions")])),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static void Run(WorkbookDocument doc) { doc.EnableScriptsAndCalculate(); doc.RunAction("main"); }

    [Fact]
    public void FindReplaceAndRemoveDuplicatesAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('A1:C3').setValues([['cat','dog',42],['CAT','',0],['caterpillar','',0]]);
              const r=s.getRange('A1:C3');
              if(r.find('cat',{completeMatch:true}).getAddress()!=='A1')throw Error('find partial first');
              if(s.getRange('A1:A3').find('cat',{completeMatch:true,matchCase:true}).getAddress()!=='A1')throw Error('find case');
              if(s.getRange('A1:A3').find('cat',{completeMatch:true,searchDirection:QuickExcel.SearchDirection.backwards}).getAddress()!=='A2')throw Error('find backwards');
              if(s.getRange('A1').find('dog').getAddress()!=='B1')throw Error('find from cell');
              if(s.getRange('A1:A3').replaceAll('cat','dog',{completeMatch:true})!==2)throw Error('replace count');
              if(s.getRange('A1').getValue()!=='dog'||s.getRange('A2').getValue()!=='dog'||s.getRange('A3').getValue()!=='caterpillar')throw Error('replace values');
              s.getRange('A4').setValue('cat cat');
              if(s.getRange('A4').replaceAll('cat','dog')!==2||s.getRange('A4').getValue()!=='dog dog')throw Error('replace occurrences');
              if(s.getRange('C1').replaceAll('42','7',{completeMatch:true})!==1||s.getRange('C1').getValue()!==7)throw Error('replace number');
              s.getRange('E1:F4').setValues([['Key','Value'],['a',1],['b',2],['a',3]]);
              const result=s.getRange('E1:F4').removeDuplicates([0],true);
              if(result.getRemoved()!==1||result.getUniqueRemaining()!==2)throw Error('duplicates result');
              if(s.getRange('E2:F3').getValues().map(row=>row.join(':')).join(',')!=='a:1,b:2')throw Error('duplicates values');
            }
            """);
        Run(doc);
        Assert.Equal("dog", doc.GetCellValue(0, 0, 0));
        Assert.Equal("7", doc.GetCellValue(0, 0, 2));
        Assert.Null(doc.Snapshot.Sheets[0].GetCell(3, 4));
        Assert.True(doc.Undo()); Assert.Equal("42", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void AutoFillAndFlashFillAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('B1:B2').setValues([[1],[2]]);
              s.getRange('B1:B2').autoFill('B1:B6',QuickExcel.AutoFillType.fillSeries);
              if(s.getRange('B1:B6').getValues().map(r=>r[0]).join(',')!=='1,2,3,4,5,6')throw Error('fill series');
              s.getRange('C1:C2').setValues([[4],[5]]);
              s.getRange('C1:C2').autoFill(s.getRange('C1:C4'),QuickExcel.AutoFillType.fillCopy);
              if(s.getRange('C1:C4').getValues().map(r=>r[0]).join(',')!=='4,5,4,5')throw Error('fill copy');
              s.getRange('D1:D2').setValues([[10],[20]]);
              s.getRange('E1:E6').setValues([[1],[2],[3],[4],[5],[6]]);
              s.getRange('D1:D2').autoFill();
              if(s.getRange('D5').getValue()!==50)throw Error('auto fill extent');
              s.getRange('B8:B9').setValues([['John Smith'],['Jane Doe']]);
              s.getRange('C8').setValue('John');
              s.getRange('C8:C9').flashFill();
              if(s.getRange('C8').getValue()!=='John'||s.getRange('C9').getValue()!=='Jane')throw Error('flash fill');
            }
            """);
        Run(doc);
        Assert.Equal("6", doc.GetCellValue(0, 5, 1));
        Assert.Equal("5", doc.GetCellValue(0, 3, 2));
        Assert.Equal("50", doc.GetCellValue(0, 4, 3));
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
    }

    [Fact]
    public void HyperlinksAndWorksheetActivationAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const data=workbook.getActiveWorksheet(), next=workbook.addWorksheet('Next');
              next.activate();
              if(workbook.getActiveWorksheet().getName()!=='Next')throw Error('activate');
              const cell=next.getRange('A1');
              cell.setHyperlink({address:'https://example.com',screenTip:'tip',textToDisplay:'Go'});
              const link=cell.getHyperlink();
              if(link.address!=='https://example.com'||link.screenTip!=='tip'||link.textToDisplay!=='Go')throw Error('hyperlink');
              cell.setHyperlink({documentReference:'Next!A1'});
              if(cell.getHyperlink().documentReference!=='Next!A1')throw Error('internal link');
              cell.setHyperlink({});
              if(cell.getHyperlink()!==undefined)throw Error('clear link');
              data.activate();
              if(workbook.getActiveWorksheet().getName()!=='Data')throw Error('reactivate');
            }
            """);
        Run(doc);
        Assert.Equal("Data", doc.Snapshot.Sheets[0].Name);
        Assert.Equal(doc.Snapshot.Sheets[0].Id, doc.Snapshot.ActiveSheetId);
        Assert.True(doc.Undo()); Assert.Single(doc.Snapshot.Sheets);
    }

    [Fact]
    public void ReplaceAllRejectsFormulaResults()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet();
              s.getRange('A1').setFormula('="cat"');
              s.getRange('A1').replaceAll('cat','dog',{completeMatch:true});
            }
            """);
        doc.EnableScriptsAndCalculate();
        var error = Record.Exception(() => doc.RunAction("main"));
        Assert.NotNull(error); Assert.Contains("公式结果", error.Message);
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.False(doc.CanUndo);
    }
}
