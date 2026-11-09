using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptP3Tests
{
    private static WorkbookDocument Document(string source) => new(Book(Sheet("Data", (0, 0, N(42))))
        .WithScriptModules(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source, ScriptLanguage.JavaScript) },
            ["main"], bindings: [new("main", "actions")])),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static void Run(WorkbookDocument doc) { doc.EnableScriptsAndCalculate(); doc.RunAction("main"); }

    [Fact]
    public void WorksheetManagementAndNamedRangesAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const data=workbook.getActiveWorksheet(), report=workbook.addWorksheet('Report');
              if(workbook.getWorksheets().length!==2||report.getName()!=='Report'||report.getPosition()!==1)throw Error('add');
              if(workbook.getFirstWorksheet().getName()!=='Data'||workbook.getLastWorksheet().getName()!=='Report')throw Error('edges');
              if(data.getNext().getName()!=='Report'||report.getPrevious().getName()!=='Data'||report.getNext()!==undefined)throw Error('adjacent');
              const extra=workbook.addWorksheet();
              if(extra.getName()!=='Sheet3'||extra.getPosition()!==2)throw Error('default name');
              extra.delete();
              report.setName('Report2');
              if(workbook.getWorksheet('Report2')===undefined||report.getPosition()!==1)throw Error('rename');
              report.setPosition(0);
              if(workbook.getFirstWorksheet().getName()!=='Report2')throw Error('reorder');
              report.setPosition(1);
              report.setTabColor('#123456');if(report.getTabColor()!=='#123456')throw Error('tab');
              report.setShowGridlines(false);if(report.getShowGridlines())throw Error('gridlines');
              report.setVisibility(QuickExcel.SheetVisibility.hidden);
              if(report.getVisibility()!==QuickExcel.SheetVisibility.hidden||workbook.getLastWorksheet(true).getName()!=='Data')throw Error('visibility');
              report.setVisibility(QuickExcel.SheetVisibility.visible);
              report.getRangeByIndexes(1,1,2,2).setValues([[1,2],[3,4]]);
              if(report.getRangeByIndexes(1,1,2,2).getAddress()!=='B2:C3')throw Error('indexes');
              if(report.getRanges('B2:C3,E5').getAreaCount()!==2)throw Error('ranges');
              report.getRange('A1').setValue(11);
              const local=report.addNamedItem('LocalValue', report.getRange('A1'));
              if(local.getScope()!==QuickExcel.NamedItemScope.worksheet||local.getWorksheet().getName()!=='Report2')throw Error('local scope');
              if(local.getValue()!==11||local.getRange().getAddress()!=='A1')throw Error('local value');
              const global=workbook.addNamedItem('GlobalValue', data.getRange('A1'));
              if(global.getScope()!==QuickExcel.NamedItemScope.workbook||global.getWorksheet()!==undefined)throw Error('global scope');
              if(global.getValue()!==42||global.getType()!==QuickExcel.NamedItemType.range)throw Error('global value');
              if(!global.getFormula().startsWith('=')||workbook.getNamedItem('GlobalValue').getName()!=='GlobalValue')throw Error('formula');
              if(data.getNames().length!==0||report.getNames().length!==1)throw Error('names');
              global.setFormula('=' + data.getName() + '!$A$1');
              if(global.getRange().getAddress()!=='A1'||workbook.getNamedItem('GlobalValue')===undefined)throw Error('set formula');
              const textRef=workbook.addNamedItem('TextReference','Data!$A$1');
              if(textRef.getRange().getAddress()!=='A1'||textRef.getValue()!==42)throw Error('text reference');
              report.getFreezePanes().freezeRows(1);report.getFreezePanes().freezeColumns(2);
              if(report.getFreezePanes().getLocation().getAddress()!=='A1:B1')throw Error('freeze');
              report.getFreezePanes().unfreeze();
              if(report.getFreezePanes().getLocation()!==undefined)throw Error('unfreeze');
              local.delete();
              if(report.getNames().length!==0||workbook.getNamedItem('GlobalValue')===undefined)throw Error('delete name');
              report.delete();
              if(workbook.getWorksheets().length!==1)throw Error('delete sheet');
            }
            """);
        Run(doc); Assert.Single(doc.Snapshot.Sheets); Assert.Equal("Data", doc.Snapshot.Sheets[0].Name);
        Assert.True(doc.Undo()); Assert.Single(doc.Snapshot.Sheets); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void RangeNavigationMergeHiddenAndWholeRowColumnEditsAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const s=workbook.getActiveWorksheet(), r=s.getRange('B2:C3');
              r.setValues([[1,2],[3,4]]);
              if(r.getAddressLocal()!=='B2:C3'||r.getCell(1,1).getAddress()!=='C3')throw Error('cell');
              if(r.getRow(0).getAddress()!=='B2:C2'||r.getColumn(1).getAddress()!=='C2:C3')throw Error('row column');
              if(r.getOffsetRange(1,1).getAddress()!=='C3:D4'||r.getResizedRange(1,1).getAddress()!=='B2:D4')throw Error('offset');
              if(r.getAbsoluteResizedRange(1,4).getAddress()!=='B2:E2'||r.getEntireRow().getColumnCount()!==2048)throw Error('resize');
              if(r.getEntireColumn().getRowCount()!==200000)throw Error('entire column');
              if(r.getBoundingRect('D5').getAddress()!=='B2:D5'||r.getIntersection('C3:E5').getAddress()!=='C3')throw Error('bounds');
              if(r.getLastCell().getAddress()!=='C3'||r.getLastRow().getAddress()!=='B3:C3'||r.getLastColumn().getAddress()!=='C2:C3')throw Error('last');
              if(r.getRowsAbove().getAddress()!=='B1:C1'||r.getRowsBelow(2).getAddress()!=='B4:C5')throw Error('rows');
              if(r.getColumnsBefore().getAddress()!=='A2:A3'||r.getColumnsAfter(2).getAddress()!=='D2:E3')throw Error('columns');
              if(r.getUsedRange().getAddress()!=='B2:C3'||r.getTables().length!==0)throw Error('used tables');
              const merged=s.getRange('E5:F6');merged.merge();
              if(merged.getMergedAreas().getAreaCount()!==1)throw Error('merge');
              merged.unmerge();if(merged.getMergedAreas().getAreaCount()!==0)throw Error('unmerge');
              const row=s.getRange('A3').getEntireRow();row.setRowHidden(true);
              if(!row.getRowHidden())throw Error('hide row');row.setRowHidden(false);
              const column=s.getRange('C1').getEntireColumn();column.setColumnHidden(true);
              if(!column.getColumnHidden())throw Error('hide column');column.setColumnHidden(false);
              s.getRange('A10').setValue(7);s.getRange('A10').getEntireRow().insert(QuickExcel.InsertShiftDirection.down);
              if(s.getRange('A11').getValue()!==7)throw Error('insert row');
              s.getRange('A10').getEntireRow().delete(QuickExcel.DeleteShiftDirection.up);
              if(s.getRange('A10').getValue()!==7)throw Error('delete row');
              s.getRange('D1').setValue(9);s.getRange('D1').getEntireColumn().insert(QuickExcel.InsertShiftDirection.right);
              if(s.getRange('E1').getValue()!==9)throw Error('insert column');
              s.getRange('D1').getEntireColumn().delete(QuickExcel.DeleteShiftDirection.left);
              if(s.getRange('D1').getValue()!==9)throw Error('delete column');
            }
            """);
        Run(doc); Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Equal("7", doc.GetCellValue(0, 9, 0));
        Assert.Equal("9", doc.GetCellValue(0, 0, 3)); Assert.Equal("4", doc.GetCellValue(0, 2, 2));
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(9, 0)); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void DeletingTheActiveWorksheetSelectsAVisibleSurvivor()
    {
        using var doc = Document("""
            function main(workbook) {
              const data=workbook.getActiveWorksheet();
              workbook.addWorksheet('Next');
              data.delete();
              if(workbook.getWorksheets().length!==1||workbook.getActiveWorksheet().getName()!=='Next')throw Error('active delete');
            }
            """);
        Run(doc); Assert.Single(doc.Snapshot.Sheets); Assert.Equal("Next", doc.Snapshot.Sheets[0].Name);
        Assert.Equal(doc.Snapshot.Sheets[0].Id, doc.Snapshot.ActiveSheetId);
    }

    [Fact]
    public void HidingTheOnlyVisibleWorksheetIsRejected()
    {
        using var doc = Document("function main(workbook){workbook.getActiveWorksheet().setVisibility(QuickExcel.SheetVisibility.hidden);}");
        doc.EnableScriptsAndCalculate();
        var error = Record.Exception(() => doc.RunAction("main"));
        Assert.NotNull(error); Assert.Contains("可见工作表", error.Message);
        Assert.Equal(SheetVisibility.Visible, doc.Snapshot.Sheets[0].Visibility); Assert.False(doc.CanUndo);
    }
}
