using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptP1Tests
{
    private static WorkbookDocument Document(string source, ScriptLanguage language = ScriptLanguage.JavaScript) => new(Book(Sheet("Data", (0, 0, N(42))))
        .WithScriptModules(new(new Dictionary<string, ScriptModule> { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js",
            language == ScriptLanguage.TypeScript ? source.Replace("function main(workbook)", "function main(workbook: QuickExcel.Workbook)") : source, language) }, ["main"], scriptApiVersion: 3, bindings: [new("main", "actions")])),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static void Run(WorkbookDocument doc) { doc.EnableScriptsAndCalculate(); doc.RunAction("main"); }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void FormatsEnumsMatricesBordersAndProtectionRoundTripAsOneUndo(ScriptLanguage language)
    {
        using var doc = Document("""
            function main(workbook) {
              const r=workbook.getActiveWorksheet().getRange('B2:C3'), f=r.getFormat(), font=f.getFont();
              r.setValues([[1,2],[3,4]]);
              font.setBold(true);font.setItalic(true);font.setName('Arial');font.setSize(14);font.setColor('#123456');
              font.setStrikethrough(true);font.setUnderline(QuickExcel.RangeUnderlineStyle.double);
              if(!font.getBold()||!font.getItalic()||font.getName()!=='Arial'||font.getSize()!==14||font.getColor()!=='#123456'||!font.getStrikethrough()||font.getUnderline()!==QuickExcel.RangeUnderlineStyle.double)throw Error('font');
              f.setHorizontalAlignment(QuickExcel.HorizontalAlignment.center); f.setVerticalAlignment(QuickExcel.VerticalAlignment.top);
              f.setWrapText(true);f.setShrinkToFit(true);f.setIndentLevel(2);f.setTextOrientation(-45);
              f.setRowHeight(30); f.setColumnWidth(60);
              if(f.getHorizontalAlignment()!==QuickExcel.HorizontalAlignment.center||f.getVerticalAlignment()!==QuickExcel.VerticalAlignment.top||!f.getWrapText()||!f.getShrinkToFit()||f.getIndentLevel()!==2||f.getTextOrientation()!==-45||f.getRowHeight()!==30||f.getColumnWidth()!==60)throw Error('format');
              f.getProtection().setLocked(false);f.getProtection().setFormulaHidden(true);
              if(f.getProtection().getLocked()||!f.getProtection().getFormulaHidden())throw Error('cell protection');
              const border=f.getRangeBorder(QuickExcel.BorderIndex.edgeBottom);
              border.setStyle(QuickExcel.BorderLineStyle.continuous);border.setWeight(QuickExcel.BorderWeight.thick);border.setColor('#ff0000');
              if(border.getWeight()!==QuickExcel.BorderWeight.thick||border.getColor()!=='#FF0000')throw Error('border');
              f.getRangeBorder(QuickExcel.BorderIndex.insideVertical).setStyle(QuickExcel.BorderLineStyle.dash);
              r.setNumberFormats([['0.00','0%'],['@','0.000']]);
              if(r.getNumberFormats()[1][1]!=='0.000'||r.getTexts()[0][1]!=='200%')throw Error('number');
              if(r.getValueTypes()[0][0]!==QuickExcel.RangeValueType.integer||r.getRowIndex()!==1||r.getColumnCount()!==2)throw Error('range');
            }
            """, language);
        Run(doc); Assert.True(doc.Undo()); Assert.False(doc.CanUndo); Assert.Null(doc.Snapshot.Sheets[0].GetCell(1, 1));
        Assert.True(doc.Redo()); var style = doc.Snapshot.Styles[doc.Snapshot.Sheets[0].GetStyleId(1, 1)]; Assert.True(style.Bold); Assert.False(style.Locked);
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsq");
        try { new TypedWorkbookStorage().Save(path, doc.Snapshot); var reopened = new TypedWorkbookStorage().Open(path); Assert.Equal(3, reopened.ScriptModules.ScriptApiVersion); Assert.True(reopened.Styles[reopened.Sheets[0].GetStyleId(1, 1)].Bold); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void TablesStableColumnsSortingFiltersRenameAndConversionAreTransactional()
    {
        using var doc = Document("""
            function main(workbook) {
              const sheet=workbook.getActiveWorksheet();sheet.getRange('B2:C5').setValues([['Name','Qty'],['c',3],['a',1],['b',2]]);
              const t=sheet.addTable('B2:C5',true), id=t.getId(), col=t.getColumnByName('Qty');
              t.setName('Inventory');col.setName('Count');
              if(t.getId()!==id||col.getName()!=='Count'||workbook.getTable('Inventory').getId()!==id||col.getIndex()!==1)throw Error('stable identity');
              t.setPredefinedTableStyle('TableStyleMedium2');t.setShowBandedColumns(true);t.setHighlightFirstColumn(true);
              if(!t.getShowBandedColumns()||!t.getHighlightFirstColumn())throw Error('table style');
              t.getSort().apply([{key:1,ascending:false}]);
              if(t.getRangeBetweenHeaderAndTotal().getValues()[0][1]!==3)throw Error('sort');
              col.getFilter().applyCustomFilter('>1');
              if(!t.getAutoFilter().getIsDataFiltered()||col.getFilter().getCriteria().criterion1!=='>1')throw Error('filter');
              t.clearFilters();if(t.getAutoFilter().getIsDataFiltered())throw Error('clear filter');
              sheet.getRange('F1').setFormula('=SUM(Inventory[Count])');
              t.resize('B2:C6'); if(t.getRowCount()!==4)throw Error('resize');
              const converted=t.convertToRange(); if(converted.getAddress()!=='B2:C6'||sheet.getTables().length!==0)throw Error('convert');
              let stale=false;try{t.getName();}catch(e){stale=e.code==='InvalidReference';}if(!stale)throw Error('stale handle');
            }
            """);
        Run(doc); doc.Calculate(); Assert.Equal("6", doc.GetCellValue(0, 0, 5));
        Assert.DoesNotContain("Inventory", doc.Snapshot.Sheets[0].GetCell(0, 5)!.Content.FormulaText);
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(1, 1)); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void ProtectionChangesAreCheckedSequentiallyAndCanBeUndone()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),p=s.getProtection();
              s.getRange('B1').getFormat().getProtection().setLocked(false);
              p.protect({allowAutoFilter:true}); if(!p.getProtected()||!p.getOptions().allowAutoFilter)throw Error('protect');
              s.getRange('B1').setValue(7);p.unprotect();s.getRange('A1').setValue(9);p.protect();
            }
            """);
        Run(doc); Assert.True(doc.Snapshot.Sheets[0].Data.Protection!.Sheet); Assert.Equal("9", doc.GetCellValue(0, 0, 0)); Assert.True(doc.Undo());
        Assert.Null(doc.Snapshot.Sheets[0].Data.Protection); Assert.Equal("42", doc.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData("s.getRange('A1').getFormat().getFont().setSize(-1)")]
    [InlineData("s.getRange('A1:A20001').getFormat().getFont().setBold(true)")]
    [InlineData("s.getProtection().protect();s.getRange('A1').setValue(9)")]
    [InlineData("s.getRange('A1').getSort().apply([{key:2}])")]
    public void CaughtWriteFailuresRollbackWorkbookAndVirtualFiles(string failure)
    {
        using var doc = Document("function main(workbook,host){const s=workbook.getActiveWorksheet();s.getRange('B1').setValue(7);host.files.writeText('//mem/test','x');try{"+failure+";}catch{}}");
        doc.EnableScriptsAndCalculate(); Assert.NotNull(Record.Exception(() => doc.RunAction("main")));
        Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1)); Assert.Null(doc.Snapshot.Sheets[0].Data.Protection); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void WorksheetFilterAndRangeSortUseRelativeColumnIndexes()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('D2:E5');
              r.setValues([['Key','Value'],['b',2],['a',1],['c',3]]);r.getSort().apply([{key:0}],false,true);
              if(r.getValues()[1][0]!=='a')throw Error('sort key');
              const f=s.getAutoFilter();f.apply(r,1,{filterOn:QuickExcel.FilterOn.values,values:['2','3']});
              if(!f.getEnabled()||!f.getIsDataFiltered()||f.getRange().getAddress()!=='D2:E5')throw Error('filter');
              f.clearColumnCriteria(1);if(f.getIsDataFiltered())throw Error('clear');f.remove();if(f.getEnabled())throw Error('remove');
            }
            """);
        Run(doc); Assert.Equal("a", doc.GetCellValue(0, 2, 3)); Assert.True(doc.Undo());
    }

    [Fact]
    public void FormulaSortUsesCurrentDraftKeysWithoutReplayingUdfsDuringCommit()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet();
              s.getRange('B1:C3').setValues([['a',0],['b',0],['c',0]]);
              s.getRange('C1:C3').setFormulas([['=3'],['=1'],['=2']]);
              s.getRange('B1:C3').getSort().apply([{key:1}]);
              if(s.getRange('B1:B3').getValues().map(r=>r[0]).join(',')!=='b,c,a')throw Error('formula sort');
            }
            """);
        Run(doc); Assert.Equal("b", doc.GetCellValue(0, 0, 1)); Assert.True(doc.Undo());
    }

    [Fact]
    public void DeletedTableHandlesCannotAliasNewTablesAndClearModesAreExplicit()
    {
        using var doc = Document("""
            function main(workbook,host){const s=workbook.getActiveWorksheet();s.getRange('B1:C2').setValues([['A','B'],[1,2]]);
              const old=s.addTable('B1:C2',true);old.convertToRange();const current=s.addTable('B1:C2',true);
              let stale=false;try{old.getName();}catch(e){stale=e.code==='InvalidReference';}if(!stale)throw Error('id reused');
            }
            """);
        Run(doc); Assert.Single(doc.Snapshot.Sheets[0].Data.Tables);
        doc.RunConsole("const r=workbook.getActiveWorksheet().getRange('E1');r.setFormula('=1/0');if(host.values.getErrorCode(r)!=='#DIV/0!')throw Error('error code');r.getFormat().getFont().setBold(true);r.clear(QuickExcel.ClearApplyTo.contents);if(!r.getFormat().getFont().getBold()||r.getValue()!=='')throw Error('contents');r.setValue(4);r.clear(QuickExcel.ClearApplyTo.formats);if(r.getValue()!==4||r.getFormat().getFont().getBold())throw Error('formats');r.clear();if(r.getValue()!=='')throw Error('all');");
    }

    [Theory]
    [InlineData("for(let i=0;i<3;i++)r.getFormat().getFont().setBold(true)", "20000")]
    [InlineData("for(let i=0;i<11;i++)r.getNumberFormats()", "100000")]
    public void CumulativeFormatBudgetsAreEnforced(string operations, string limit)
    {
        using var doc = Document("function main(workbook){const r=workbook.getActiveWorksheet().getRange('B1:B10000');" + operations + ";}");
        doc.EnableScriptsAndCalculate(); var error = Record.Exception(() => doc.RunAction("main")); Assert.NotNull(error); Assert.Contains(limit, error.Message); Assert.False(doc.CanUndo);
        Assert.Single(doc.Snapshot.Sheets[0].Cells);
    }

    [Fact]
    public void MaximumRawActionSourceBudgetIncludesRoomForHostRegistration()
    {
        var source = "function main(workbook){workbook.getActiveWorksheet().getRange('B1').setValue(7);}\n//";
        using var doc = Document(source.PadRight(1024 * 1024, 'x')); Run(doc); Assert.Equal("7", doc.GetCellValue(0, 0, 1));
    }

    [Fact]
    public void FormulaFiltersUseDraftValuesAndRespectAllowedProtectedTableFiltering()
    {
        using var doc = Document("""
            function main(workbook,host){const s=workbook.getActiveWorksheet();
              s.getRange('B1:C3').setValues([['Name','Value'],['a',0],['b',0]]);
              s.getRange('C2:C3').setFormulas([['=1+1'],['=3+1']]);
              const table=s.addTable('B1:C3',true);s.getProtection().protect({allowAutoFilter:true});
              table.getColumnByName('Value').getFilter().applyCustomFilter('>2');
              if(!table.getAutoFilter().getIsDataFiltered())throw Error('filter');
              host.files.writeText('//xlsq/protected.txt','success');
            }
            """);
        Run(doc); Assert.True(doc.Snapshot.Sheets[0].Data.Protection!.Sheet);
        Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/protected.txt"));
        doc.Calculate(); Assert.Contains(1, WorksheetFilterEngine.Evaluate(doc.Snapshot, doc.Snapshot.ActiveSheetId));
        Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].Data.Protection); Assert.Empty(doc.Snapshot.PackageFiles.Entries);
    }

    [Fact]
    public void TableAppendPreservesRowsAndConversionPreservesEffectiveFormatting()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet();s.getRange('B1:C2').setValues([['A','B'],[1,2]]);
              const t=s.addTable('B1:C2',true);t.setPredefinedTableStyle('TableStyleMedium2');
              if(!t.getHeaderRowRange().getFormat().getFill().getColor())throw Error('effective table style');
              t.addRows(-1,[[3,4],[5,6]]);t.addRow(-1,[7,8]);
              if(t.getRowCount()!==4||t.getRange().getValues()[4][1]!==8)throw Error('append');t.convertToRange();
            }
            """);
        Run(doc); var sheet = doc.Snapshot.Sheets[0];
        Assert.Equal("8", doc.GetCellValue(0, 4, 2)); Assert.Empty(sheet.Data.Tables);
        Assert.NotNull(doc.Snapshot.Styles[sheet.GetStyleId(0, 1)].FillColor); Assert.True(doc.Undo()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
    }

    [Fact]
    public void FormattingAfterPrivateCalculationIsVisibleToFilterValidation()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('B2');
              s.getRange('B1').setValue('Value');r.setFormula('=1+1');r.setNumberFormat('0* ');
              if(r.getValue()!==2)throw Error('value');r.setNumberFormat('0');
              s.getAutoFilter().apply('B1:B2',0,{filterOn:QuickExcel.FilterOn.values,values:['2']});
            }
            """);
        Run(doc); doc.Calculate(); Assert.Empty(WorksheetFilterEngine.Evaluate(doc.Snapshot, doc.Snapshot.ActiveSheetId));
    }
}
