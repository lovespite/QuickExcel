using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptP2Tests
{
    private static WorkbookDocument Document(string source, ScriptLanguage language = ScriptLanguage.JavaScript) => new(Book(Sheet("Data", (0, 0, N(42))))
        .WithScriptModules(new(new Dictionary<string, ScriptModule> { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js",
            language == ScriptLanguage.TypeScript ? TypeScriptSource(source) : source, language) }, ["main"], bindings: [new("main", "actions")])),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static string TypeScriptSource(string source) => System.Text.RegularExpressions.Regex.Replace(
        source.Replace("function main(workbook)", "function main(workbook: QuickExcel.Workbook)"),
        @"(\.get(?:CellValue|Custom|Preset)\(\)|\.getConditionalFormat\(id\)|\.getChart\('Report'\))", "$1!");
    private static void Run(WorkbookDocument doc)
    {
        try { doc.EnableScriptsAndCalculate(); doc.RunAction("main"); }
        catch (ScriptCompilationException e) { throw new Xunit.Sdk.XunitException(string.Join("\n", e.Diagnostics.Select(d => $"{d.Line}:{d.Column} {d.Message}"))); }
    }
    private static WorkbookSnapshot Reopen(WorkbookSnapshot snapshot, string extension = "xlsq")
    {
        var folder = Path.Combine(AppContext.BaseDirectory, "p2", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "roundtrip." + extension); new TypedWorkbookStorage().Save(path, snapshot); return new TypedWorkbookStorage().Open(path);
    }

    [Theory] [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void ChartObjectsRetainIdentityFormattingAndStorage(ScriptLanguage language)
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet();s.getRange('B1:C3').setValues([['Name','Value'],['a',1],['b',2]]);
              const c=s.addChart(QuickExcel.ChartType.columnClustered,s.getRange('B1:C3'),QuickExcel.ChartSeriesBy.columns),id=c.getId();
              c.setName('Report'); c.getTitle().setText('Totals');c.getTitle().getFormat().getFont().setBold(true);
              c.getTitle().getFormat().getFont().setColor('#123456');c.getFormat().getFill().setSolidColor('#eeeeee');
              c.getLegend().setPosition(QuickExcel.ChartLegendPosition.bottom);c.setWidth(300);c.setHeight(200);c.setPosition('E3');
              const axis=c.getAxes().getValueAxis();axis.setMinimum(0);axis.setMaximum(10);axis.setNumberFormat('0.00');if(axis.getMaximum()!==10)throw Error('axis');
              const series=c.getSeries()[0];if(series.getName()!=='Value')throw Error('name reference');series.setName('Amount');series.getFormat().getFill().setSolidColor('#00aa00');
              s.getRange('D1:D2').setValues([[8],[9]]);series.setValues(s.getRange('D1:D2'));
              if(c.getId()!==id||s.getChart('Report').getId()!==id||c.getWidth()!==300||series.getName()!=='Amount'||!c.getTitle().getFormat().getFont().getBold())throw Error('chart state');
            }
            """, language);
        Run(doc); var chart = Assert.Single(doc.Snapshot.Sheets[0].Data.Charts); Assert.Equal("Totals", chart.Title); Assert.True(chart.Appearance.Title.Bold);
        var reopened = Reopen(doc.Snapshot); var saved = Assert.Single(reopened.Sheets[0].Data.Charts); Assert.Equal("Report", saved.Name); Assert.Equal("FF00AA00", saved.Series[0].Color);
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.Sheets[0].Data.Charts); Assert.True(doc.Redo()); Assert.Equal(chart.ScriptId, doc.Snapshot.Sheets[0].Data.Charts[0].ScriptId);
    }

    [Theory] [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void ConditionalRulesFormatsPriorityAndReferencesRoundtrip(ScriptLanguage language)
    {
        using var doc = Document("""
            function main(workbook){const r=workbook.getActiveWorksheet().getRange('B2:B4');r.setValues([[1],[5],[9]]);
              const a=r.addConditionalFormat(QuickExcel.ConditionalFormatType.cellValue),id=a.getId();
              a.getCellValue().setRule({formula1:'3',operator:QuickExcel.ConditionalCellValueOperator.greaterThan});
              a.getCellValue().getFormat().getFill().setColor('#ff0000');a.getCellValue().getFormat().getFont().setBold(true);a.setStopIfTrue(true);
              const border=a.getCellValue().getFormat().getConditionalRangeBorder(QuickExcel.ConditionalRangeBorderIndex.edgeBottom);
              border.setStyle(QuickExcel.ConditionalRangeBorderLineStyle.dash);border.setColor('#123456');if(border.getColor()!=='#123456')throw Error('border');
              const b=r.addConditionalFormat(QuickExcel.ConditionalFormatType.custom);b.getCustom().getRule().setFormula('=B2<2');
              b.getCustom().getFormat().getFont().setItalic(true);b.setPriority(1);
              if(a.getPriority()!==2||a.getId()!==id||r.getConditionalFormat(id).getId()!==id||r.getConditionalFormats().length!==2)throw Error('priority');
              const p=r.addConditionalFormat(QuickExcel.ConditionalFormatType.presetCriteria);p.getPreset().setRule({criterion:QuickExcel.ConditionalFormatPresetCriterion.uniqueValues});
              p.delete();
            }
            """, language);
        Run(doc); var rules = doc.Snapshot.Sheets[0].Data.ConditionalFormats; Assert.Equal(2, rules.Length);
        Assert.True(rules.Single(r => r.Kind == ConditionalFormatKind.CellIs).Style.Bold);
        var saved = Reopen(doc.Snapshot).Sheets[0].Data.ConditionalFormats; Assert.Equal(2, saved.Length); Assert.Equal("=(B2<2)", saved.Single(r => r.Kind == ConditionalFormatKind.Expression).Formula1);
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.Sheets[0].Data.ConditionalFormats);
    }

    [Theory] [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void ValidationRulesNestedRangeArgumentsAndInspectionRoundtrip(ScriptLanguage language)
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('B1:B2'),v=r.getDataValidation();
              r.setValues([[3],[9]]);v.setRule({wholeNumber:{formula1:1,formula2:5,operator:QuickExcel.DataValidationOperator.between}});
              v.setErrorAlert({title:'Range',message:'Use 1–5',showAlert:false,style:QuickExcel.DataValidationAlertStyle.stop});
              v.setPrompt({title:'Hint',message:'Enter a number',showPrompt:true});v.setIgnoreBlanks(false);
              if(v.getValid()||v.getInvalidCells().getAddress()!=='B2'||v.getType()!==QuickExcel.DataValidationType.wholeNumber)throw Error('validation');
              s.getRange('D1:D2').setValues([['a'],['b']]);const list=s.getRange('E1').getDataValidation();
              list.setRule({list:{source:s.getRange('D1:D2'),inCellDropDown:true}});list.setIgnoreBlanks(true);
              if(!list.getValid()||!v.getPrompt().showPrompt||v.getErrorAlert().showAlert)throw Error('options');
            }
            """, language);
        Run(doc); Assert.Equal(2, doc.Snapshot.Sheets[0].Data.Validations.Length);
        var saved = Reopen(doc.Snapshot).Sheets[0].Data.Validations; Assert.Equal(2, saved.Length); Assert.Equal("Range", saved.Single(v => v.Kind == ValidationKind.Whole).ErrorTitle);
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.Sheets[0].Data.Validations);
    }

    [Theory] [InlineData(ScriptLanguage.JavaScript)] [InlineData(ScriptLanguage.TypeScript)]
    public void ThreadCommentsRepliesResolutionAndStorageAreUndoable(ScriptLanguage language)
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),c=s.addComment('B2','Review',QuickExcel.ContentType.plain),id=c.getId();
              const reply=c.addCommentReply('Agreed');reply.setContent('Checked');c.setContent('Final review');c.setResolved(true);
              if(c.getId()!==id||!reply.getResolved()||reply.getParentComment().getId()!==id||workbook.getCommentByReplyId(reply.getId()).getId()!==id||c.getCreationDate().getFullYear()<2020)throw Error('comments');
              if(workbook.getComments().length!==1||c.getReplies()[0].getContent()!=='Checked'||s.getCommentByCell('B2').getId()!==id)throw Error('lookup');
            }
            """, language);
        Run(doc); var comments = doc.Snapshot.Sheets[0].Data.ThreadComments; Assert.Equal(2, comments.Length); Assert.Empty(doc.Snapshot.Sheets[0].Annotations.Notes);
        var reopened = Reopen(doc.Snapshot); Assert.Equal(comments.ToArray(), reopened.Sheets[0].Data.ThreadComments.ToArray());
        // Editing an imported plain thread is still transactional and persists its real thread IDs.
        using var imported = new WorkbookDocument(reopened, new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        imported.RunConsole("workbook.getActiveWorksheet().getCommentByCell('B2').setContent('Reopened');");
        Assert.Equal("Reopened", Reopen(imported.Snapshot).Sheets[0].Data.ThreadComments.Single(c => c.ParentId is null).Content);
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.Sheets[0].Data.ThreadComments); Assert.True(doc.Redo());
        imported.RunConsole("workbook.getActiveWorksheet().getCommentByCell('B2').delete();");
        Assert.Empty(Reopen(imported.Snapshot).Sheets[0].Data.ThreadComments);
    }

    [Fact]
    public void InvalidObjectWritesRollbackCellsFilesAndAllObjectsEvenWhenCaught()
    {
        using var doc = Document("""
            function main(workbook,host){const s=workbook.getActiveWorksheet();s.getRange('B1').setValue(7);workbook.files.writeText('//xlsq/p2.txt','staged');
              s.addComment('C1','staged');try{s.getRange('B1').getDataValidation().setRule({wholeNumber:{formula1:1,operator:999}});}catch(e){}
            }
            """);
        doc.EnableScriptsAndCalculate(); Assert.ThrowsAny<Exception>(() => doc.RunAction("main")); Assert.False(doc.CanUndo); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1)); Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Empty(doc.Snapshot.Sheets[0].Data.ThreadComments);
    }

    [Fact]
    public void DeletedChartAndConditionalHandlesNeverAliasReplacementObjects()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('B1:C2');r.setValues([['x','y'],[1,2]]);
              const a=s.addChart(QuickExcel.ChartType.line,r);a.delete();const b=s.addChart(QuickExcel.ChartType.line,r);
              let stale=false;try{a.getName()}catch(e){stale=e.code==='InvalidReference'}if(!stale||a===b)throw Error('chart identity');
              const cf=r.addConditionalFormat(QuickExcel.ConditionalFormatType.custom);cf.delete();r.addConditionalFormat(QuickExcel.ConditionalFormatType.custom);
              stale=false;try{cf.getPriority()}catch(e){stale=e.code==='InvalidReference'}if(!stale)throw Error('rule identity');
            }
            """);
        Run(doc); Assert.Single(doc.Snapshot.Sheets[0].Data.Charts); Assert.Single(doc.Snapshot.Sheets[0].Data.ConditionalFormats);
    }

    [Fact]
    public void PastedChartReceivesAnIndependentScriptIdentity()
    {
        var sheet = Sheet("Data", (0, 0, T("Name")), (0, 1, T("Value")), (1, 0, T("a")), (1, 1, N(1)));
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 1, 1), WorksheetChartKind.Column, new() { Kind = DrawingAnchorKind.OneCell, Width = 4_762_500, Height = 2_857_500 });
        var book = Book(sheet.WithData(sheet.Data with { Charts = [chart] })); using var session = new DocumentSession(book);
        session.Execute(new PasteChartCommand(ChartClipboard.Capture(book, chart), sheet.Id, new(5, 5)));
        var pasted = session.Snapshot.GetSheet(sheet.Id).Data.Charts[1]; Assert.NotEqual(chart.ScriptId, pasted.ScriptId);
        Assert.True(session.Undo()); Assert.True(session.Redo()); Assert.Equal(pasted.ScriptId, session.Snapshot.GetSheet(sheet.Id).Data.Charts[1].ScriptId);
    }

    [Theory]
    [InlineData("s.getRange('B1:B10001').addConditionalFormat(QuickExcel.ConditionalFormatType.custom).getCustom().getFormat().getFont().setBold(true)")]
    [InlineData("const v=s.getRange('B1:B10001').getDataValidation();v.setRule({list:{source:'a,b',inCellDropDown:true}});v.setIgnoreBlanks(false)")]
    public void ObjectRuleWritesUseCumulativeCellBudgets(string operations)
    {
        using var doc = Document("function main(workbook){const s=workbook.getActiveWorksheet();" + operations + ";}");
        doc.EnableScriptsAndCalculate(); var error = Assert.ThrowsAny<Exception>(() => doc.RunAction("main")); Assert.Contains("20000", error.Message); Assert.False(doc.CanUndo);
    }

    [Theory]
    [InlineData("s.getRange('B1').clearAllConditionalFormats()")]
    [InlineData("s.getRange('B1').getDataValidation().clear()")]
    public void PartialRuleRemovalNeverDeletesOutsideTheRequestedRange(string operation)
    {
        using var doc = Document("function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('B1:B3');r.addConditionalFormat(QuickExcel.ConditionalFormatType.custom);r.getDataValidation().setRule({list:{source:'a,b',inCellDropDown:true}});" + operation + ";}");
        doc.EnableScriptsAndCalculate(); Assert.ThrowsAny<Exception>(() => doc.RunAction("main")); Assert.False(doc.CanUndo); Assert.Empty(doc.Snapshot.Sheets[0].Data.Validations); Assert.Empty(doc.Snapshot.Sheets[0].Data.ConditionalFormats);
    }

    [Theory]
    [InlineData("s.addComment('D1','blocked')")]
    [InlineData("s.addChart(QuickExcel.ChartType.line,s.getRange('B1:C2'))")]
    public void ObjectProtectionCannotBeBypassedThroughTheScriptModel(string operation)
    {
        using var doc = Document("function main(workbook){const s=workbook.getActiveWorksheet();s.getRange('B1:C2').setValues([['x','y'],[1,2]]);s.getProtection().protect();" + operation + ";}");
        doc.EnableScriptsAndCalculate(); Assert.ThrowsAny<Exception>(() => doc.RunAction("main")); Assert.False(doc.CanUndo); Assert.Null(doc.Snapshot.Sheets[0].Data.Protection);
    }

    [Fact]
    public void TextConditionalRulesAndMultipleSheetCommentsSaveWithoutLosingNotes()
    {
        using var doc = Document("""
            function main(workbook){const s=workbook.getActiveWorksheet(),r=s.getRange('B1:B2');r.setValues([['prefix'],['other']]);
              const f=r.addConditionalFormat(QuickExcel.ConditionalFormatType.containsText).getTextComparison();
              f.setRule({operator:QuickExcel.ConditionalTextOperator.beginsWith,text:'pre'});f.getFormat().getFont().setColor('#123456');
              s.addComment('D1','sheet one');
            }
            """);
        Run(doc); var second = Sheet("Other");
        var original = doc.Snapshot; var snapshot = original.WithSheets(original.Sheets.Add(second.WithData(second.Data with {
            ThreadComments = [new WorksheetComment(Guid.NewGuid(), new(0, 3), "sheet two", "Local", DateTimeOffset.UtcNow, Guid.NewGuid())] })));
        var reopened = Reopen(snapshot); Assert.Equal(2, reopened.Sheets.Sum(s => s.Data.ThreadComments.Length)); Assert.Equal(ConditionalFormatKind.BeginsWith, reopened.Sheets[0].Data.ConditionalFormats[0].Kind);
        var path = Path.Combine(AppContext.BaseDirectory, "p2", Guid.NewGuid() + ".xlsx"); new TypedWorkbookStorage().ExportPortable(path, snapshot);
        Assert.Equal(2, new TypedWorkbookStorage().Open(path).Sheets.Sum(s => s.Data.ThreadComments.Length));
    }
}
