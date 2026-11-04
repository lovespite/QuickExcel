using System.Collections.Immutable;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptWorkbookModelTests
{
    private static WorkbookDocument Document(string source, ScriptLanguage language = ScriptLanguage.JavaScript, IWorkbookStorage? storage = null)
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        { ["main"] = new(language == ScriptLanguage.TypeScript ? "scripts/main.ts" : "scripts/main.js", source, language) },
            ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);
        return new(Book(Sheet("Data", (0, 0, N(42)))).WithScriptModules(modules),
            new(storage ?? new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    }
    private static string? Fill(WorkbookSnapshot book, int row, int column) => book.Styles[book.Sheets[0].GetStyleId(row, column)].FillColor;

    [Theory]
    [InlineData(ScriptLanguage.JavaScript)]
    [InlineData(ScriptLanguage.TypeScript)]
    public void MainEntrySupportsActiveCellWorksheetRangeAndSequentialFillWithOneUndo(ScriptLanguage language)
    {
        var parameter = language == ScriptLanguage.TypeScript ? "workbook: QuickExcel.Workbook" : "workbook";
        using var doc = Document("function main(" + parameter + ") {" + """
            const selectedCell=workbook.getActiveCell(), selectedSheet=workbook.getActiveWorksheet(), selectedRange=workbook.getSelectedRange();
            if(selectedCell.getAddress() !== 'B2' || selectedRange.getAddress() !== 'B2:C3') throw Error('selection');
            if(selectedSheet.getName() !== 'Data' || selectedSheet.getCell('A1').getValue() !== 42) throw Error('read');
            selectedCell.getFormat().getFill().setColor('#ff0000');
            if(selectedCell.getFormat().getFill().getColor() !== '#FF0000') throw Error('first write');
            selectedRange.getFormat().getFill().setColor('#00ff00');
            if(selectedCell.getFormat().getFill().getColor() !== '#00FF00') throw Error('overlap');
            console.info(selectedSheet.getCell(0,0).getValue());
            }
            """, language);
        doc.EnableScriptsAndCalculate();
        doc.Session.UpdateNavigation(doc.Snapshot.ActiveSheetId, new(1, 1));
        doc.Session.UpdateScriptSelection(doc.Snapshot.ActiveSheetId, new(1, 1), [new(1, 1, 2, 2)]);
        Assert.Equal("main", Assert.Single(doc.ListActions()).Name);
        var revision = doc.Snapshot.Revision;
        Assert.True(doc.RunAction("main").Changed);
        Assert.Equal(revision + 1, doc.Snapshot.Revision);
        Assert.Equal("FF00FF00", Fill(doc.Snapshot, 1, 1)); Assert.Equal("FF00FF00", Fill(doc.Snapshot, 2, 2));
        Assert.Equal("42", doc.GetCellValue(0, 0, 0));
        Assert.True(doc.Undo()); Assert.False(doc.CanUndo); Assert.Null(Fill(doc.Snapshot, 1, 1));
        Assert.True(doc.Redo()); Assert.Equal("FF00FF00", Fill(doc.Snapshot, 2, 2));
    }

    [Fact]
    public void ModelReadsPriorWritesAndFilesStayLazy()
    {
        var storage = new CountingStorage();
        using var doc = Document("""
            function main(workbook, host){
              const sheet=workbook.getWorksheet('Data'), cell=sheet.getCell('A1');
              cell.setValue(7);
              if(cell.getValue()!==7) throw Error('view');
              sheet.getRange('B1').setValue(8);
              if(sheet.getCell('B1').getValue()!==8) throw Error('patch order');
              sheet.getRange('B1:C2').setValues([[9,10],[11,12]]);
              if(sheet.getCell('B1').getValue()!==9 || sheet.getRange('B1:C2').getValues()[1][1]!==12) throw Error('matrix');
              host.files.writeText('//mem/seen','yes');
            }
            """, storage: storage);
        doc.EnableScriptsAndCalculate(); doc.RunAction("main");
        Assert.Equal(0, storage.Captures);
        Assert.Equal("7", doc.GetCellValue(0, 0, 0)); Assert.Equal("9", doc.GetCellValue(0, 0, 1));
        Assert.True(doc.Undo()); Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
    }

    [Theory]
    [InlineData("throw Error('abort')")]
    [InlineData("try { cell.setValue(NaN); } catch {}")]
    [InlineData("try { cell.getFormat().getFill().setColor('red'); } catch {}")]
    [InlineData("try { cell.setValues([[1,2]]); } catch {}")]
    [InlineData("try { cell.setFormula('=1+'); } catch {}")]
    [InlineData("while(true){}")]
    public void FailureRollsBackValuesAndFormattingEvenWhenAnInvalidWriteIsCaught(string failure)
    {
        using var doc = Document("function main(workbook){const cell=workbook.getActiveWorksheet().getCell('A1');cell.setValue(7);cell.getFormat().getFill().setColor('#ff0000');" + failure + ";}");
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        Assert.NotNull(Record.Exception(() => doc.RunAction("main")));
        Assert.Equal(revision, doc.Snapshot.Revision); Assert.Equal("42", doc.GetCellValue(0, 0, 0));
        Assert.Null(Fill(doc.Snapshot, 0, 0)); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void LiveFormulaReadsCalculatePendingInputsAndNewFormulasBeforeCommit()
    {
        using var doc = Document("""
            function main(workbook) {
              const sheet=workbook.getActiveWorksheet(), formula=sheet.getCell('B1');
              if(formula.getValue()!==43) throw Error('initial calculation');
              sheet.getCell('A1').setValue(7);
              if(formula.getFormula()!=='=A1+1') throw Error('formula text');
              if(formula.getValue()!==8 || formula.getValues()[0][0]!==8) throw Error('pending calculation');
              formula.setFormula('=A1+2');
              if(formula.getFormula()!=='=A1+2') throw Error('pending formula text');
              if(formula.getValue()!==9) throw Error('new formula');
              sheet.getCell('A1').setValue(10);
              if(formula.getValue()!==12) throw Error('second invalidation');
            }
            """);
        doc.SetCell(0, 0, 1, "=A1+1");
        doc.EnableScriptsAndCalculate(); doc.RunAction("main"); doc.Calculate();
        Assert.Equal("12", doc.GetCellValue(0, 0, 1));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FormulaReadAfterWriteUsesTheDraftAndPreservesAtomicFilesHistoryAndSnapshots(bool udf, bool fail)
    {
        var source = "host.functions.register('TWICE',n=>n*2);" + """
            function main(workbook, host) {
                const sheet=workbook.getActiveWorksheet(), cell=sheet.getCell('a1');
                cell.setValue(Number(cell.getValue())+1);
                const result=sheet.getCell('a2').getValue();
                console.log(result);
                if(result!==12 || sheet.getRange('A1:A2').getValues()[1][0]!==12) throw Error('not current');
                sheet.getCell('B1').setValue(result);
                host.files.writeText('//xlsq/result.txt',String(result));
                host.files.writeText('//mem/result',String(result));
            """ + (fail ? "throw Error('after calculation');" : "") + "}";
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, N(5)), (1, 0, F(udf ? "=TWICE(A1)" : "=A1*2"))))
            .WithScriptModules(M4ActionTests.Actions(source, udf)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        if (udf && fail) doc.Session.SetScriptsEnabled(true); // Also cover a cold formula worker started during the action.
        else doc.EnableScriptsAndCalculate();
        var revision = doc.Snapshot.Revision;
        var changes = new List<DocumentChange>(); doc.Session.Changed += (_, change) => changes.Add(change);
        if (fail)
        {
            Assert.Throws<ScriptRuntimeFailureException>(() => doc.RunAction("A"));
            Assert.Equal(revision, doc.Snapshot.Revision); Assert.Equal("5", doc.GetCellValue(0, 0, 0));
            Assert.Equal("", doc.GetCellValue(0, 0, 1)); Assert.False(doc.CanUndo);
            Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Empty(doc.Session.TransientFiles.Entries);
            Assert.DoesNotContain(changes, c => !c.CalculationOnly);
        }
        else
        {
            doc.RunAction("A");
            Assert.Equal(revision + 1, doc.Snapshot.Revision); Assert.Single(changes);
            Assert.Equal("6", doc.GetCellValue(0, 0, 0)); Assert.Equal("12", doc.GetCellValue(0, 0, 1));
            Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/result.txt"));
            Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "12");
            Assert.True(doc.Undo()); Assert.False(doc.CanUndo); Assert.Empty(doc.Snapshot.PackageFiles.Entries);
            Assert.Equal("5", doc.GetCellValue(0, 0, 0)); Assert.Equal("", doc.GetCellValue(0, 0, 1));
        }
    }

    [Fact]
    public void FormulaValueReadsIncludeNewAndResizedSpillsAndTypedErrors()
    {
        using var doc = Document("""
            function main(workbook) {
                const s=workbook.getActiveWorksheet();
                s.getCell('B1').setFormula('=SEQUENCE(3)');
                if(s.getCell('B3').getValue()!==3) throw Error('new spill');
                s.getCell('B1').setFormula('=SEQUENCE(2)');
                const values=s.getRange('B1:B3').getValues();
                if(values[0][0]!==1 || values[1][0]!==2 || values[2][0]!=='') throw Error('shrinking spill');
                s.getCell('C1').setFormula('=1/0');
                if(s.getCell('C1').getValue()!=='#DIV/0!' || s.getCell('C1').getValues()[0][0]!=='#DIV/0!') throw Error('typed error');
                s.getCell('D1').setValue(values[1][0]);
            }
            """);
        doc.EnableScriptsAndCalculate(); doc.RunAction("main");
        Assert.Equal("2", doc.GetCellValue(0, 0, 3));
    }

    [Fact]
    public void UdfTimeoutDuringFormulaReadRollsBackEvenIfTheScriptTriesToCatchIt()
    {
        const string source = """
            host.functions.register('SLOW',n=>{if(n===6)while(true){};return n*2;});
            function main(workbook){
                const s=workbook.getActiveWorksheet();s.getCell('A1').setValue(6);
                try{s.getCell('A2').getValue();}catch{}
                s.getCell('B1').setValue(99);
            }
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, N(5)), (1, 0, F("=SLOW(A1)"))))
            .WithScriptModules(M4ActionTests.Actions(source, true)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        Assert.Throws<TimeoutException>(() => doc.RunAction("main"));
        Assert.Equal(revision, doc.Snapshot.Revision); Assert.Equal("5", doc.GetCellValue(0, 0, 0));
        Assert.Equal("", doc.GetCellValue(0, 0, 1)); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void ConsoleCalculatesBuiltinsWithoutLoadingStoredUdfModules()
    {
        using var doc = Document("throw Error('stored modules must not execute');");
        doc.RunConsole("const s=workbook.getActiveWorksheet();s.getCell('B1').setFormula('=A1*2');s.getCell('C1').setValue(s.getCell('B1').getValue());");
        Assert.Equal("84", doc.GetCellValue(0, 0, 2)); Assert.False(doc.ScriptsEnabled);
    }

    [Fact]
    public void BlankReadResolvesAnInitialUncalculatedSpillBeforeAnyWrite()
    {
        const string source = """
            function main(workbook){
                const s=workbook.getActiveWorksheet();
                if(s.getCell('B2').getValue()!==2) throw Error('initial spill');
                s.getCell('A1').setValue(3);
                if(s.getCell('B3').getValue()!==3) throw Error('grown spill');
            }
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, N(2)), (0, 1, F("=SEQUENCE(A1)"))))
            .WithScriptModules(M4ActionTests.Actions(source)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.Session.SetScriptsEnabled(true); doc.RunAction("main");
        Assert.Equal("3", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void BlockedFormulaReadsNeverReturnOldCachesAndFormulaTextDoesNotCalculate()
    {
        const string source = """
            function main(workbook){
                const s=workbook.getActiveWorksheet(), f=s.getCell('B1');
                s.getCell('A1').setValue(6);
                if(f.getFormula()!=='=MISSING(A1)') throw Error('formula text');
                let rejected=0;
                try{f.getValue();}catch{rejected++;}
                try{f.getValues();}catch{rejected++;}
                if(rejected!==2) throw Error('old result escaped');
            }
            """;
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0, 0, N(5)), (0, 1, F("=MISSING(A1)", 123))))
            .WithScriptModules(M4ActionTests.Actions(source)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.Session.SetScriptsEnabled(true); doc.RunAction("main");
        Assert.Equal("6", doc.GetCellValue(0, 0, 0));
        Assert.NotEqual(CalculationState.Fresh, doc.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State);
    }

    [Fact]
    public void MultipleAreasAreNotSilentlyExpandedToBoundingRectangle()
    {
        using var doc = Document("""
            function main(workbook) {
              let rejected=false; try { workbook.getSelectedRange(); } catch { rejected=true; }
              if(!rejected) throw Error('multiple areas');
              for(const range of workbook.getSelectedRanges().getAreas()) range.getFormat().getFill().setColor('#ff0000');
            }
            """);
        doc.EnableScriptsAndCalculate();
        doc.Session.UpdateScriptSelection(doc.Snapshot.ActiveSheetId, new(0, 0), [new(0, 0, 0, 0), new(2, 2, 2, 2)]);
        doc.RunAction("main");
        Assert.Equal("FFFF0000", Fill(doc.Snapshot, 0, 0)); Assert.Equal("FFFF0000", Fill(doc.Snapshot, 2, 2)); Assert.Null(Fill(doc.Snapshot, 1, 1));
    }

    [Fact]
    public void SupersededActionCannotReadOrCommitItsStagedModel()
    {
        using var doc = Document("function main(){}"); doc.EnableScriptsAndCalculate();
        var request = doc.Session.BeginAction("main");
        request.Model!.SetFill(doc.Snapshot.ActiveSheetId, new(0, 0, 0, 0), "#ff0000");
        doc.SetCell(0, 1, 0, "user edit");
        Assert.NotNull(Record.Exception(() => request.Model.ReadCell(doc.Snapshot.ActiveSheetId, new(0, 0))));
        Assert.NotNull(Record.Exception(() => doc.Session.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, "main", [], []))));
        Assert.Null(Fill(doc.Snapshot, 0, 0)); Assert.Equal("user edit", doc.GetCellValue(0, 1, 0));
    }

    private sealed class CountingStorage : IWorkbookStorage, IWorkbookVirtualFiles
    {
        private readonly TypedWorkbookStorage _storage = new();
        public int Captures { get; private set; }
        public VirtualFileSet Capture(WorkbookSnapshot snapshot) { Captures++; return _storage.Capture(snapshot); }
        public WorkbookSnapshot Apply(WorkbookSnapshot snapshot, IReadOnlyList<VirtualFileEdit> edits) => _storage.Apply(snapshot, edits);
        public WorkbookSnapshot Open(string path) => _storage.Open(path);
        public void Save(string path, WorkbookSnapshot snapshot) => _storage.Save(path, snapshot);
        public void Export(string path, WorkbookSnapshot snapshot) => _storage.Export(path, snapshot);
    }
}
