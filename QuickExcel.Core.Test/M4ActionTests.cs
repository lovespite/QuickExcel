using System.Collections.Immutable;
using System.IO.Compression;
using System.Text.Json.Nodes;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4ActionTests
{
    internal static ScriptModuleSet Actions(string source, bool functions = false) => new(
        new Dictionary<string, ScriptModule> { ["A"] = new("scripts/main.js", source) }, ["A"], scriptApiVersion: 3,
        bindings: functions ? [new("A", "formula-functions"), new("A", "actions")] : [new("A", "actions")]);
    private static WorkbookServices Services(IScriptActionExecutor? actions = null) => new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), actions ?? new ScriptWorkerActionExecutor());
    private static WorkbookDocument Document(string source, params WorksheetSnapshot[] sheets) => new(Book(sheets.Length == 0 ? [Sheet("S")] : sheets).WithScriptModules(Actions(source)), Services());
    private static ScriptActionResult Result(ScriptActionRequest request, params ScriptActionEdit[] edits) => new(request.OperationId, request.Stamp.InputRevision, request.Name!, edits.ToImmutableArray(), []);

    [Fact]
    public void RealActionReadsLiveModelAndCommitsAcrossSheetsWithOneUndo()
    {
        var s = Sheet("Data", (0, 0, T("  产品  ")), (1, 0, N(2)), (2, 0, T("00123")), (3, 0, F("=1+1", 2)));
        var other = Sheet("Out");
        using var doc = Document("""
            function main(workbook) {
              const a=workbook.getWorksheet('Data'),b=workbook.getWorksheet('Out');
              if(a.getRange('A4').getFormula()!=='=1+1') throw 2;
              a.getRange('B1:B3').setValues(a.getRange('A1:A3').getValues().map(r=>r.map(v=>typeof v==='string'?v.trim():v)));
              b.getRange('A1:C1').setValues([[true,'',3]]);
              b.getRange('A2').setFormula('=Data!B2*2');
            }
            """, s, other);
        Assert.Throws<InvalidOperationException>(() => doc.RunAction("A"));
        doc.EnableScriptsAndCalculate();
        var actions = doc.ListActions();
        Assert.Equal("A", Assert.Single(actions).Name);
        var before = doc.Snapshot; var events = 0;
        doc.Session.Changed += (_, e) => { if (!e.CalculationOnly) events++; };
        var result = doc.RunAction("A");
        Assert.True(result.Changed); Assert.Equal(before.Revision + 1, doc.Snapshot.Revision); Assert.Equal(1, events);
        Assert.Equal(CellValue.FromText("产品"), doc.Snapshot.GetSheet(s.Id).GetCell(0, 1)!.Value);
        Assert.Equal(CellValue.FromText("00123"), doc.Snapshot.GetSheet(s.Id).GetCell(2, 1)!.Value);
        Assert.Equal(CellValue.FromBoolean(true), doc.Snapshot.GetSheet(other.Id).GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText(""), doc.Snapshot.GetSheet(other.Id).GetCell(0, 1)!.Value);
        Assert.Equal(CellValue.FromNumber(3), doc.Snapshot.GetSheet(other.Id).GetCell(0, 2)!.Value);
        Assert.Equal("=Data!B2*2", doc.Snapshot.GetSheet(other.Id).GetCell(1, 0)!.Content.FormulaText);
        Assert.True(doc.Undo()); Assert.False(doc.Session.CanUndo);
        Assert.Null(doc.Snapshot.GetSheet(s.Id).GetCell(0, 1)); Assert.Empty(doc.Snapshot.GetSheet(other.Id).Cells);
        doc.DisableScripts(); // Redo is recorded data, with no runtime or renewed authorization.
        Assert.True(doc.Redo()); Assert.False(doc.ScriptsEnabled);
        Assert.Equal("产品", doc.GetCellValue(0, 0, 1)); Assert.Equal(3, events);
    }

    [Fact]
    public void WindowsWorkbookFunctionsAndActionsReuseSeparateStateAcrossRounds()
    {
        var source = """
            let n=0;
            host.functions.register('NEXT',()=>++n);
            function main(workbook){ workbook.getActiveWorksheet().getRange('B1').setValue(++n); }
            """;
        using var doc = new WorkbookDocument(Book(Sheet("S", (0, 0, F("=NEXT()", 99)))).WithScriptModules(Actions(source, true)), Services());
        doc.EnableScriptsAndCalculate(); Assert.Equal("1", doc.GetCellValue(0, 0, 0));
        Assert.Single(doc.ListActions()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
        doc.RunAction("A"); Assert.Equal("1", doc.GetCellValue(0, 0, 1));
        Assert.Equal(OperatingSystem.IsWindows(), doc.RunAction("A").Changed);
        doc.Calculate(); Assert.Equal(OperatingSystem.IsWindows() ? "2" : "1", doc.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData("throw new Error('failure')")]
    [InlineData("while(true){}")]
    [InlineData("return Promise.resolve()")]
    [InlineData("Promise.resolve().then(()=>sheet.getRange('A2').setValue(7))")]
    [InlineData("sheet.getRange('A200001').setValue(2)")]
    [InlineData("workbook.getWorksheet('missing').getRange('B1')")]
    [InlineData("sheet.getRange('B1:C1').setValues([[1]])")]
    [InlineData("sheet.getRange('B1').setValue(NaN)")]
    [InlineData("sheet.getRange('B1').setValue({})")]
    [InlineData("sheet.getRange('B1').setValue('x'.repeat(32768))")]
    [InlineData("sheet.getRange('B1:B20001').setValues(Array.from({length:20001},()=>[1]))")]
    [InlineData("sheet.getRange('A1:B100000').getValues()")]
    [InlineData("sheet.name='changed'")]
    [InlineData("sheet.getRange('B1').setFormula('=1+')")]
    [InlineData("try { sheet.getRange('B1:C1').setValues([[1,{}]]) } catch {}")]
    public void FailureAfterFirstWriteNeverCommitsPartialPatch(string tail)
    {
        using var doc = Document("function main(workbook){const sheet=workbook.getActiveWorksheet();sheet.getRange('A1').setValue(1);" + tail + ";}", Sheet("S", (0, 0, N(42))));
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        Assert.NotNull(Record.Exception(() => doc.RunAction("A")));
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.Session.CanUndo);
        doc.SetCell(0, 1, 0, "editable"); Assert.Equal("editable", doc.GetCellValue(0, 1, 0));
    }

    [Theory]
    [InlineData("duplicate")] [InlineData("sheet")] [InlineData("coordinate")] [InlineData("text")]
    [InlineData("both")] [InlineData("neither")] [InlineData("formula")] [InlineData("count")]
    [InlineData("operation")] [InlineData("revision")] [InlineData("name")]
    public void HostRevalidatesEveryProposedEditBeforeMutation(string fault)
    {
        var sheet = Sheet("S", (0, 0, N(42)));
        using var session = new DocumentSession(Book(sheet).WithScriptModules(Actions("")));
        session.SetScriptsEnabled(true);
        var request = session.BeginAction("A"); var before = session.Snapshot;
        var good = new ScriptActionEdit(new(sheet.Id, 0, 0), CellValue.FromNumber(1), null);
        var bad = new ScriptActionEdit(new(sheet.Id, 1, 0), CellValue.FromNumber(2), null);
        bad = fault switch
        {
            "duplicate" => good,
            "sheet" => bad with { Address = new(SheetId.New(), 1, 0) },
            "coordinate" => bad with { Address = new(sheet.Id, -1, 0) },
            "text" => bad with { Value = CellValue.FromText(new string('x', 32768)) },
            "both" => bad with { Formula = "=1" },
            "neither" => bad with { Value = null },
            "formula" => bad with { Value = null, Formula = "=1+" },
            _ => bad
        };
        var result = Result(request, good, bad);
        result = fault switch
        {
            "operation" => result with { OperationId = Guid.NewGuid() },
            "revision" => result with { BaseRevision = result.BaseRevision + 1 },
            "name" => result with { Name = "other" },
            "count" => result with { Edits = Enumerable.Repeat(good, 20001).ToImmutableArray() },
            _ => result
        };
        Assert.NotNull(Record.Exception(() => session.CommitAction(request, result)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData("edit")] [InlineData("source")] [InlineData("binding")] [InlineData("disable")]
    [InlineData("new-action")] [InlineData("calculate")] [InlineData("close")] [InlineData("cancel")]
    public void RevisionAuthorizationAndGenerationChangesRejectLatePatch(string change)
    {
        var sheet = Sheet("S", (0, 0, N(42)));
        using var session = new DocumentSession(Book(sheet).WithScriptModules(Actions("")));
        session.SetScriptsEnabled(true); var request = session.BeginAction("A");
        var result = Result(request, new ScriptActionEdit(new(sheet.Id, 0, 0), CellValue.FromNumber(9), null));
        using var cancel = new CancellationTokenSource();
        switch (change)
        {
            case "edit": session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 0), CellContent.FromLiteral(CellValue.FromNumber(7)))])); break;
            case "source": session.Execute(new UpdateScriptSourceCommand("// changed")); break;
            case "binding": session.Execute(new UpdateScriptModulesCommand(Actions("", true))); break;
            case "disable": session.SetScriptsEnabled(false); break;
            case "new-action": session.BeginAction("A"); break;
            case "calculate": session.BeginCalculation(); break;
            case "close": session.Dispose(); break;
            case "cancel": cancel.Cancel(); break;
        }
        Assert.NotNull(Record.Exception(() => session.CommitAction(request, result, cancel.Token)));
        if (!session.IsClosed) Assert.Equal(CellValue.FromNumber(42), session.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Value);
    }

    [Fact]
    public void SuccessfulAndEmptyPatchesAreSingleUseAndHistoryBudgetIsAtomic()
    {
        var sheet = Sheet("S", (0, 0, N(42)));
        using var session = new DocumentSession(Book(sheet).WithScriptModules(Actions("")));
        session.SetScriptsEnabled(true); var request = session.BeginAction("A");
        Assert.False(session.CommitAction(request, Result(request)).Changed); Assert.False(session.CanUndo);
        Assert.NotNull(Record.Exception(() => session.CommitAction(request, Result(request))));
        request = session.BeginAction("A"); var patch = Result(request, new ScriptActionEdit(new(sheet.Id, 0, 0), CellValue.FromNumber(7), null));
        Assert.True(session.CommitAction(request, patch).Changed);
        Assert.NotNull(Record.Exception(() => session.CommitAction(request, patch)));
        using var tiny = new DocumentSession(Book(sheet).WithScriptModules(Actions("")), historyBudgetBytes: 1);
        tiny.SetScriptsEnabled(true); var tinyRequest = tiny.BeginAction("A"); var before = tiny.Snapshot;
        Assert.Throws<InvalidOperationException>(() => tiny.CommitAction(tinyRequest, Result(tinyRequest, patch.Edits.ToArray())));
        Assert.Same(before, tiny.Snapshot); Assert.False(tiny.CanUndo);
    }

    [Fact]
    public async Task CancelRunningWorkerLeavesWorkbookEditable()
    {
        using var doc = Document("function main(workbook){workbook.getActiveWorksheet().getRange('A1').setValue(9);while(true){}}", Sheet("S", (0, 0, N(42))));
        doc.EnableScriptsAndCalculate(); using var cancel = new CancellationTokenSource(150);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Task.Run(() => doc.RunAction("A", cancel.Token)));
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.False(doc.Session.CanUndo);
        doc.SetCell(0, 1, 0, "available");
    }

    [Theory]
    [InlineData("duplicate")] [InlineData("sheet")] [InlineData("coordinate")] [InlineData("type")]
    [InlineData("both")] [InlineData("formula")] [InlineData("error")] [InlineData("operation")]
    [InlineData("base")] [InlineData("name")] [InlineData("count")] [InlineData("session")]
    [InlineData("worker")] [InlineData("request")] [InlineData("revision")] [InlineData("generation")]
    [InlineData("calculation")]
    [InlineData("model-operation")] [InlineData("model-sequence")] [InlineData("model-round")]
    [InlineData("model-method")] [InlineData("model-sheet")] [InlineData("model-value")]
    public void FaultyWorkerCannotBypassProtocolOrHostValidation(string fault)
    {
        var executor = new ScriptWorkerActionExecutor(Path.Combine(AppContext.BaseDirectory, "WorkerFixture", "WorkerFixture.exe"));
        using var doc = new WorkbookDocument(Book(Sheet("S", (0, 0, N(42)))).WithScriptModules(Actions(fault)), Services(executor));
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        Assert.NotNull(Record.Exception(() => doc.RunAction("A")));
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.Session.CanUndo);
    }

    [Fact]
    public void SharedModuleDependenciesInitializeSeparatelyForEachRoleAndRespectBindings()
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["action"] = new("scripts/action.js", "import { next } from 'helper'; host.functions.register('UNBOUND',()=>0); export function main(workbook){console.log('committed');workbook.getActiveWorksheet().getRange('A1').setValue(next());}"),
            ["helper"] = new("scripts/helper.js", "let n=6; export const next=()=>++n;"),
            ["udf"] = new("scripts/udf.js", "import { next } from 'helper'; host.functions.register('NEXTVALUE',()=>next());")
        }, ["udf", "action"], scriptApiVersion: 3, bindings: [new("udf", "formula-functions"), new("action", "actions")]);
        using var doc = new WorkbookDocument(Book(Sheet("S")).WithScriptModules(modules), Services());
        doc.Session.SetScriptsEnabled(true);
        Assert.Equal("action", Assert.Single(doc.ListActions()).Name); doc.RunAction("action");
        Assert.Equal("7", doc.GetCellValue(0, 0, 0)); Assert.Contains("committed", doc.Session.Diagnostics);
        doc.SetCell(0, 1, 0, "=NEXTVALUE()"); doc.Calculate();
        Assert.Equal("7", doc.GetCellValue(0, 1, 0));
    }

    [Fact]
    public void LiveFormulaReadCalculatesMissingCacheAndWritesPreserveStylesAndLiteralText()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S",
            [new(new(0, 0), new(CellContent.FromFormula("=1+2", 1)))], rows: [new(1, new(StyleId: 1))]);
        var source = """
            function main(workbook){
              const sheet=workbook.getActiveWorksheet();
              if(sheet.getRange('A1').getValue()!==3) throw 1;
              if(sheet.getRange('A1').getFormula()!=='=1+2') throw 2;
              if(sheet.getRange('B1').getValue()!=='') throw 3;
              sheet.getRange('A1').setValue('=literal'); sheet.getRange('A2').setValue('00123');
            }
            """;
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904, [sheet], source,
            [new(), new() { NumberFormatCode = "@" }], Actions(source));
        using var doc = new WorkbookDocument(snapshot, Services()); doc.Session.SetScriptsEnabled(true);
        doc.RunAction("A");
        Assert.Equal(CellContentKind.Literal, doc.Snapshot.Sheets[0].GetCell(0, 0)!.Content.Kind);
        Assert.Equal("=literal", doc.GetCellValue(0, 0, 0));
        Assert.Equal(1, doc.Snapshot.Sheets[0].GetCell(0, 0)!.Content.StyleId);
        Assert.Equal(1, doc.Snapshot.Sheets[0].GetCell(1, 0)!.Content.StyleId);
        Assert.True(doc.Undo()); Assert.Equal("=1+2", doc.Snapshot.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void MergedChildOrInCellImageTargetsRejectTheWholePatch(bool image)
    {
        var target = new WorksheetSnapshot(SheetId.New(), "Protected", mergedCells: image ? [] : ["A1:B1"],
            images: image ? [new SheetImage([1], "png", 0, 1, placeInCell: true)] : []);
        using var doc = Document("function main(workbook){const sheets=workbook.getWorksheets();sheets[0].getRange('A1').setValue(1);sheets[1].getRange('B1').setValue(2);}", Sheet("Safe"), target);
        doc.Session.SetScriptsEnabled(true);
        Assert.Throws<ScriptRuntimeFailureException>(() => doc.RunAction("A"));
        Assert.Empty(doc.Snapshot.Sheets[0].Cells); Assert.False(doc.Session.CanUndo);
    }

    [Fact]
    public void ActionsRoundtripFeatureBindingsAndSourceWithoutExecutionOrAuthorization()
    {
        var path = Path.Combine(Path.GetTempPath(), "qx-actions-" + Guid.NewGuid() + ".xlsq");
        var modules = Actions("host.functions.register('F',()=>7);function main(){throw 1;}", true);
        try
        {
            using (var doc = new WorkbookDocument(Book(Sheet("S")).WithScriptModules(modules), Services())) doc.Save(path);
            using (var zip = ZipFile.OpenRead(path))
            using (var reader = new StreamReader(zip.GetEntry("manifest.json")!.Open()))
            {
                var manifest = JsonNode.Parse(reader.ReadToEnd())!;
                Assert.Equal(3, manifest["scriptApiVersion"]!.GetValue<int>());
                Assert.Contains("script-actions", manifest["requiredFeatures"]!.AsArray().Select(v => v!.GetValue<string>()));
            }
            using var reopened = WorkbookDocument.Open(path, Services());
            Assert.False(reopened.ScriptsEnabled); Assert.Equal(modules.Bindings.ToArray(), reopened.Snapshot.ScriptModules.Bindings.ToArray());
            Assert.Throws<InvalidOperationException>(() => reopened.ListActions());
            reopened.EnableScriptsAndCalculate(); Assert.Single(reopened.ListActions()); // body is not run
            reopened.ScriptSource += "\n// edit"; Assert.False(reopened.ScriptsEnabled);
            reopened.Save(path); Assert.True(Services().Storage.Open(path).ScriptModules.HasActions);
        }
        finally { File.Delete(path); }
    }
}
