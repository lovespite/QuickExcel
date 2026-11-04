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
        new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) }, ["main"], scriptApiVersion: 2,
        bindings: functions ? [new("main", "formula-functions"), new("main", "actions")] : [new("main", "actions")]);
    private static WorkbookServices Services(IScriptActionExecutor? actions = null) => new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), actions ?? new ScriptWorkerActionExecutor());
    private static WorkbookDocument Document(string source, params WorksheetSnapshot[] sheets) => new(Book(sheets.Length == 0 ? [Sheet("S")] : sheets).WithScriptModules(Actions(source)), Services());
    private static ScriptActionResult Result(ScriptActionRequest request, params ScriptActionEdit[] edits) => new(request.OperationId, request.Stamp.InputRevision, request.Name!, edits.ToImmutableArray(), []);

    [Fact]
    public void RealActionReadsFrozenTypedSnapshotAndCommitsAcrossSheetsWithOneUndo()
    {
        var s = Sheet("Data", (0, 0, T("  产品  ")), (1, 0, N(2)), (2, 0, T("00123")), (3, 0, F("=1+1", 2)));
        var other = Sheet("Out");
        using var doc = Document("""
            quickExcel.actions.register({name:'规范商品名',description:'一次事务'}, ({snapshot,patch})=>{
              const a=snapshot.sheets[0].id,b=snapshot.sheets[1].id;
              if(!Object.isFrozen(snapshot) || !Object.isFrozen(snapshot.sheets) || !Object.isFrozen(snapshot.getRange(a,'A1:A3'))) throw 1;
              if(snapshot.getCell(a,'A4').formula!=='=1+1') throw 2;
              patch.setValues(a,'B1:B3',snapshot.getRange(a,'A1:A3').map(r=>r.map(v=>typeof v==='string'?v.trim():v)));
              patch.setValues(b,'A1:E1',[[true,null,'',3,{kind:'error',code:'DivisionByZero'}]]);
              patch.setFormula(b,'A2','=Data!B2*2');
            });
            """, s, other);
        Assert.Throws<InvalidOperationException>(() => doc.RunAction("规范商品名"));
        doc.EnableScriptsAndCalculate();
        var actions = doc.ListActions();
        Assert.Equal("一次事务", Assert.Single(actions).Description);
        var before = doc.Snapshot; var events = 0;
        doc.Session.Changed += (_, e) => { if (!e.CalculationOnly) events++; };
        var result = doc.RunAction("规范商品名");
        Assert.True(result.Changed); Assert.Equal(before.Revision + 1, doc.Snapshot.Revision); Assert.Equal(1, events);
        Assert.Equal(CellValue.FromText("产品"), doc.Snapshot.GetSheet(s.Id).GetCell(0, 1)!.Value);
        Assert.Equal(CellValue.FromText("00123"), doc.Snapshot.GetSheet(s.Id).GetCell(2, 1)!.Value);
        Assert.Equal(CellValue.FromBoolean(true), doc.Snapshot.GetSheet(other.Id).GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText(""), doc.Snapshot.GetSheet(other.Id).GetCell(0, 2)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.DivisionByZero), doc.Snapshot.GetSheet(other.Id).GetCell(0, 4)!.Value);
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
            quickExcel.functions.register('NEXT',()=>++n);
            quickExcel.actions.register('写入',({snapshot,patch})=>patch.setValue(snapshot.sheets[0].id,'B1',++n));
            """;
        using var doc = new WorkbookDocument(Book(Sheet("S", (0, 0, F("=NEXT()", 99)))).WithScriptModules(Actions(source, true)), Services());
        doc.EnableScriptsAndCalculate(); Assert.Equal("1", doc.GetCellValue(0, 0, 0));
        Assert.Single(doc.ListActions()); Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
        doc.RunAction("写入"); Assert.Equal("1", doc.GetCellValue(0, 0, 1));
        Assert.Equal(OperatingSystem.IsWindows(), doc.RunAction("写入").Changed);
        doc.Calculate(); Assert.Equal(OperatingSystem.IsWindows() ? "2" : "1", doc.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData("throw new Error('failure')")]
    [InlineData("while(true){}")]
    [InlineData("return Promise.resolve()")]
    [InlineData("Promise.resolve().then(()=>patch.setValue(id,'A2',7))")]
    [InlineData("patch.setValue(id,'A1',2)")]
    [InlineData("patch.setValue('missing','B1',3)")]
    [InlineData("patch.setValues(id,'B1:C1',[[1]])")]
    [InlineData("patch.setValue(id,'B1',NaN)")]
    [InlineData("patch.setValue(id,'B1',{})")]
    [InlineData("patch.setValue(id,'B1','x'.repeat(32768))")]
    [InlineData("patch.setValue(id,'A100001',1)")]
    [InlineData("patch.setValues(id,'B1:B20001',Array.from({length:20001},()=>[1]))")]
    [InlineData("snapshot.getRange(id,'A1:B100000')")]
    [InlineData("snapshot.sheets[0].name='changed'")]
    [InlineData("patch.setFormula(id,'B1','=1+')")]
    [InlineData("try { patch.setValues(id,'B1:C1',[[1,{}]]) } catch {}")]
    public void FailureAfterFirstWriteNeverCommitsPartialPatch(string tail)
    {
        using var doc = Document("quickExcel.actions.register('A',({snapshot,patch})=>{const id=snapshot.sheets[0].id;patch.setValue(id,'A1',1);" + tail + ";});", Sheet("S", (0, 0, N(42))));
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
        using var doc = Document("quickExcel.actions.register('A',({snapshot,patch})=>{patch.setValue(snapshot.sheets[0].id,'A1',9);while(true){}});", Sheet("S", (0, 0, N(42))));
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
            ["action"] = new("scripts/action.js", "import { next } from 'helper'; quickExcel.functions.register('UNBOUND',()=>0); quickExcel.actions.register('A',({snapshot,patch})=>{console.log('committed');patch.setValue(snapshot.sheets[0].id,'A1',next());});"),
            ["helper"] = new("scripts/helper.js", "let n=6; export const next=()=>++n;"),
            ["udf"] = new("scripts/udf.js", "import { next } from 'helper'; quickExcel.actions.register('unbound',()=>{}); quickExcel.functions.register('NEXTVALUE',()=>next());")
        }, ["udf", "action"], scriptApiVersion: 2, bindings: [new("udf", "formula-functions"), new("action", "actions")]);
        using var doc = new WorkbookDocument(Book(Sheet("S")).WithScriptModules(modules), Services());
        doc.Session.SetScriptsEnabled(true);
        Assert.Single(doc.ListActions()); doc.RunAction("A");
        Assert.Equal("7", doc.GetCellValue(0, 0, 0)); Assert.Contains("committed", doc.Session.Diagnostics);
        doc.SetCell(0, 1, 0, "=NEXTVALUE()"); doc.Calculate();
        Assert.Equal("7", doc.GetCellValue(0, 1, 0));
    }

    [Fact]
    public void SnapshotDistinguishesMissingCacheAndWritesPreserveStylesAndLiteralText()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S",
            [new(new(0, 0), new(CellContent.FromFormula("=1+2", 1)))], rows: [new(1, new(StyleId: 1))]);
        var source = """
            quickExcel.actions.register('A',({snapshot,patch})=>{
              const id=snapshot.sheets[0].id, c=snapshot.getCell(id,'A1');
              if(c.hasValue || c.value!==null || c.formula!=='=1+2' || c.state==='Fresh') throw 1;
              if(!snapshot.getCell(id,'B1').hasValue) throw 2;
              patch.setValue(id,'A1','=literal'); patch.setValue(id,'A2','00123');
            });
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
        using var doc = Document("quickExcel.actions.register('A',({snapshot,patch})=>{patch.setValue(snapshot.sheets[0].id,'A1',1);patch.setValue(snapshot.sheets[1].id,'B1',2);});", Sheet("Safe"), target);
        doc.Session.SetScriptsEnabled(true);
        Assert.Throws<ScriptRuntimeFailureException>(() => doc.RunAction("A"));
        Assert.Empty(doc.Snapshot.Sheets[0].Cells); Assert.False(doc.Session.CanUndo);
    }

    [Fact]
    public void ActionsRoundtripFeatureBindingsAndSourceWithoutExecutionOrAuthorization()
    {
        var path = Path.Combine(Path.GetTempPath(), "qx-actions-" + Guid.NewGuid() + ".xlsq");
        var modules = Actions("quickExcel.functions.register('F',()=>7);quickExcel.actions.register('A',()=>{throw 1;});", true);
        try
        {
            using (var doc = new WorkbookDocument(Book(Sheet("S")).WithScriptModules(modules), Services())) doc.Save(path);
            using (var zip = ZipFile.OpenRead(path))
            using (var reader = new StreamReader(zip.GetEntry("manifest.json")!.Open()))
            {
                var manifest = JsonNode.Parse(reader.ReadToEnd())!;
                Assert.Equal(2, manifest["scriptApiVersion"]!.GetValue<int>());
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
