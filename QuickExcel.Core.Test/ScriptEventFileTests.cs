using System.Collections.Immutable;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptEventFileTests
{
    private static readonly TypedWorkbookStorage Storage = new();
    private static WorkbookDocument Document(string source) => new(Book(Sheet("Data")).WithScriptModules(M4ActionTests.Actions(source)),
        new(Storage, new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static VirtualFileEdit Write(string path, string value) => new("write", path, Encoding.UTF8.GetBytes(value).ToImmutableArray());
    private static string Text(VirtualFileSet files, string path) => Encoding.UTF8.GetString(files.Entries[path].Data.AsSpan());

    [Theory]
    [InlineData("C:/secret")][InlineData("/etc/passwd")][InlineData("file:///etc/passwd")][InlineData("//server/share")]
    [InlineData("//xlsq/../outside")][InlineData("//mem/a/../../x")][InlineData("//temp/a\\b")][InlineData("//xlsq/%2e%2e/x")]
    [InlineData("//mem/a//b")][InlineData("//mem/a:stream")][InlineData("//mem/a.")][InlineData("//temp/a ")]
    [InlineData("//mem/./x")][InlineData("//mem/a\0b")][InlineData("//MEM/a")][InlineData("//temp/")]
    public void HostAndWireRejectAmbiguousOrEscapingPaths(string path)
    {
        Assert.Throws<ArgumentException>(() => VirtualPath.Validate(path));
        Assert.Throws<ProtocolException>(() => VirtualFileProtocol.Path(path));
    }

    [Fact]
    public void ParentAliasesTypesAndBudgetsAreValidatedBeforeMutation()
    {
        var files = VirtualFileSet.Empty.Apply([Write("//mem/A/x", "one")]);
        Assert.Throws<ArgumentException>(() => files.Apply([Write("//mem/a/y", "two")]));
        Assert.Throws<ArgumentException>(() => files.Apply([Write("//mem/A", "file")]));
        Assert.Throws<ArgumentException>(() => files.Apply([new("remove", "//mem/A")]));
        Assert.Throws<ArgumentException>(() => files.Apply([new("write", "//mem/large", new byte[VirtualFileSet.MaxBytes].ToImmutableArray())]));
        Assert.Equal("one", Text(files, "//mem/A/x"));
    }

    [Fact]
    public void WorkerFilesAreTransactionalAndOnlyPackageVolumePersists()
    {
        using var doc = Document("""
            quickExcel.actions.register('write', ({fs,snapshot,patch})=>{
              fs.mkdir('//xlsq/data'); fs.mkdir('//xlsq/empty');
              fs.writeText('//xlsq/data/中文.txt','汉字 😀\nline');
              fs.writeJson('//temp/session.json',{count:7});
              fs.writeBytes('//mem/cache',new Uint8Array([0,128,255]));
              if(fs.readText('//xlsq/data/中文.txt')!=='汉字 😀\nline') throw Error('UTF8');
              if(fs.readJson('//temp/session.json').count!==7 || fs.readBytes('//mem/cache')[2]!==255) throw Error('read');
              if(!fs.list('//xlsq/').some(x=>x.path==='//xlsq/data' && x.type==='directory')) throw Error('list');
              if(!fs.list('//xlsq/data',{recursive:true}).some(x=>x.size>0)) throw Error('recursive');
              if(!fs.readJson('//xlsq/manifest.json').scripts.modules.main) throw Error('system');
              fs.copy('//mem/cache','//temp/copy'); fs.move('//temp/copy','//temp/moved'); fs.remove('//temp/moved');
              fs.move('//mem/cache','//mem/cache'); if(!fs.exists('//mem/cache')) throw Error('self move');
              patch.setValue(snapshot.sheets[0].id,'A1','committed');
            });
            quickExcel.actions.register('next',({fs,patch,snapshot})=>{
              patch.setValue(snapshot.sheets[0].id,'A2',fs.readJson('//temp/session.json').count);
            });
            """);
        doc.EnableScriptsAndCalculate();
        doc.RunAction("write");
        Assert.Equal("committed", doc.GetCellValue(0, 0, 0));
        Assert.Equal("汉字 😀\nline", Text(doc.Snapshot.PackageFiles, "//xlsq/data/中文.txt"));
        Assert.True(doc.Snapshot.PackageFiles.Entries["//xlsq/empty"].IsDirectory);
        Assert.Equal(2, doc.Session.TransientFiles.Entries.Count);
        var saved = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "event-files", Guid.NewGuid() + ".xlsq");
        Directory.CreateDirectory(Path.GetDirectoryName(saved)!);
        doc.Save(saved);
        using (var zip = ZipFile.OpenRead(saved))
        {
            Assert.Contains(zip.Entries, e => e.FullName == "data/中文.txt");
            Assert.Contains(zip.Entries, e => e.FullName == "empty/");
            Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("session.json") || e.FullName.Contains("cache"));
        }
        Assert.True(doc.Snapshot.PackageFiles.ContentEquals(Storage.Open(saved).PackageFiles));
        doc.RunAction("next"); Assert.Equal("7", doc.GetCellValue(0, 1, 0));
        Assert.True(doc.Undo()); Assert.True(doc.Undo());
        Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Equal("", doc.GetCellValue(0, 0, 0));
        // Ephemeral exchange/cache files are intentionally outside workbook undo history.
        Assert.Equal(2, doc.Session.TransientFiles.Entries.Count);
        Assert.True(doc.Redo()); Assert.Equal("committed", doc.GetCellValue(0, 0, 0));
        using var reopened = WorkbookDocument.Open(saved, new(Storage, new WorkbookCalculator(new ScriptWorkerRuntimeFactory())));
        Assert.Empty(reopened.Session.TransientFiles.Entries);
        doc.Dispose(); Assert.Throws<ObjectDisposedException>(() => doc.Session.TransientFiles);
    }

    [Theory]
    [InlineData("throw Error('abort')")]
    [InlineData("try { fs.writeText('//temp/../escape','x'); } catch {}")]
    [InlineData("fs.writeJson('//xlsq/manifest.json',{})")]
    [InlineData("fs.writeBytes('//mem/big',new Uint8Array(8388608))")]
    [InlineData("return Promise.resolve()")]
    public void FailedWorkerNeverCommitsAnyVolumeOrCell(string tail)
    {
        using var doc = Document("quickExcel.actions.register('A',({fs,patch,snapshot})=>{fs.writeText('//mem/a','memory');fs.writeText('//temp/a','temp');fs.writeText('//xlsq/a','package');patch.setValue(snapshot.sheets[0].id,'A1',42);" + tail + ";});");
        doc.EnableScriptsAndCalculate(); var before = doc.Snapshot;
        Assert.NotNull(Record.Exception(() => doc.RunAction("A")));
        Assert.Equal(before.Revision, doc.Snapshot.Revision); Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Empty(doc.Snapshot.Sheets[0].Cells); Assert.Empty(doc.Session.TransientFiles.Entries); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void MemoryOnlyTransactionsDoNotDirtyOrEnterHistoryAndRejectStaleReplies()
    {
        using var doc = Document("quickExcel.actions.register('A',({fs})=>fs.writeText('//mem/cache','cached'));");
        doc.EnableScriptsAndCalculate(); var revision = doc.Session.Revision;
        Assert.False(doc.RunAction("A").Changed); Assert.Equal(revision, doc.Session.Revision); Assert.False(doc.IsDirty); Assert.False(doc.CanUndo);
        var request = doc.Session.BeginAction("A");
        var result = new ScriptActionResult(request.OperationId, request.Stamp.InputRevision, "A", [], []) { FileEdits = [Write("//mem/cache", "late")] };
        doc.SetCell(0, 0, 0, "new");
        Assert.NotNull(Record.Exception(() => doc.Session.CommitAction(request, result)));
        Assert.Equal("cached", Text(doc.Session.TransientFiles, "//mem/cache"));
    }

    [Fact]
    public void SystemSourceEditsRefreshHashesRevokeTrustAndPersist()
    {
        using var doc = Document("quickExcel.actions.register('A',({fs})=>{fs.writeText('//xlsq/scripts/main.js',\"quickExcel.actions.register('B',()=>{});\");});");
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.ScriptModules.Revision;
        doc.RunAction("A"); Assert.False(doc.ScriptsEnabled);
        Assert.Contains("'B'", doc.ScriptSource); Assert.True(doc.Snapshot.ScriptModules.Revision > revision);
        Assert.Null(doc.Snapshot.ScriptModules.Signature); Assert.False(doc.CanUndo);
        var files = Storage.Capture(doc.Snapshot);
        Assert.Equal(doc.ScriptSource, Text(files, "//xlsq/scripts/main.js"));
    }

    [Fact]
    public void RawWorkbookFileReplacementIsParsedAndUndoable()
    {
        using var doc = Document("quickExcel.actions.register('A',()=>{});");
        doc.EnableScriptsAndCalculate();
        var source = Storage.Capture(doc.Snapshot);
        using var input = new MemoryStream(source.Entries["//xlsq/workbook.xlsx"].Data.ToArray());
        using var output = new MemoryStream();
        using (var original = new ZipArchive(input, ZipArchiveMode.Read))
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
        {
            foreach (var entry in original.Entries)
            {
                using var stream = entry.Open(); using var data = new MemoryStream(); stream.CopyTo(data);
                var bytes = data.ToArray();
                if (entry.FullName == "xl/worksheets/sheet1.xml")
                {
                    var xml = System.Xml.Linq.XDocument.Parse(Encoding.UTF8.GetString(bytes)); var ns = xml.Root!.Name.Namespace;
                    xml.Descendants(ns + "sheetData").Single().Add(new System.Xml.Linq.XElement(ns + "row", new System.Xml.Linq.XAttribute("r", 1),
                        new System.Xml.Linq.XElement(ns + "c", new System.Xml.Linq.XAttribute("r", "A1"), new System.Xml.Linq.XElement(ns + "v", 17))));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString());
                }
                using var destination = zip.CreateEntry(entry.FullName).Open(); destination.Write(bytes);
            }
        }
        var request = doc.Session.BeginAction("A");
        doc.Session.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, "A", [], [])
        { FileEdits = [new("write", "//xlsq/workbook.xlsx", output.ToArray().ToImmutableArray())] });
        Assert.Equal("17", doc.GetCellValue(0, 0, 0)); Assert.True(doc.ScriptsEnabled);
        Assert.True(doc.Undo()); Assert.Equal("", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void EventsAreTrustedOrderedSingleTransactionAndDoNotReenter()
    {
        using var doc = Document("""
            function log(ctx) {
              const entries=ctx.fs.exists('//mem/events')?ctx.fs.readJson('//mem/events'):[];
              entries.push(ctx.event.name); ctx.fs.writeJson('//mem/events',entries);
            }
            for(const name of ['workbook.open','sheet.activate','sheet.deactivate','sheet.added','sheet.deleted','cell.click','cell.edit','cell.commit','cells.changed']) quickExcel.events.on(name,log);
            quickExcel.events.on('cell.commit',({snapshot,patch,event})=>{
              if(!Object.isFrozen(event) || event.ranges[0]!=='A1') throw Error('event');
              patch.setValue(snapshot.sheets[0].id,'B1','event');
            });
            """);
        doc.Session.RaiseScriptEvent("cell.click", doc.Snapshot.ActiveSheetId, ["A1"]); Assert.False(doc.Session.HasScriptEvents);
        doc.EnableScriptsAndCalculate(); Assert.Empty(doc.ListActions()); doc.DispatchScriptEvents();
        var sheet = doc.AddSheet("Second"); doc.Session.UpdateNavigation(sheet.Id, new(0, 0));
        doc.Session.Execute(new DeleteSheetCommand(sheet.Id));
        foreach (var name in new[] { "cell.click", "cell.edit" }) doc.Session.RaiseScriptEvent(name, doc.Snapshot.ActiveSheetId, ["A1"]);
        doc.SetCell(0, 0, 0, "input"); doc.Session.RaiseScriptEvent("cell.commit", doc.Snapshot.ActiveSheetId, ["A1"]);
        doc.DispatchScriptEvents();
        Assert.Equal("event", doc.GetCellValue(0, 0, 1)); Assert.False(doc.Session.HasScriptEvents);
        var events = JsonNode.Parse(Text(doc.Session.TransientFiles, "//mem/events"))!.AsArray().Select(n => n!.GetValue<string>()).ToArray();
        Assert.Equal(new[] { "workbook.open", "sheet.activate", "sheet.added", "sheet.deactivate", "sheet.activate", "sheet.deleted", "sheet.deactivate", "sheet.activate", "cell.click", "cell.edit", "cells.changed", "cell.commit" }, events);
        Assert.True(doc.Undo()); Assert.Equal("", doc.GetCellValue(0, 0, 1));
        doc.DisableScripts(); Assert.False(doc.Session.HasScriptEvents);
    }

    [Theory]
    [InlineData("workbook.saved")][InlineData("workbook.calculated")][InlineData("workbook.close")]
    public void CompletionNotificationsCannotDirtyTheWorkbook(string name)
    {
        using var doc = Document($"quickExcel.events.on('{name}',({{fs}})=>fs.writeText('//xlsq/late','no')); ");
        doc.EnableScriptsAndCalculate(); doc.ListActions();
        while (doc.Session.TakeScriptEvent() is not null) { }
        doc.Session.RaiseScriptEvent(name);
        Assert.Throws<InvalidOperationException>(() => doc.DispatchScriptEvents());
        Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.False(doc.IsDirty);
    }

    [Fact]
    public void MultipleHandlersShareFilesAndRollbackTogether()
    {
        using var doc = Document("""
            quickExcel.events.on('cell.click',({fs})=>fs.writeText('//mem/a','first'));
            quickExcel.events.on('cell.click',({fs})=>{
              if(fs.readText('//mem/a')!=='first') throw Error('missing');
              fs.writeText('//xlsq/a','second'); throw Error('rollback');
            });
            """);
        doc.EnableScriptsAndCalculate(); doc.DispatchScriptEvents();
        doc.Session.RaiseScriptEvent("cell.click", doc.Snapshot.ActiveSheetId, ["A1"]);
        Assert.NotNull(Record.Exception(() => doc.DispatchScriptEvents()));
        Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Empty(doc.Session.TransientFiles.Entries); Assert.False(doc.CanUndo);
    }

    [Fact]
    public async Task CancelledFileWorkerAndHistoryOverflowNeverCommitTransientData()
    {
        using var doc = Document("quickExcel.actions.register('A',({fs})=>{fs.writeText('//mem/a','no');fs.writeText('//xlsq/a','no');while(true){}});");
        doc.EnableScriptsAndCalculate(); using var cancellation = new CancellationTokenSource();
        var run = Task.Run(() => doc.RunAction("A", cancellation.Token));
        cancellation.CancelAfter(150);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await run);
        Assert.Empty(doc.Session.TransientFiles.Entries); Assert.Empty(doc.Snapshot.PackageFiles.Entries);
        using var tiny = new DocumentSession(doc.Snapshot, 1, virtualFiles: Storage);
        tiny.SetScriptsEnabled(true); var request = tiny.BeginAction("A");
        Assert.Throws<InvalidOperationException>(() => tiny.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, "A", [], [])
        { FileEdits = [Write("//mem/a", "memory"), Write("//xlsq/a", "permanent")] }));
        Assert.Empty(tiny.TransientFiles.Entries); Assert.Empty(tiny.Snapshot.PackageFiles.Entries);
    }

    [Fact]
    public void TamperedPersistentFileHashAndNamespaceCollisionsAreRejectedOnOpen()
    {
        using var doc = Document("quickExcel.actions.register('A',({fs})=>{fs.mkdir('//xlsq/data');fs.writeText('//xlsq/data/a','good');});");
        doc.EnableScriptsAndCalculate(); doc.RunAction("A");
        var path = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "event-files", Guid.NewGuid() + ".xlsq");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!); doc.Save(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            zip.GetEntry("data/a")!.Delete(); using var stream = zip.CreateEntry("data/a").Open(); stream.Write(Encoding.UTF8.GetBytes("tampered"));
        }
        Assert.Throws<InvalidDataException>(() => Storage.Open(path));
        Assert.Throws<ArgumentException>(() => Storage.Apply(doc.Snapshot, [Write("//xlsq/scripts", "not a directory")]));
    }

    [Fact]
    public void SourceChangesCannotHideNonUndoableCellOrDataChanges()
    {
        using var doc = Document("quickExcel.actions.register('A',({fs,patch,snapshot})=>{fs.writeText('//xlsq/scripts/main.js','// changed');patch.setValue(snapshot.sheets[0].id,'A1',1);});");
        doc.EnableScriptsAndCalculate(); var source = doc.ScriptSource;
        Assert.Throws<InvalidOperationException>(() => doc.RunAction("A"));
        Assert.Equal(source, doc.ScriptSource); Assert.Empty(doc.Snapshot.Sheets[0].Cells); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void VirtualSystemEditsPreserveV1MigrationOverwriteGuard()
    {
        const string source = "quickExcel.actions.register('A',({fs})=>fs.writeText('//xlsq/scripts/main.js','// replacement'));";
        using var seed = Document(source);
        var path = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "event-files", Guid.NewGuid() + ".xlsq");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        Fixtures.WorkbookPackageFixture.WriteV1(path, Storage.Capture(seed.Snapshot).Entries["//xlsq/workbook.xlsx"].Data.ToArray(), source);
        using var doc = WorkbookDocument.Open(path, new(Storage, new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.Session.Execute(new UpdateScriptModulesCommand(M4ActionTests.Actions(source)));
        doc.EnableScriptsAndCalculate(); doc.RunAction("A");
        Assert.True(doc.Snapshot.FileState!.RequiresMigration);
        Assert.Throws<InvalidOperationException>(() => doc.Save(path));
    }
}
