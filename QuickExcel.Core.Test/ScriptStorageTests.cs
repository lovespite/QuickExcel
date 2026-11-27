using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptStorageTests
{
    private static WorkbookDocument Document(WorkbookSnapshot? snapshot = null) => new(snapshot ?? Book(Sheet("Data")),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    private static void Commit(DocumentSession session, ScriptActionRequest request) =>
        session.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, request.Name!, [], []));
    private static string Artifact(string name)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../script-storage-tests", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(root); return Path.Combine(root, name);
    }

    [Fact]
    public void BothGlobalStoresImplementEveryMethodAndPreserveExactStrings()
    {
        using var doc = Document();
        doc.RunConsole("""
            for (const store of [localStorage, sessionStorage]) {
                if (store.get('missing') !== undefined) throw Error('missing');
                store.set('', ''); store.set('__proto__', '中文😀'); store.set('a', 'first'); store.set('A', '\u0000');
                store.set('a', 'replaced'); store.remove('absent');
                const keys = store.keys(), values = store.values(), entries = store.entries();
                if (JSON.stringify(keys) !== JSON.stringify(['', 'A', '__proto__', 'a'])) throw Error('keys');
                if (JSON.stringify(entries) !== JSON.stringify(keys.map((k,i)=>[k,values[i]]))) throw Error('entries');
                entries[0][1]='changed'; keys.length=0; values[0]='changed';
                if (store.get('') !== '' || store.get('__proto__') !== '中文😀') throw Error('copy');
                store.remove('a'); if (store.get('a') !== undefined) throw Error('remove');
                store.clear(); if (store.keys().length !== 0) throw Error('clear');
                store.set('final','yes');
            }
            if (host.localStorage !== localStorage || host.sessionStorage !== sessionStorage) throw Error('host');
            """);
        Assert.Equal("yes", doc.Snapshot.LocalStorage.Get("final"));
        Assert.Equal("yes", doc.Session.SessionStorage.Get("final"));
        Assert.Single(doc.Snapshot.LocalStorage.Entries);
    }

    [Fact]
    public void EntryQuotaUsesUtf8AndReplacementReclaimsThePreviousSize()
    {
        var value = new string('x', ScriptStorage.MaxEntryBytes - 1);
        var storage = ScriptStorage.Empty.Set("k", value);
        Assert.Equal(ScriptStorage.MaxEntryBytes, storage.Bytes);
        Assert.Throws<ScriptStorageQuotaException>(() => storage.Set("k", value + "x"));
        Assert.Throws<ScriptStorageQuotaException>(() => storage.Set("中", value));
        Assert.Throws<EncoderFallbackException>(() => storage.Set("bad", "\ud800"));
        var replaced = storage.Set("k", "短"); Assert.Equal(4, replaced.Bytes);
        Assert.Equal(0, replaced.Remove("k").Bytes);
        Assert.Equal(value, storage.Get("k"));
    }

    [Fact]
    public void SessionAcceptsExactly100MiBAndLocalHasNoTotalQuota()
    {
        var value = new string('x', ScriptStorage.MaxEntryBytes - 3);
        var storage = ScriptStorage.Empty;
        for (var i = 0; i < 100; i++) storage = storage.Set(i.ToString("D3"), value, ScriptStorage.MaxSessionBytes);
        Assert.Equal(ScriptStorage.MaxSessionBytes, storage.Bytes);
        Assert.Same(storage, storage.Set("000", value, ScriptStorage.MaxSessionBytes));
        Assert.Throws<ScriptStorageQuotaException>(() => storage.Set("overflow", "", ScriptStorage.MaxSessionBytes));
        Assert.Equal(ScriptStorage.MaxSessionBytes, storage.Remove("000").Set("100", value, ScriptStorage.MaxSessionBytes).Bytes);
        var larger = storage.Set("100", value);
        Assert.Equal(101L * ScriptStorage.MaxEntryBytes, larger.Bytes);
        var file = Artifact("over-100-mib.xlsq");
        var io = new TypedWorkbookStorage(); io.Save(file, Book(Sheet("Data")).WithLocalStorage(larger));
        Assert.True(larger.ContentEquals(io.Open(file).LocalStorage));
    }

    [Fact]
    public void LocalCellAndFileWritesUndoTogetherWhileSessionSurvivesUndoAndScriptReset()
    {
        using var doc = Document();
        doc.RunConsole("localStorage.set('saved','one'); sessionStorage.set('live','one'); workbook.getActiveCell().setValue(42); workbook.files.writeText('//xlsq/data','ok');");
        Assert.True(doc.IsDirty); Assert.True(doc.Undo());
        Assert.Empty(doc.Snapshot.LocalStorage.Entries); Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Equal("", doc.GetCellValue(0, 0, 0));
        Assert.Equal("one", doc.Session.SessionStorage.Get("live"));
        Assert.True(doc.Redo()); Assert.Equal("one", doc.Snapshot.LocalStorage.Get("saved"));
        doc.Session.SetScriptsEnabled(false);
        doc.RunConsole("if(sessionStorage.get('live')!=='one'||localStorage.get('saved')!=='one') throw Error('lost');");
    }

    [Fact]
    public void SessionOnlyWritesDoNotDirtyWorkbookAndDoNotLeakAcrossSessionsOrReopen()
    {
        using var doc = Document();
        Assert.False(doc.RunConsole("sessionStorage.set('memory','yes')").Changed);
        Assert.False(doc.IsDirty); Assert.False(doc.CanUndo);
        doc.RunConsole("if(sessionStorage.get('memory')!=='yes')throw Error('lost'); localStorage.set('saved','yes');");
        var file = Artifact("roundtrip.xlsq"); doc.Save(file);
        using var fresh = Document(new TypedWorkbookStorage().Open(file));
        Assert.Empty(fresh.Session.SessionStorage.Entries); Assert.Equal("yes", fresh.Snapshot.LocalStorage.Get("saved"));
        fresh.RunConsole("if(sessionStorage.get('memory')!==undefined||localStorage.get('saved')!=='yes')throw Error('scope');");
        doc.Dispose(); Assert.True(doc.Session.IsClosed);
    }

    [Theory]
    [InlineData("throw Error('stop');")]
    [InlineData("try { sessionStorage.set('large','x'.repeat(1048576)); } catch {}")]
    [InlineData("try { localStorage.set('bad', 42); } catch {}")]
    public void ExceptionsAndCaughtInvalidWritesRollBackAllState(string failure)
    {
        using var doc = Document();
        Assert.ThrowsAny<Exception>(() => doc.RunConsole("localStorage.set('x','1'); sessionStorage.set('x','1'); workbook.getActiveCell().setValue(9); " + failure));
        Assert.Empty(doc.Snapshot.LocalStorage.Entries); Assert.Empty(doc.Session.SessionStorage.Entries);
        Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void HostQuotaAndStaleCommitCannotPublishSessionOrLocalChanges()
    {
        using var session = new DocumentSession(Book(Sheet("Data")));
        var request = session.BeginConsole("1"); var model = request.Model!;
        for (var i = 0; i < 100; i++) model.WriteStorage(true, "set", i.ToString("D3"), new string('x', ScriptStorage.MaxEntryBytes - 3));
        model.WriteStorage(false, "set", "local", "x");
        Assert.Throws<ScriptStorageQuotaException>(() => model.WriteStorage(true, "set", "too much", "x"));
        Assert.Throws<InvalidOperationException>(() => Commit(session, request));
        Assert.Empty(session.SessionStorage.Entries); Assert.Empty(session.Snapshot.LocalStorage.Entries);
        request = session.BeginConsole("2"); request.Model!.WriteStorage(true, "set", "old", "x"); request.Model.WriteStorage(false, "set", "old", "x");
        _ = session.BeginConsole("3");
        Assert.Throws<OperationCanceledException>(() => Commit(session, request));
        Assert.Empty(session.SessionStorage.Entries); Assert.Empty(session.Snapshot.LocalStorage.Entries);
    }

    [Fact]
    public void EnumerationPagesLargerStoresWithinTheExistingRoundBudget()
    {
        var storage = ScriptStorage.Empty;
        for (var i = 0; i < 24; i++) storage = storage.Set(i.ToString("D3"), new string('x', 128 * 1024));
        using var doc = Document(Book(Sheet("Data")).WithLocalStorage(storage));
        doc.RunConsole("const entries=localStorage.entries(); if(entries.length!==24||entries.some(e=>e[1].length!==131072))throw Error('paging');");
        Assert.False(doc.IsDirty);
    }

    [Fact]
    public void ExcessiveEnumerationPreservesTheBridgeBudgetAndRollsBackWrites()
    {
        var storage = ScriptStorage.Empty;
        for (var i = 0; i < 24; i++) storage = storage.Set(i.ToString("D3"), new string('\0', 128 * 1024));
        using var doc = Document(Book(Sheet("Data")).WithLocalStorage(storage));
        Assert.Throws<ScriptRuntimeFailureException>(() => doc.RunConsole("sessionStorage.set('lost','x'); localStorage.set('lost','x'); localStorage.entries();"));
        Assert.Empty(doc.Session.SessionStorage.Entries); Assert.Null(doc.Snapshot.LocalStorage.Get("lost"));
        Assert.False(doc.IsDirty);
    }

    [Fact]
    public void StorageSurvivesStructureStyleAndRawPackageEditsAndCannotBeSilentlySavedAsXlsx()
    {
        using var doc = Document();
        doc.RunConsole("localStorage.set('keep','yes'); workbook.getActiveCell().getFormat().getFont().setBold(true);");
        doc.Session.Execute(new ShiftAxisCommand(doc.Snapshot.ActiveSheetId, SheetAxis.Row, 0, 1));
        Assert.Equal("yes", doc.Snapshot.LocalStorage.Get("keep"));
        doc.RunConsole("const manifest=workbook.files.readJson('//xlsq/manifest.json'); workbook.files.writeJson('//xlsq/manifest.json',manifest);");
        Assert.Equal("yes", doc.Snapshot.LocalStorage.Get("keep"));
        Assert.ThrowsAny<Exception>(() => doc.RunConsole("workbook.files.writeText('//xlsq/state/local-storage/000000.bin','bad');"));
        Assert.Equal("yes", doc.Snapshot.LocalStorage.Get("keep"));
        var emptyScripts = ScriptModuleSet.Single("", 3);
        var snapshot = doc.Snapshot.WithScriptModules(emptyScripts);
        Assert.True(TypedWorkbookStorage.RequiresScriptContainer(snapshot));
        Assert.Throws<InvalidOperationException>(() => new TypedWorkbookStorage().Save(Artifact("blocked.xlsx"), snapshot));
    }

    [Fact]
    public void InitializationActionsEventsAndConsoleShareStoresWhileUdfAccessIsRejected()
    {
        var sources = new[] {
            ("main", "initialization", "export function main(){localStorage.set('n','1');sessionStorage.set('n','10');}"),
            ("action", "actions", "export function main(){localStorage.set('n',String(+localStorage.get('n')+1));sessionStorage.set('n',String(+sessionStorage.get('n')+1));}"),
            ("events", "events", "export function onCellClick(context){sessionStorage.set('n',String(+sessionStorage.get('n')+1));} export function onWorkbookSaved(context){sessionStorage.set('saved','yes');}"),
            ("udf", "formula-functions", "export function READ(){try{localStorage.get('n');return 'bad';}catch(e){return e.code;}}") };
        var modules = new ScriptModuleSet(sources.Select(s => KeyValuePair.Create(s.Item1, new ScriptModule($"scripts/{s.Item1}.js", s.Item3))),
            sources.Select(s => s.Item1), scriptApiVersion: 4, bindings: sources.Select(s => new ScriptBinding(s.Item1, s.Item2)));
        using var doc = Document(Book(Sheet("Data", (0, 0, F("=JS.udf.READ()")))).WithScriptModules(modules));
        doc.EnableScriptsAndCalculate(); doc.RunAction("action");
        doc.Session.RaiseScriptEvent("cell.click", doc.Snapshot.ActiveSheetId, ["A1"]); doc.DispatchScriptEvents();
        Assert.Equal("2", doc.Snapshot.LocalStorage.Get("n")); Assert.Equal("12", doc.Session.SessionStorage.Get("n"));
        Assert.Equal("InvalidReference", doc.GetCellValue(0, 0, 0));
        doc.Session.RaiseScriptEvent("workbook.saved"); doc.DispatchScriptEvents();
        Assert.Equal("yes", doc.Session.SessionStorage.Get("saved"));
        doc.RunConsole("if(sessionStorage.get('n')!=='12'||localStorage.get('n')!=='2')throw Error('shared');");
    }

    [Theory]
    [InlineData("workbook.saved", false)]
    [InlineData("workbook.calculated", false)]
    [InlineData("workbook.close", true)]
    public void ReadOnlyEventsRejectForbiddenStorageWrites(string name, bool sessionStore)
    {
        using var session = new DocumentSession(Book(Sheet("Data")).WithScriptModules(M4ActionTests.Actions("function main(){}")));
        session.SetScriptsEnabled(true);
        var request = session.BeginAction(ScriptEventNames.Prefix + name, new(name, session.SessionId, 1, session.Revision, null, null, [], false));
        request.Model!.WriteStorage(sessionStore, "set", "x", "1");
        Assert.Throws<InvalidOperationException>(() => Commit(session, request));
        Assert.Empty(session.SessionStorage.Entries); Assert.Empty(session.Snapshot.LocalStorage.Entries);
    }

    [Fact]
    public async Task TypeScriptChecksTheCompleteStorageContract()
    {
        const string source = """
            export function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host): void {
                const store: IStorage = localStorage;
                const other: QuickExcel.IStorage = host.sessionStorage;
                store.set('a', 'b'); const value: string | undefined = store.get('a');
                const entries: Array<[string,string]> = store.entries();
                const keys: string[] = store.keys(); const values: string[] = store.values();
                other.set(keys[0], values[0]); sessionStorage.remove('a'); host.localStorage.clear();
            }
            """;
        var modules = new ScriptModuleSet([KeyValuePair.Create("main", new ScriptModule("scripts/main.ts", source, ScriptLanguage.TypeScript))],
            ["main"], scriptApiVersion: 4, bindings: [new("main", "actions")]);
        await ScriptCompilation.PrepareAsync(modules, new(Guid.NewGuid(), 0, 0, 0), default);
    }

    [Fact]
    public void VersionThreeCanUseStorageWithoutChangingFileApiVersion()
    {
        var modules = M4ActionTests.Actions("function main(){host.localStorage.set('v','3');sessionStorage.set('v','3');}");
        using var doc = Document(Book(Sheet("Data")).WithScriptModules(modules));
        doc.EnableScriptsAndCalculate(); doc.RunAction("A");
        Assert.Equal("3", doc.Snapshot.LocalStorage.Get("v"));
        Assert.Equal("3", doc.Session.SessionStorage.Get("v"));
        var file = Artifact("version-three.xlsq"); doc.Save(file);
        var saved = new TypedWorkbookStorage().Open(file); Assert.Equal(3, saved.ScriptModules.ScriptApiVersion);
        Assert.Equal("3", saved.LocalStorage.Get("v"));
    }

    [Fact]
    public void ReadOnlyFormulaGuardRejectsStorageWritesAndDirectSaveBypasses()
    {
        var original = Book(Sheet("Data")).WithLocalStorage(ScriptStorage.Empty.Set("kept", "yes"));
        var state = new WorkbookFileState("artifacts/readonly-storage.xlsq", 2, [], original, [], [],
            preservedFormulas: [new(PreservedFormulaKind.Array, "xl/worksheets/sheet1.xml", "Data", "A1", "A1:A2", "SUM(A1)")]);
        using var session = new DocumentSession(original.WithFileState(state));
        foreach (var transient in new[] { false, true })
        {
            var request = session.BeginConsole("1");
            Assert.Equal("yes", request.Model!.ReadStorage(false).Get("kept"));
            Assert.Throws<InvalidOperationException>(() => request.Model.WriteStorage(transient, "set", "bad", "no"));
            Assert.Throws<InvalidOperationException>(() => Commit(session, request));
        }
        Assert.Throws<InvalidOperationException>(() => new TypedWorkbookStorage().Save(Artifact("readonly.xlsq"), session.Snapshot.WithLocalStorage(ScriptStorage.Empty)));
        Assert.Equal("yes", session.Snapshot.LocalStorage.Get("kept")); Assert.Empty(session.SessionStorage.Entries);
    }

    [Theory]
    [InlineData("hash")]
    [InlineData("duplicate")]
    [InlineData("length")]
    [InlineData("feature")]
    public void CorruptStoragePackagesAreRejected(string fault)
    {
        var file = Artifact("malformed.xlsq"); var io = new TypedWorkbookStorage();
        io.Save(file, Book(Sheet("Data")).WithLocalStorage(ScriptStorage.Empty.Set("a", "b")));
        Dictionary<string, byte[]> parts;
        using (var zip = ZipFile.OpenRead(file)) parts = zip.Entries.ToDictionary(e => e.FullName, e => { using var stream = e.Open(); using var data = new MemoryStream(); stream.CopyTo(data); return data.ToArray(); });
        var manifest = JsonNode.Parse(parts["manifest.json"])!;
        var reference = manifest["localStorage"]![0]!; var path = reference["path"]!.GetValue<string>();
        if (fault == "feature") manifest["requiredFeatures"]!.AsArray().Remove(manifest["requiredFeatures"]!.AsArray().Single(n => n!.GetValue<string>() == "script-storage"));
        else
        {
            using var data = new MemoryStream(); using (var writer = new BinaryWriter(data, Encoding.UTF8, true))
            {
                writer.Write(1); writer.Write(fault == "duplicate" ? 2 : 1);
                for (var i = 0; i < (fault == "duplicate" ? 2 : 1); i++) { writer.Write(fault == "length" ? 1048577 : 1); writer.Write((byte)'a'); writer.Write(1); writer.Write((byte)'c'); }
            }
            parts[path] = data.ToArray();
            if (fault != "hash") reference["sha256"] = Convert.ToHexStringLower(SHA256.HashData(parts[path]));
        }
        parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString());
        using (var zip = new ZipArchive(File.Create(file), ZipArchiveMode.Create))
            foreach (var pair in parts) { using var stream = zip.CreateEntry(pair.Key).Open(); stream.Write(pair.Value); }
        Assert.Throws<InvalidDataException>(() => io.Open(file));
    }
}
