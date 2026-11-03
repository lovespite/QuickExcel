using System.Diagnostics;
using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4ModuleTests
{
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void SyntaxErrorsRetainLogicalModuleAndLine(bool worker)
    {
        var package = Package("import 'broken';", ("broken", "// first line\nexport const value = ;"));
        var error = Assert.ThrowsAny<Exception>(() => Create(package, worker));
        Assert.Contains("SyntaxError", error.Message);
        Assert.Contains("broken:2", error.Message);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void ErrorStackGetterIsNotExecutedByDiagnosticFormatting(bool worker)
    {
        var package = Package("const error = new Error('expected'); Object.defineProperty(error, 'stack', {get(){while(true){}}}); throw error;");
        var error = Assert.ThrowsAny<Exception>(() => Create(package, worker));
        Assert.Contains("expected", error.Message);
        Assert.DoesNotContain("time limit", error.Message);
    }

    private static ScriptModuleSet Package(string main, params (string Id, string Source)[] helpers) => new(
        new[] { new KeyValuePair<string, ScriptModule>("main", new("scripts/main.js", main)) }
            .Concat(helpers.Select(h => new KeyValuePair<string, ScriptModule>(h.Id, new($"scripts/lib/{h.Id}.js", h.Source)))),
        ["main"], scriptApiVersion: 2);
    private static IScriptRuntime Create(ScriptModuleSet modules, bool worker = true, CancellationToken token = default) =>
        (worker ? (IScriptRuntimeFactory)new ScriptWorkerRuntimeFactory() : new QuickJsRuntimeFactory())
            .Create(new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 1, 2, 3), TimeSpan.FromSeconds(5), token));

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void NamedDefaultNamespaceAndReexportsUsePackageIds(bool worker)
    {
        var modules = Package("""
            import discount, { label } from 'exports';
            import * as math from 'price';
            const privateValue = 7;
            quickExcel.functions.register('VALUE', n => discount(n) + math.extra);
            quickExcel.functions.register('TEXT', () => label);
            quickExcel.functions.register('SCOPE', () => this === undefined && !('privateValue' in globalThis));
            """,
            ("exports", "export {default, label} from 'price';"),
            ("price", "export default n => n * .8; export const extra = 2; export const label = '中文 🧮';"),
            ("unused", "throw new Error('unreachable module must not run');"));
        using var runtime = Create(modules, worker);
        Assert.Equal(82d, runtime.Invoke("VALUE", [100d]));
        Assert.Equal("中文 🧮", runtime.Invoke("TEXT", []));
        Assert.Equal(true, runtime.Invoke("SCOPE", []));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void EntryOrderSharedDependenciesAndCyclesFollowEsModuleSemantics(bool worker)
    {
        var package = Package("import { next } from 'shared'; console.log('main'); quickExcel.functions.register('NEXT', next);",
            ("second", "import { next } from 'shared'; console.log('second:' + next());"),
            ("shared", "import { odd } from 'cycle'; console.log('shared'); let n=0; export function next(){return ++n}; export function even(n){return n===0 || odd(n-1)}; quickExcel.functions.register('EVEN',even);"),
            ("cycle", "import { even } from 'shared'; export function odd(n){return n!==0 && even(n-1)};"));
        package = new(package.Modules, ["second", "main", "shared"], scriptApiVersion: 2);
        for (var round = 0; round < 2; round++)
        {
            using var runtime = Create(package, worker);
            Assert.Equal(new[] { "shared", "second:1", "main" }, runtime.GetLogs());
            Assert.Equal(2d, runtime.Invoke("NEXT", []));
            Assert.Equal(true, runtime.Invoke("EVEN", [8d]));
        }
    }

    [Theory]
    [InlineData("missing")] [InlineData("./price")] [InlineData("../price")]
    [InlineData("scripts/lib/price.js")] [InlineData("file:///C:/secret.js")]
    [InlineData("https://example.invalid/code.js")] [InlineData("std")] [InlineData("os")]
    [InlineData("Price")] [InlineData("C:/price.js")]
    public void UndeclaredPathsUrlsAndCaseAliasesCannotResolve(string id)
    {
        var modules = Package($"import '{id}'; quickExcel.functions.register('VALUE',()=>1);", ("price", "export const value = 1;"));
        Assert.Equal("Load", Assert.Throws<ScriptRuntimeFailureException>(() => Create(modules)).Code);
    }

    [Theory]
    [InlineData("this is not valid JavaScript !!!")]
    [InlineData("throw new Error('failure');")]
    [InlineData("quickExcel.functions.register('MODULEVALUE',()=>2);")]
    [InlineData("while(true){}")]
    [InlineData("await Promise.resolve(3);")]
    [InlineData("await new Promise(()=>{});")]
    [InlineData("import('main');")]
    [InlineData("Promise.resolve().then(()=>console.log('late'));")]
    [InlineData("import value from 'main';")]
    [InlineData("import 'main' with { type: 'json' };")]
    public void FailedEntryDiscardsWholeRegistryAndCachedCalculation(string badEntry)
    {
        var package = Package("quickExcel.functions.register('MODULEVALUE',()=>7);", ("bad", badEntry));
        package = new(package.Modules, ["main", "bad"], scriptApiVersion: 2);
        var sheet = Sheet("S", (0, 0, F("=MODULEVALUE()", 42)), (0, 1, F("=3", 9)));
        using var document = new WorkbookDocument(Book(sheet).WithScriptModules(package),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory())));
        Assert.Throws<ScriptRuntimeFailureException>(document.EnableScriptsAndCalculate);
        Assert.False(document.ScriptsEnabled);
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.Equal("9", document.GetCellValue(0, 0, 1));
        document.Session.Execute(new UpdateScriptModulesCommand(Package("quickExcel.functions.register('MODULEVALUE',()=>8);")));
        document.EnableScriptsAndCalculate();
        Assert.Equal("8", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void ModuleFunctionsRetainTypedRangesAndFrozenRegistration()
    {
        var modules = Package("import { shape } from 'helpers'; quickExcel.functions.register({name:'JS.SHAPE',apiVersion:2,parameters:[{name:'r',type:'range'}],returns:'text'},shape);",
            ("helpers", "export const shape = r => [r.rows,r.columns,r.values[1]===null,r.values[2]==='',Object.isFrozen(r.values)].join('|'); quickExcel.functions.register('LATE',()=>quickExcel.functions.register('NEW',()=>3));"));
        var data = Sheet("Data", (0, 0, N(2)), (1, 0, T("")), (1, 1, N(4)));
        var output = Sheet("Out", (0, 0, F("=JS.SHAPE(Data!A1:B2)")), (1, 0, F("=LATE()")));
        var result = new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(Request(Book(data, output).WithScriptModules(modules), true));
        Assert.Equal(CellValue.FromText("2|2|true|true|true"), C(result, output).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Script), C(result, output, 1).Value);
    }

    [Fact]
    public void RoundtripKeepsModuleModeEntryOrderAndNeverRestoresAuthorization()
    {
        var path = Path.Combine(Path.GetTempPath(), "qx-modules-" + Guid.NewGuid() + ".xlsq");
        var modules = Package("import {next} from 'state'; quickExcel.functions.register('NEXT',next);", ("state", "let n=0; export const next=()=>++n;"));
        var sheet = Sheet("S", (0, 0, F("=NEXT()", 10)), (1, 0, F("=NEXT()+A1", 20)));
        var services = new WorkbookServices(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()));
        try
        {
            using var document = new WorkbookDocument(Book(sheet).WithScriptModules(modules), services);
            document.Save(path);
            using var reopened = WorkbookDocument.Open(path, services);
            Assert.False(reopened.ScriptsEnabled);
            Assert.Equal(2, reopened.Snapshot.ScriptModules.ScriptApiVersion);
            Assert.Equal(modules.Modules, reopened.Snapshot.ScriptModules.Modules);
            Assert.Equal(modules.EntryModules.ToArray(), reopened.Snapshot.ScriptModules.EntryModules.ToArray());
            for (var round = 0; round < 2; round++)
            {
                reopened.EnableScriptsAndCalculate();
                Assert.Equal("1", reopened.GetCellValue(0, 0, 0));
                Assert.Equal("3", reopened.GetCellValue(0, 1, 0));
            }
            reopened.ScriptSource += "\n// changed";
            Assert.False(reopened.ScriptsEnabled);
            Assert.Equal(2, reopened.Snapshot.ScriptModules.ScriptApiVersion);
            reopened.Save(path);
            Assert.Equal(modules.Modules["state"], services.Storage.Open(path).ScriptModules.Modules["state"]);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void HelperAndBindingChangesInvalidateAuthorizationAndInFlightResult()
    {
        var original = Package("import 'helper';", ("helper", "quickExcel.functions.register('VALUE',()=>1);"));
        using var session = new DocumentSession(Book(Sheet("S")).WithScriptModules(original));
        session.SetScriptsEnabled(true);
        var old = session.BeginCalculation();
        var changed = Package("import 'helper';", ("helper", "quickExcel.functions.register('VALUE',()=>2);"));
        session.Execute(new UpdateScriptModulesCommand(changed));
        Assert.True(old.CancellationToken.IsCancellationRequested);
        Assert.False(session.ScriptsEnabled);
        Assert.False(session.TryCommitCalculation(new(old.Stamp, [])));
        Assert.Equal(1, session.Snapshot.ScriptModules.Revision);
        Assert.Equal(0, session.Snapshot.ScriptModules.BindingRevision);
        Assert.False(session.CanUndo);
        session.SetScriptsEnabled(true);
        var binding = session.BeginCalculation();
        session.Execute(new UpdateScriptModulesCommand(new(changed.Modules, ["helper", "main"], scriptApiVersion: 2)));
        Assert.True(binding.CancellationToken.IsCancellationRequested);
        Assert.False(session.ScriptsEnabled);
        Assert.Equal(1, session.Snapshot.ScriptModules.BindingRevision);
    }

    [Fact]
    public async Task CancellationInterruptsModuleInitialization()
    {
        using var cancel = new CancellationTokenSource();
        cancel.CancelAfter(150);
        var timer = Stopwatch.StartNew();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Task.Run(() => Create(Package("import 'loop';", ("loop", "while(true){}")), token: cancel.Token)));
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(3));
    }

    [Fact]
    public void ModuleDigestBindsSourcesPathsAndEntryOrderButNotMapOrder()
    {
        var bundle = new LoadModules([new("a", "scripts/a.js", "//中文"), new("b", "scripts/b.js", "//b")], ["a", "b"], "");
        var hash = bundle.ComputeDigest();
        Assert.Equal(hash, (bundle with { Modules = bundle.Modules.Reverse().ToArray() }).ComputeDigest());
        Assert.NotEqual(hash, (bundle with { EntryModules = ["b", "a"] }).ComputeDigest());
        Assert.NotEqual(hash, (bundle with { Modules = [bundle.Modules[0] with { Source = "// changed" }, bundle.Modules[1]] }).ComputeDigest());
        Assert.NotEqual(hash, (bundle with { Modules = [bundle.Modules[0] with { Path = "scripts/other.js" }, bundle.Modules[1]] }).ComputeDigest());
        Assert.Throws<ProtocolException>(() => bundle.Validate());
        (bundle with { Sha256 = hash }).Validate();
    }

    [Theory]
    [InlineData("duplicate-id")] [InlineData("duplicate-path")] [InlineData("unknown-entry")]
    [InlineData("duplicate-entry")] [InlineData("traversal")] [InlineData("url")]
    [InlineData("source-budget")] [InlineData("module-count")] [InlineData("null-module")]
    public void ModuleProtocolRejectsInvalidMappingsBeforeExecution(string scenario)
    {
        var item = new ModuleSource("a", "scripts/a.js", "//a");
        var load = new LoadModules([item], ["a"], "");
        load = scenario switch
        {
            "duplicate-id" => load with { Modules = [item, item with { Path = "scripts/b.js" }] },
            "duplicate-path" => load with { Modules = [item, item with { Id = "b", Path = "scripts/A.js" }] },
            "unknown-entry" => load with { EntryModules = ["b"] },
            "duplicate-entry" => load with { EntryModules = ["a", "a"] },
            "traversal" => load with { Modules = [item with { Path = "scripts/../a.js" }] },
            "url" => load with { Modules = [item with { Id = "https://a" }] },
            "source-budget" => load with { Modules = [item with { Source = new string('中', 350000) }] },
            "module-count" => load with { Modules = Enumerable.Range(0, 65).Select(i => new ModuleSource("m" + i, $"scripts/{i}.js", "")).ToArray() },
            "null-module" => load with { Modules = [null!] },
            _ => throw new Exception()
        };
        Assert.Throws<ProtocolException>(() => load.ComputeDigest());
    }

    [Fact]
    public void MaximumModuleCountAndUtf8BudgetAreAccepted()
    {
        var modules = Enumerable.Range(0, 64).ToDictionary(i => "m" + i, i => new ScriptModule($"scripts/{i}.js", ""));
        var source = "quickExcel.functions.register('VALUE',()=>64); //";
        modules["m0"] = modules["m0"] with { Source = source + new string('x', 1024 * 1024 - Encoding.UTF8.GetByteCount(source)) };
        using var runtime = Create(new ScriptModuleSet(modules, modules.Keys, scriptApiVersion: 2));
        Assert.Equal(64d, runtime.Invoke("VALUE", []));
    }

    [Theory]
    [InlineData("digest")] [InlineData("entry")] [InlineData("path")] [InlineData("legacy-protocol")]
    public async Task RealWorkerRejectsTamperedModuleMessages(string fault)
    {
        using var process = Process.Start(new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"))
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true })!;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        try
        {
            var identity = M4ProtocolTests.Identity();
            var hello = WorkerProtocol.Message(identity, "hello", new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [1, 2], WorkerProtocol.MaxFrameBytes));
            if (fault == "legacy-protocol")
            {
                var bytes = System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(hello with { ProtocolVersion = 1 }, WorkerProtocol.Json);
                var header = new byte[4]; System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(header, bytes.Length);
                await process.StandardInput.BaseStream.WriteAsync(header, timeout.Token);
                await process.StandardInput.BaseStream.WriteAsync(bytes, timeout.Token);
                await process.StandardInput.BaseStream.FlushAsync(timeout.Token);
            }
            else
            {
                await WorkerProtocol.WriteAsync(process.StandardInput.BaseStream, hello, timeout.Token);
                _ = await WorkerProtocol.ReadAsync(process.StandardOutput.BaseStream, timeout.Token);
                var load = new LoadModules([new("main", "scripts/main.js", "while(true){}")], ["main"], "");
                load = load with { Sha256 = load.ComputeDigest() };
                load = fault switch
                {
                    "digest" => load with { Sha256 = new string('0', 64) },
                    "entry" => load with { EntryModules = ["missing"] },
                    "path" => load with { Modules = [load.Modules[0] with { Path = "scripts/../main.js" }] },
                    _ => throw new Exception()
                };
                await WorkerProtocol.WriteAsync(process.StandardInput.BaseStream,
                    WorkerProtocol.Message(identity with { RequestId = Guid.NewGuid() }, "loadModules", load), timeout.Token);
            }
            await process.WaitForExitAsync(timeout.Token);
            Assert.Equal(4, process.ExitCode); // Protocol rejection precedes the nonterminating source.
            Assert.Equal("", await process.StandardOutput.ReadToEndAsync(timeout.Token));
            Assert.Equal("", await process.StandardError.ReadToEndAsync(timeout.Token));
        }
        finally { if (!process.HasExited) { process.Kill(entireProcessTree: true); process.WaitForExit(1000); } }
    }

    [Fact]
    public void RejectedPackageEditPreservesSourcesRevisionAndAuthorization()
    {
        using var session = new DocumentSession(Book(Sheet("S")).WithScriptModules(Package("// original")));
        session.SetScriptsEnabled(true);
        var snapshot = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new UpdateScriptModulesCommand(Package(new string('x', 1024 * 1024 + 1)))));
        Assert.Same(snapshot, session.Snapshot);
        Assert.True(session.ScriptsEnabled);
    }
}
