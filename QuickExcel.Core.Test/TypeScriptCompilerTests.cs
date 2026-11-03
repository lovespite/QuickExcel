using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Calculation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TypeScriptCompilerTests
{
    [Fact]
    public async Task DenseOneMebibyteSourceCompilesWithinTheFrozenIsolatedBudget()
    {
        var source = new System.Text.StringBuilder();
        for (var i=0;;i++)
        {
            var line=$"export const v{i}:number={i};\n";
            if(source.Length+line.Length>1024*1024)break;
            source.Append(line);
        }
        var request = Request(new CompilerModule("main", "typescript", source.ToString()));
        var result = await new ScriptCompiler().CompileAsync(request, Stamp);
        Assert.True(result.Result.Success);
        Assert.Single(result.Result.Outputs);
        CompilerProtocol.ValidateReply(result, request);
    }

    [Fact]
    public void TypeScriptActionsAndEventsUseValidatedSourcesAndAtomicFileCellPatches()
    {
        const string source = """
            enum Amount { Value = 42 }
            quickExcel.actions.register('write', ({snapshot,patch,fs})=>{
              const n:number=Amount.Value;
              patch.setValue(snapshot.sheets[0].id,'A1',n);
              fs.writeText('//xlsq/value.txt',String(n));
            });
            quickExcel.events.on('cell.click',({snapshot,patch,fs})=>{
              const message:string=fs.readText('//xlsq/value.txt');
              patch.setValue(snapshot.sheets[0].id,'B1',message);
              fs.writeText('//xlsq/event.txt',message);
            });
            """;
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        { ["main"] = new("scripts/main.ts", source, ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);
        using var doc = new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate();
        doc.RunAction("write");
        Assert.Equal("42", doc.GetCellValue(0, 0, 0));
        Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/value.txt"));
        Assert.True(doc.Undo());
        Assert.Empty(doc.Snapshot.PackageFiles.Entries);
        Assert.Equal("", doc.GetCellValue(0, 0, 0));
        Assert.True(doc.Redo());
        doc.Session.RaiseScriptEvent("cell.click", doc.Snapshot.ActiveSheetId, ["A1"]);
        doc.DispatchScriptEvents();
        Assert.Equal("42", doc.GetCellValue(0, 0, 1));
        Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/event.txt"));
        Assert.True(doc.Undo());
        Assert.False(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/event.txt"));
        Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/value.txt"));
        Assert.Equal("", doc.GetCellValue(0, 0, 1));
    }

    [Theory]
    [InlineData("execute-hello")]
    [InlineData("compile-role")]
    [InlineData("compiler-identity")]
    [InlineData("revision")]
    [InlineData("execute-message")]
    public async Task CompilerRoleIsLockedBeforeSourceAndRejectsRoundChanges(string fault)
    {
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        using var process = WorkerProcess.Start(Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"), deadline.Token, compilation: true);
        var identity = M4ProtocolTests.Identity();
        var hello = fault switch
        {
            "execute-hello" => new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [1, 2], WorkerProtocol.MaxFrameBytes),
            "compile-role" => CompilerProtocol.Hello with { Role = "execute" },
            "compiler-identity" => CompilerProtocol.Hello with { Compiler = new string('0', 64) },
            _ => CompilerProtocol.Hello
        };
        await WorkerProtocol.WriteAsync(process.Input, WorkerProtocol.Message(identity, "hello", hello), deadline.Token);
        if (fault is "revision" or "execute-message")
        {
            var accepted = await WorkerProtocol.ReadAsync(process.Output, deadline.Token);
            WorkerProtocol.ValidateReply(accepted, identity, "hello");
            identity = identity with { RequestId = Guid.NewGuid(), Revision = fault == "revision" ? identity.Revision + 1 : identity.Revision };
            var request = Request(new CompilerModule("main", "typescript", "throw new Error('never execute');"));
            await WorkerProtocol.WriteAsync(process.Input, WorkerProtocol.Message(identity, fault == "execute-message" ? "load" : "compile", request), deadline.Token);
        }
        await Assert.ThrowsAsync<EndOfStreamException>(() => WorkerProtocol.ReadAsync(process.Output, deadline.Token));
    }

    private static CompileRequest Request(params CompilerModule[] modules)
    {
        var request = new CompileRequest(modules, ["main"], "");
        return request with { Sha256 = request.ComputeDigest() };
    }
    private static CalculationStamp Stamp => new(Guid.NewGuid(), 1, 1, 1);

    [Fact]
    public async Task CrossModuleTypesAreCheckedBeforeAnySourceIsExecuted()
    {
        var compiler = new ScriptCompiler();
        var request = Request(new("shared", "typescript", "export const twice=(n:number):number=>n*2;"),
            new("main", "typescript", "import {twice} from 'shared'; quickExcel.functions.register('F',()=>twice(4));"));
        var result = await compiler.CompileAsync(request, Stamp);
        Assert.True(result.Result.Success);
        var modules = new ScriptModuleSet(result.Result.Outputs.ToDictionary(m => m.Id, m => new ScriptModule("scripts/" + m.Id + ".js", m.Javascript)), ["main"], scriptApiVersion: 2);
        using var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, Stamp, TimeSpan.FromSeconds(5), default));
        Assert.Equal(8d, runtime.Invoke("F", []));
        var hit = await compiler.CompileAsync(request, Stamp);
        Assert.Equal(result.OutputSha256, hit.OutputSha256);
        Assert.Equal(1, compiler.ProcessLaunchCount);
        var noRun = await compiler.CompileAsync(Request(new CompilerModule("main", "typescript", "throw new Error('MUST NOT RUN'); export const n:number=2;")), Stamp);
        Assert.True(noRun.Result.Success);
    }

    [Theory]
    [InlineData("export const n:number='bad';")]
    [InlineData("export const n = document.title;")]
    [InlineData("export const n = process.env.PATH;")]
    [InlineData("import {n} from './external'; export {n};")]
    [InlineData("export const n = import('main');")]
    [InlineData("export const n = await Promise.resolve(1);")]
    public async Task UnsupportedOrInvalidSourceReturnsDiagnosticsAndNoArtifacts(string source)
    {
        var result = await new ScriptCompiler().CompileAsync(Request(new CompilerModule("main", "typescript", source)), Stamp);
        Assert.False(result.Result.Success);
        Assert.NotEmpty(result.Result.Diagnostics);
        Assert.Empty(result.Result.Outputs);
    }

    [Fact]
    public async Task CancellationAndSourceIdentityDoNotPopulateCache()
    {
        var request = Request(new CompilerModule("main", "typescript", "export const n:number=1;"));
        var compiler = new ScriptCompiler();
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => compiler.CompileAsync(request, Stamp, cancel.Token));
        Assert.Equal(0, compiler.ProcessLaunchCount);
        await Assert.ThrowsAsync<ProtocolException>(() => compiler.CompileAsync(request with { Modules = [new("main", "typescript", "export const n=2;")] }, Stamp));
        Assert.Equal(0, compiler.ProcessLaunchCount);
    }

    [Fact]
    public async Task InFlightCancellationKillsCompilerAndRejectsArtifacts()
    {
        var compiler = new ScriptCompiler();
        var request = Request(new CompilerModule("main", "typescript", "export const n:number=1;"));
        using var cancel = new CancellationTokenSource();
        var work = compiler.CompileAsync(request, Stamp, cancel.Token);
        while (compiler.ProcessLaunchCount == 0 && !work.IsCompleted) await Task.Delay(5);
        cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work);
        Assert.True((await compiler.CompileAsync(request, Stamp)).Result.Success);
        Assert.Equal(2, compiler.ProcessLaunchCount);
    }

    [Fact]
    public void RuntimeAcceptsOnlyCurrentCompiledSourcesAndMapsErrorsToOriginalLines()
    {
        const string source = "enum E { A, B }\n\nquickExcel.functions.register('F',()=>{\n  const n:number=E.B;\n  throw new Error('mapped '+n);\n});";
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        { ["main"] = new("scripts/main.ts", source, ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 2);
        using var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, Stamp, TimeSpan.FromSeconds(5), default));
        var error = Assert.Throws<InvalidOperationException>(() => runtime.Invoke("F", []));
        Assert.Contains("main:5:", error.Message);
        var bad = modules.WithPrimarySource("export const n:number='wrong';");
        Assert.Throws<ScriptCompilationException>(() => new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(bad, Stamp, TimeSpan.FromSeconds(5), default)));
    }

    [Theory]
    [InlineData("file")]
    [InlineData("source")]
    [InlineData("vlq")]
    [InlineData("hash")]
    public async Task TamperedArtifactsCannotBecomeExecutable(string fault)
    {
        var request = Request(new CompilerModule("main", "typescript", "export const n:number=1;"));
        var reply = await new ScriptCompiler().CompileAsync(request, Stamp);
        var output = reply.Result.Outputs[0];
        var map = System.Text.Json.Nodes.JsonNode.Parse(output.SourceMap)!;
        if (fault == "file") map["file"] = "file:///outside.js";
        if (fault == "source") map["sources"]![0] = "https://outside";
        if (fault == "vlq") map["mappings"] = "!";
        output = output with { SourceMap = map.ToJsonString() };
        reply = reply with { Result = reply.Result with { Outputs = [output] } };
        reply = reply with { OutputSha256 = fault == "hash" ? new string('0', 64) : CompilerProtocol.OutputDigest(reply.Result) };
        Assert.Throws<ProtocolException>(() => CompilerProtocol.ValidateReply(reply, request));
    }
}
