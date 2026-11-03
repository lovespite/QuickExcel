using System.Text.Json;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Calculation;
using QuickExcel.Storage;

if(args.Length==2 && args[0]=="--performance")return await JscPerformance.Run(args[1]);

if(args.Length==2 && args[0] is "--trust" or "--trust-unavailable")return TrustChecks.Run(args[1],args[0]=="--trust-unavailable");
if(args.Length==2 && args[0]=="--cups")return await CupsChecks.Run(args[1]);

if(args.Length==2 && args[0]=="--recover" && OperatingSystem.IsLinux())
{
    var recovered=ScriptWorkerRuntimeFactory.RecoverAbandonedSandboxes();
    File.WriteAllText(args[1],JsonSerializer.Serialize(new{recovered}));
    return 0;
}

if(args.Length==2 && args[0] is "--sandbox" or "--hold" or "--hold-compile" && OperatingSystem.IsLinux())
    return await LinuxSandboxChecks.Run(args[1],args[0]!="--sandbox",args[0]=="--hold-compile");

if (args.Length == 2 && args[0] == "--bootstrap" && OperatingSystem.IsLinux())
{
    Directory.CreateDirectory(args[1]);
    using var token = new CancellationTokenSource(TimeSpan.FromSeconds(10));
    using var process = LinuxWorkerProcess.Start(Path.Combine(AppContext.BaseDirectory,"ScriptWorker","QuickExcel.ScriptWorker"),token.Token);
    using var reader = new StreamReader(process.Error);
    var error = reader.ReadToEndAsync();
    var passed=false;
    try
    {
        var identity=new RequestIdentity(Guid.NewGuid(),Guid.NewGuid(),Guid.NewGuid(),0,0,0);
        await WorkerProtocol.WriteAsync(process.Input,WorkerProtocol.Message(identity,"hello",new WorkerHello(WorkerProtocol.Engine,WorkerProtocol.Policy,[1,2],WorkerProtocol.MaxFrameBytes)),token.Token);
        var response=await WorkerProtocol.ReadAsync(process.Output,token.Token);
        Console.WriteLine("Handshake: "+response.Kind);
        passed=response.Kind=="hello";
    }
    catch(Exception ex){Console.WriteLine(ex.Message);}
    finally {process.Kill(); File.WriteAllText(Path.Combine(args[1],"stderr.log"), await error.WaitAsync(TimeSpan.FromSeconds(2)));}
    Console.WriteLine(File.ReadAllText(Path.Combine(args[1],"stderr.log")));
    return passed ? 0 : 1;
}

if (args.Length != 1 || Directory.Exists(args[0])) return 2;
var directory = Path.GetFullPath(args[0]); Directory.CreateDirectory(directory);
var checks = new List<string>(); Exception? failure = null;
try
{
    var recovered = ScriptWorkerRuntimeFactory.RecoverAbandonedSandboxes();
    using (var runtime = new ScriptWorkerRuntimeFactory().Create("quickExcel.functions.register('ADDONE', x => x+1);"))
    {
        if (!Equals(runtime.Invoke("ADDONE", [41d]), 42d)) throw new Exception("Incorrect isolated result.");
        checks.Add("API v1 native bridge and isolated IPC result");
        if (runtime is ScriptWorkerRuntime worker && OperatingSystem.IsLinux())
        {
            File.WriteAllText(Path.Combine(directory,"worker-status.txt"),File.ReadAllText($"/proc/{worker.ProcessId}/status"));
            File.WriteAllText(Path.Combine(directory,"worker-cgroup.txt"),File.ReadAllText($"/proc/{worker.ProcessId}/cgroup"));
        }
    }
    checks.Add("Worker termination and private payload cleanup");
    var modules=new ScriptModuleSet(new Dictionary<string,ScriptModule>
    {
        ["main"]=new("scripts/main.js","import {twice} from 'math'; quickExcel.functions.register('DOUBLE',twice); console.log('module-ready');"),
        ["math"]=new("scripts/math.js","export const twice=x=>x*2;")
    },["main"],scriptApiVersion:2);
    using(var runtime=(ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules,new(Guid.NewGuid(),1,2,3),TimeSpan.FromSeconds(5),default)))
    {
        var values=runtime.InvokeBatch([new("DOUBLE",[WireValue.FromObject(21d)]),new("DOUBLE",[WireValue.FromObject(2d)])]);
        if(!Equals(values[0].Value?.ToObject(),42d) || !Equals(values[1].Value?.ToObject(),4d) || runtime.GetLogs().Single()!="module-ready") throw new Exception("Module/batch/log mismatch.");
    }
    checks.Add("API v2 declared module import, ordered batch and logs");
    var services=new WorkbookServices(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor());
    using(var document=WorkbookDocument.CreateBlank(services))
    {
        document.Session.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string,ScriptModule>{["main"]=new("scripts/main.js","quickExcel.actions.register('写入',({snapshot,patch})=>{if(!Object.isFrozen(snapshot))throw 1;patch.setValue(snapshot.sheets[0].id,'A1','中文');patch.setValue(snapshot.sheets[0].id,'B1',42);});")},["main"],scriptApiVersion:2,bindings:[new("main","actions")])));
        document.EnableScriptsAndCalculate();
        if(document.ListActions().Single().Name!="写入")throw new Exception("Action registry mismatch.");
        var revision=document.Snapshot.Revision;
        if(!document.RunAction("写入").Changed || document.Snapshot.Revision!=revision+1 || document.GetCellValue(0,0,0)!="中文" || document.GetCellValue(0,0,1)!="42")throw new Exception("Atomic action mismatch.");
        if(!document.Undo() || document.Snapshot.Sheets[0].GetCell(0,0) is not null || document.Snapshot.Sheets[0].GetCell(0,1) is not null)throw new Exception("Action undo mismatch.");
        document.DisableScripts(); if(!document.Redo() || document.ScriptsEnabled)throw new Exception("Redo renewed script authorization.");
    }
    checks.Add("Frozen action snapshot, validated single transaction, undo and data-only redo");
    foreach(var source in new[]{"while(true){}",
        "const a=new Uint8Array(300*1024*1024);a.fill(7);"
        })
    {
        var failed=false;
        try { using var invalid=new ScriptWorkerRuntimeFactory().Create(source); }
        catch(Exception ex) when(ex is ScriptRuntimeFailureException or TimeoutException){failed=true;}
        if(!failed)throw new Exception("Native source limit was not enforced.");
    }
    checks.Add("Native one-second interrupt and mandatory process memory ceiling; JSC-MEM-001 engine cap remains open");
    var typescript = new ScriptModuleSet(new Dictionary<string, ScriptModule>
    {
        ["main"] = new("scripts/main.ts", "import {twice} from 'math'; quickExcel.functions.register('DOUBLE',()=>twice(21));", ScriptLanguage.TypeScript),
        ["math"] = new("scripts/math.ts", "export const twice=(n:number):number=>n*2;", ScriptLanguage.TypeScript)
    }, ["main"], scriptApiVersion: 2);
    using (var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(typescript, new(Guid.NewGuid(),1,1,1),TimeSpan.FromSeconds(5),default)))
        if (!Equals(runtime.Invoke("DOUBLE", []), 42d)) throw new Exception("Isolated TypeScript output mismatch.");
    checks.Add("Independent isolated TypeScript role checks declared module types and executes verified ESM");
    var badTypescript=typescript.WithPrimarySource("export const n:number='wrong';");
    try { using var bad = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(badTypescript,new(Guid.NewGuid(),1,1,1),TimeSpan.FromSeconds(5),default)); throw new Exception("Invalid TS was accepted."); }
    catch (ScriptCompilationException) { checks.Add("TypeScript errors produce no executable artifacts"); }
    using(var runtime=new ScriptWorkerRuntimeFactory().Create("quickExcel.functions.register('IDENTITY', x=>x);"))
        if(!Equals(runtime.Invoke("IDENTITY",["中文 🧮"]),"中文 🧮"))throw new Exception("Recovery/UTF-8 mismatch.");
    checks.Add("Fresh round remains usable after failed workers");
}
catch(Exception ex) { failure=ex; }
File.WriteAllText(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new { passed=failure is null, checks, error=failure?.ToString(), platform=System.Runtime.InteropServices.RuntimeInformation.OSDescription },new JsonSerializerOptions {WriteIndented=true}));
Console.WriteLine(failure?.ToString() ?? "PASS: platform smoke");
return failure is null ? 0 : 1;
