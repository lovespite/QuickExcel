using System.Diagnostics;
using System.Text;
using System.Text.Json;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Calculation;
using QuickExcel.Storage;
using System.Security.Cryptography;

internal static class JscPerformance
{
    internal static async Task<int> Run(string directory)
    {
        if (Directory.Exists(directory)) return 2;
        Directory.CreateDirectory(directory);
        var samples = new List<object>();
        Exception? failure = null;
        try
        {
            for (var iteration = 0; iteration < 3; iteration++)
            foreach (var scenario in new[] { "small", "64-modules", "1m-source", "256-calls", "cpu-calls", "ts-small", "ts-dense-1m" })
            {
                var compile = scenario.StartsWith("ts-", StringComparison.Ordinal);
                var source = "quickExcel.functions.register('F',x=>x+1);";
                if (scenario == "1m-source") source += "\n//" + new string('x', 1024 * 1024 - 128);
                if (scenario == "cpu-calls") source = "quickExcel.functions.register('F',x=>{for(let i=0;i<10000;i++)x=(x+i)%997;return x;});";
                var modules = new List<ModuleSource> { new("main", "scripts/main.js", source) };
                if (scenario == "64-modules")
                {
                    for (var i = 0; i < 63; i++) modules.Add(new("m"+i, "scripts/m"+i+".js", "export const value="+i+";"));
                    modules[0] = modules[0] with { Source = string.Concat(Enumerable.Range(0, 63).Select(i => "import 'm"+i+"';\n")) + source };
                }
                CompileRequest? request = null;
                if (compile)
                {
                    var ts = new StringBuilder("export const value:number=3;\n");
                    if (scenario == "ts-dense-1m")
                        for (var i=0;;i++) { var line=$"export const v{i}:number={i};\n"; if(ts.Length+line.Length>1024*1024)break; ts.Append(line); }
                    request = new([new("main", "typescript", ts.ToString())], ["main"], "");
                    request = request with { Sha256 = request.ComputeDigest() };
                }
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(compile ? 6 : 5));
                var clock = Stopwatch.StartNew();
                var executable = Path.Combine(AppContext.BaseDirectory,"ScriptWorker",OperatingSystem.IsWindows()?"QuickExcel.ScriptWorker.exe":"QuickExcel.ScriptWorker");
                IWorkerProcess process = OperatingSystem.IsWindows() ? WorkerProcess.Start(executable,deadline.Token,compilation:compile) : OperatingSystem.IsLinux() ? LinuxWorkerProcess.Start(executable,deadline.Token,compile) : throw new PlatformNotSupportedException();
                var launchMs=clock.Elapsed.TotalMilliseconds;
                using var stop=deadline.Token.Register(process.Kill);
                using var sampled=Process.GetProcessById(process.Id);
                using var sampling=new CancellationTokenSource();
                long peakWorking=0, peakCommit=0; int sampledThreads=0;
                var observer=Task.Run(async()=>
                {
                    try { while(!sampling.IsCancellationRequested) { sampled.Refresh(); peakWorking=Math.Max(peakWorking,sampled.PeakWorkingSet64); if(OperatingSystem.IsWindows())peakCommit=Math.Max(peakCommit,sampled.PeakPagedMemorySize64); sampledThreads=Math.Max(sampledThreads,sampled.Threads.Count); await Task.Delay(5,sampling.Token); } }
                    catch(Exception e) when(e is InvalidOperationException or System.ComponentModel.Win32Exception or OperationCanceledException) { }
                });
                var stderr=new StreamReader(process.Error).ReadToEndAsync();
                var identity=new RequestIdentity(Guid.NewGuid(),Guid.NewGuid(),Guid.NewGuid(),1,1,1);
                async Task<WorkerEnvelope> Exchange<T>(string kind,T payload)
                {
                    identity=identity with { RequestId=Guid.NewGuid() };
                    await WorkerProtocol.WriteAsync(process.Input,WorkerProtocol.Message(identity,kind,payload),deadline.Token);
                    var reply=await WorkerProtocol.ReadAsync(process.Output,deadline.Token);
                    WorkerProtocol.ValidateReply(reply,identity,kind);
                    if(reply.Kind=="failure")throw new Exception(WorkerProtocol.Payload<WorkerFailure>(reply).Message);
                    return reply;
                }
                double helloMs=0,loadMs=0,firstBatchMs=0,repeatBatchMs=0,compilerInternalMs=0,cleanupMs=0;
                string? sampleError=null;
                try
                {
                    var phase=Stopwatch.StartNew();
                    var expected=compile?CompilerProtocol.Hello:new WorkerHello(WorkerProtocol.Engine,WorkerProtocol.Policy,[1,2],WorkerProtocol.MaxFrameBytes);
                    var accepted=WorkerProtocol.Payload<WorkerHello>(await Exchange("hello",expected));
                    if(JsonSerializer.Serialize(accepted,WorkerProtocol.Json)!=JsonSerializer.Serialize(expected,WorkerProtocol.Json))throw new Exception("Handshake mismatch.");
                    if(OperatingSystem.IsLinux() && process is LinuxWorkerProcess linux)linux.VerifyBeforeSource();
                    helloMs=phase.Elapsed.TotalMilliseconds; phase.Restart();
                    if(compile)
                    {
                        var reply=WorkerProtocol.Payload<CompiledReply>(await Exchange("compile",request!));
                        CompilerProtocol.ValidateReply(reply,request!);
                        if(!reply.Result.Success)throw new Exception("Benchmark input did not compile.");
                        compilerInternalMs=reply.Result.Milliseconds;
                    }
                    else
                    {
                        var bundle=new LoadModules(modules.ToArray(),["main"],""); bundle=bundle with {Sha256=bundle.ComputeDigest()};
                        var loaded=WorkerProtocol.Payload<LoadedScript>(await Exchange("loadModules",bundle));
                        if(loaded.Sha256!=bundle.Sha256)throw new Exception("Loaded source mismatch.");
                    }
                    loadMs=phase.Elapsed.TotalMilliseconds;
                    if(!compile)
                    {
                        var calls=new InvokeBatch(Enumerable.Range(0,scenario is "256-calls" or "cpu-calls"?256:1).Select(i=>new FunctionCall("F",[WireValue.FromObject((double)i)])).ToArray());
                        for(var repeat=0;repeat<2;repeat++)
                        {
                            phase.Restart(); var result=WorkerProtocol.Payload<BatchResult>(await Exchange("invokeBatch",calls));
                            if(result.Results.Length!=calls.Calls.Length || result.Results.Any(r=>r.Error is not null || r.Value?.Kind!="number"))throw new Exception("Batch mismatch.");
                            if(repeat==0)firstBatchMs=phase.Elapsed.TotalMilliseconds;else repeatBatchMs=phase.Elapsed.TotalMilliseconds;
                        }
                    }
                }
                catch(Exception e){failure=e;sampleError=e.ToString();}
                finally
                {
                    sampling.Cancel(); await observer;
                    var cleanup=Stopwatch.StartNew(); process.Dispose(); cleanupMs=cleanup.Elapsed.TotalMilliseconds;
                }
                var error=await stderr;
                samples.Add(new{iteration,scenario,sourceSha256=request?.Sha256 ?? new LoadModules(modules.ToArray(),["main"],"").ComputeDigest(),launchMs,helloMs,loadMs,firstBatchMs,repeatBatchMs,compilerInternalMs,cleanupMs,totalMs=clock.Elapsed.TotalMilliseconds,peakWorkingSetBytes=peakWorking,peakCommitBytes=OperatingSystem.IsWindows()?(long?)peakCommit:null,sampledMaxThreads=sampledThreads,stderr=error,error=sampleError});
                Console.WriteLine($"{scenario} {iteration}: {clock.Elapsed.TotalMilliseconds:F1} ms");
            }
            for(var iteration=0;iteration<3;iteration++)
            {
                var services=new WorkbookServices(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor());
                using var document=WorkbookDocument.CreateBlank(services);
                var source="quickExcel.actions.register('A',({fs,snapshot,patch})=>{fs.writeText('//mem/payload','x'.repeat(128*1024));patch.setValue(snapshot.sheets[0].id,'A1',42);});quickExcel.events.on('cell.click',({fs,snapshot,patch})=>{patch.setValue(snapshot.sheets[0].id,'B1',fs.readText('//mem/payload').length);});";
                document.Session.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string,ScriptModule>{["main"]=new("scripts/main.js",source)},["main"],scriptApiVersion:2,bindings:[new("main","actions")])));
                document.EnableScriptsAndCalculate();document.DispatchScriptEvents();
                var timer=Stopwatch.StartNew();document.RunAction("A");var actionMs=timer.Elapsed.TotalMilliseconds;
                document.Session.RaiseScriptEvent("cell.click",document.Snapshot.ActiveSheetId,["A1"]);
                timer.Restart();document.DispatchScriptEvents();var eventMs=timer.Elapsed.TotalMilliseconds;
                if(document.GetCellValue(0,0,0)!="42" || document.GetCellValue(0,0,1)!="131072")throw new Exception("Action/event result mismatch.");
                using var cancel=new CancellationTokenSource();
                var runtime=new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest("quickExcel.functions.register('F',()=>{while(true){}});",new(Guid.NewGuid(),1,1,1),TimeSpan.FromSeconds(5),cancel.Token));
                var invocation=Task.Run(()=>{try{runtime.Invoke("F",[]);}catch(OperationCanceledException){return;}throw new Exception("Cancellation was not observed.");});
                await Task.Delay(20);timer.Restart();cancel.Cancel();await invocation;runtime.Dispose();var cancellationMs=timer.Elapsed.TotalMilliseconds;
                var compiler=new ScriptCompiler();var request=new CompileRequest([new("main","typescript","export const n:number=1;")],["main"],"");request=request with {Sha256=request.ComputeDigest()};
                await compiler.CompileAsync(request,new(Guid.NewGuid(),1,1,1));timer.Restart();await compiler.CompileAsync(request,new(Guid.NewGuid(),2,1,1));var cacheHitMs=timer.Elapsed.TotalMilliseconds;
                if(compiler.ProcessLaunchCount!=1)throw new Exception("Compiler cache was not reused.");
                samples.Add(new{iteration,scenario="actions-events-cancel-cache",actionMs,eventMs,cancellationMs,cacheHitMs,virtualSnapshotBytes=128*1024});
            }
        }
        catch(Exception e){failure=e;}
        var payloadHashes=Directory.EnumerateFiles(Path.Combine(AppContext.BaseDirectory,"ScriptWorker")).ToDictionary(p=>Path.GetFileName(p),p=>{using var stream=File.OpenRead(p);return Convert.ToHexString(SHA256.HashData(stream));});
        File.WriteAllText(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new{passed=failure is null,engine=WorkerProtocol.Engine,policy=WorkerProtocol.Policy,compiler=CompilerProtocol.Identity,platform=System.Runtime.InteropServices.RuntimeInformation.OSDescription,logicalProcessors=Environment.ProcessorCount,payloadHashes,measurement="Sequential fresh isolated worker samples; launch/hello/source-load/batch/cleanup wall times include IPC. OS peak memory and 5 ms sampled thread counts; three repetitions are not p95.",samples,error=failure?.ToString()},new JsonSerializerOptions{WriteIndented=true}));
        if(failure is not null)Console.Error.WriteLine(failure);
        return failure is null?0:1;
    }
}
