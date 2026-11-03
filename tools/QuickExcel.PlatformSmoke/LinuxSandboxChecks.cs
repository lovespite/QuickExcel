using System.Diagnostics;
using System.Runtime.Versioning;
using System.Text.Json;
using QuickExcel.Runtime;

[SupportedOSPlatform("linux")]
internal static class LinuxSandboxChecks
{
    internal static async Task<int> Run(string directory, bool hold = false, bool compilationOnly = false)
    {
        if (Directory.Exists(directory)) return 2;
        Directory.CreateDirectory(directory); directory=Path.GetFullPath(directory);
        var results=new List<object>(); Exception? failure=null;
        try
        {
            var canary=Path.Combine(directory,"outside-canary.txt"); File.WriteAllText(canary,"controlled host-only canary");
            foreach(var compilation in hold ? new[]{compilationOnly} : new[]{false,true})
            foreach(var mode in hold ? new[]{"parent"} : new[]{"policy","threads","memory","timeout"})
            {
                var name=(compilation ? "compile-" : "execute-")+mode;
                var fixture=Path.Combine(directory,"fixture-"+name); Directory.CreateDirectory(fixture);
                foreach(var file in Directory.EnumerateFiles(Path.Combine(AppContext.BaseDirectory,"ScriptWorker"))) File.Copy(file,Path.Combine(fixture,Path.GetFileName(file)));
                File.Copy(Path.Combine(AppContext.BaseDirectory,"TestProbe","QuickExcel.ScriptWorker"),Path.Combine(fixture,"QuickExcel.ScriptWorker"),true);
                File.WriteAllText(Path.Combine(fixture,"probe.mode"),mode+"\n"+canary+"\n");
                using var process=LinuxWorkerProcess.Start(Path.Combine(fixture,"QuickExcel.ScriptWorker"),default,compilation);
                using var output=new StreamReader(process.Output); using var error=new StreamReader(process.Error);
                if(await output.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(2)) != "READY "+mode) throw new Exception("Native guard failed before probe: "+await error.ReadToEndAsync());
                process.VerifyBeforeSource();
                var membership=File.ReadAllLines($"/proc/{process.Id}/cgroup").Single(line=>line.StartsWith("0::",StringComparison.Ordinal))[3..];
                var group="/sys/fs/cgroup"+membership;
                var limits=new Dictionary<string,string>();
                foreach(var file in new[]{"memory.max","memory.swap.max","pids.max","memory.events","cgroup.procs"}) limits[file]=File.ReadAllText(Path.Combine(group,file));
                File.WriteAllText(Path.Combine(directory,name+"-status.txt"),File.ReadAllText($"/proc/{process.Id}/status"));
                process.Input.WriteByte((byte)'P'); process.Input.Flush();
                if(hold)
                {
                    File.WriteAllText(Path.Combine(directory,"live.json"),JsonSerializer.Serialize(new{host=Environment.ProcessId,worker=process.Id,start=LinuxWorkerLease.ReadStartTime(process.Id),scope=process.ProfileName,payload=process.PayloadDirectory,compilation}));
                    await Task.Delay(Timeout.Infinite); // External test kills this host, never the worker.
                }
                var timer=Stopwatch.StartNew();
                if(!process.WaitForExit(8000)) throw new Exception("Worker escaped scope lifetime limit.");
                var stdout=await output.ReadToEndAsync(); var stderr=await error.ReadToEndAsync(); var exit=process.ExitCode;
                var passed=mode switch {"policy"=>exit==0 && stdout.Contains("PASS policy 32"),"threads"=>exit==0 && stdout.Contains("PASS threads 32"),"memory"=>exit==137,"timeout"=>exit is 143 or 137 && timer.Elapsed.TotalSeconds >= (compilation ? 4 : 3) && timer.Elapsed.TotalSeconds < 8,_=>false};
                results.Add(new{mode,compilation,passed,exit,elapsedMs=timer.Elapsed.TotalMilliseconds,stdout,stderr,limits});
                if(!passed) throw new Exception($"Failed native {mode}: exit {exit}; {stdout} {stderr}");
            }
            if(File.ReadAllText(canary)!="controlled host-only canary") throw new Exception("Outside canary changed.");
        }
        catch(Exception ex){failure=ex;}
        File.WriteAllText(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new{passed=failure is null,results,error=failure?.ToString()},new JsonSerializerOptions{WriteIndented=true}));
        Console.WriteLine(failure?.ToString() ?? "PASS: native Linux sandbox checks");
        return failure is null ? 0 : 1;
    }
}
