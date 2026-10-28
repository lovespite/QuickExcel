using System.Runtime.InteropServices;
using System.Text.Json;
using QuickExcel.Runtime;

SetErrorMode(0x0002); // Test failures must never display Windows Error Reporting dialogs.
try { await Run(args); return 0; }
catch (Exception error) { Console.Error.WriteLine(error); return 1; }

static async Task Run(string[] args)
{
// Test-only host: the parent kills this process, so no Dispose/finally runs.
if (args.Length == 0)
{
    // Fault-injected worker ignores pipe EOF, so only the Job can stop it after host death.
    Console.WriteLine("ready"); Console.Out.Flush();
    await Task.Delay(Timeout.Infinite);
    return;
}
var root = args[0];
var stage = args[1];
if (stage == "recover") { Console.WriteLine(WorkerSandboxLease.Recover(root: root)); return; }
string profile, payload;
var pid = 0;
object? owner = null;
if (stage is "record" or "profile" or "partial")
{
    var lease = WorkerSandboxLease.Create(default, root);
    owner = lease;
    profile = lease.ProfileName;
    payload = lease.PayloadDirectory;
    if (stage != "record")
    {
        Marshal.ThrowExceptionForHR(CreateAppContainerProfile(profile, profile, "QuickExcel crash test", IntPtr.Zero, 0, out var sid));
        FreeSid(sid);
    }
    if (stage == "partial")
    {
        WorkerSandboxLease.CreatePrivateDirectory(payload);
        File.WriteAllText(Path.Combine(payload, "partial.dll"), "interrupted copy");
    }
}
else if (stage == "staged")
{
    var sandbox = WorkerSandbox.Create(args[2], default, root);
    owner = sandbox; profile = sandbox.ProfileName; payload = sandbox.PayloadDirectory;
}
else
{
    var worker = WorkerProcess.Start(stage == "busy" ? Environment.ProcessPath! : args[2], default, root);
    owner = worker; profile = worker.ProfileName; payload = worker.PayloadDirectory; pid = worker.Id;
    if (stage == "busy")
    {
        var reader = new StreamReader(worker.Output);
        if (await reader.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(5)) != "ready") throw new IOException("Worker did not start.");
    }
}
Console.WriteLine(JsonSerializer.Serialize(new { Profile = profile, Payload = payload, ProcessId = pid }));
Console.Out.Flush();
await Task.Delay(Timeout.Infinite);
GC.KeepAlive(owner);
}

[DllImport("kernel32.dll")] static extern uint SetErrorMode(uint mode);
[DllImport("userenv.dll", CharSet = CharSet.Unicode)]
static extern int CreateAppContainerProfile(string name, string display, string description, IntPtr capabilities, uint count, out IntPtr sid);
[DllImport("advapi32.dll")] static extern IntPtr FreeSid(IntPtr sid);
