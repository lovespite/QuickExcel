using System.Buffers.Binary;
using System.Runtime.InteropServices;
using QuickExcel.ScriptContracts;

// Deliberately faulty, test-only peer. Never copied into Desktop or a release package.
using var input = Console.OpenStandardInput();
using var output = Console.OpenStandardOutput();
var hello = await WorkerProtocol.ReadAsync(input);
await WorkerProtocol.WriteAsync(output, hello);
var loadEnvelope = await WorkerProtocol.ReadAsync(input);
var load = WorkerProtocol.Payload<LoadScript>(loadEnvelope);
var mode = load.Source;
if (mode == "duplicate") { await WorkerProtocol.WriteAsync(output, hello); return; }
if (mode == "oversize")
{
    var header = new byte[4]; BinaryPrimitives.WriteInt32LittleEndian(header, int.MaxValue);
    await output.WriteAsync(header); await output.FlushAsync(); return;
}
if (mode == "half") { await output.WriteAsync(new byte[] { 20, 0, 0, 0, 123 }); await output.FlushAsync(); return; }
if (mode == "stderr") { await Console.Error.WriteAsync(new string('x', 20000)); await Task.Delay(30000); return; }
if (mode == "hang-load") { await Task.Delay(30000); return; }
var identity = mode switch
{
    "session" => loadEnvelope.Identity with { SessionId = Guid.NewGuid() },
    "worker" => loadEnvelope.Identity with { WorkerInstanceId = Guid.NewGuid() },
    "request" => loadEnvelope.Identity with { RequestId = Guid.NewGuid() },
    "revision" => loadEnvelope.Identity with { Revision = 999 },
    "calculation" => loadEnvelope.Identity with { CalculationGeneration = 999 },
    "script" => loadEnvelope.Identity with { ScriptGeneration = 999 },
    _ => loadEnvelope.Identity
};
await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(identity, "load", new LoadedScript(load.Sha256,
    [new("F", 1, 0, 255, "any", "conservative")])));
var invoke = await WorkerProtocol.ReadAsync(input);
if (mode == "hang") { await Task.Delay(30000); return; }
if (mode == "memory")
{
    // A single allocation above the Job commit limit must fail even outside the JS engine.
    var denied = false;
    try { var pointer = Marshal.AllocHGlobal(384 * 1024 * 1024); Marshal.FreeHGlobal(pointer); }
    catch (OutOfMemoryException) { denied = true; }
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "invokeBatch", new BatchResult([new(WireValue.FromObject(denied), null)])));
    return;
}
if (mode == "bad-count")
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "invokeBatch", new BatchResult([])));
else if (mode == "bad-value")
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "invokeBatch", new BatchResult([new(new WireValue("number", System.Text.Json.JsonSerializer.SerializeToElement("wrong")), null)])));
else if (mode == "bad-logs")
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "logs", new WorkerLogs(Enumerable.Repeat("x", 101).ToArray())));
