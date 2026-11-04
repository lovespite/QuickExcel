using System.Buffers.Binary;
using System.Runtime.InteropServices;
using QuickExcel.ScriptContracts;

// Deliberately faulty, test-only peer. Never copied into Desktop or a release package.
using var input = Console.OpenStandardInput();
using var output = Console.OpenStandardOutput();
var hello = await WorkerProtocol.ReadAsync(input);
await WorkerProtocol.WriteAsync(output, hello);
var loadEnvelope = await WorkerProtocol.ReadAsync(input);
if (loadEnvelope.Kind is "loadActions" or "loadWorkbook")
{
    var modules = WorkerProtocol.Payload<LoadModules>(loadEnvelope);
    var fault = modules.Modules[0].Source.Split('\n')[0];
    if (loadEnvelope.Kind == "loadWorkbook")
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(loadEnvelope.Identity, "loadWorkbook", new LoadedWorkbook(modules.Sha256, [], [new("A", "")])));
    else await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(loadEnvelope.Identity, "loadActions", new LoadedActions(modules.Sha256, [new("A", "")])));
    var envelope = await WorkerProtocol.ReadAsync(input);
    var action = WorkerProtocol.Payload<RunAction>(envelope);
    if (fault.StartsWith("model-", StringComparison.Ordinal))
    {
        // Stage one valid mutation before injecting a bad callback. No partial edit may escape.
        var sheet = action.Snapshot.Sheets[0].Id;
        var first = new ModelCall(action.OperationId, 1, WorkerProtocol.EncodePayload(new { method = "setValue", args = new object[] { sheet, "A1", 9 } }));
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(envelope.Identity, "model", first));
        var firstReply = WorkerProtocol.Payload<ModelReply>(await WorkerProtocol.ReadAsync(input));
        if (!firstReply.Ok) throw new InvalidOperationException("Fault fixture could not stage its initial write.");
        var next = new ModelCall(action.OperationId, 2, WorkerProtocol.EncodePayload(new { method = "setValue", args = new object[] { sheet, "A1", 10 } }));
        next = fault switch
        {
            "model-operation" => next with { OperationId = Guid.NewGuid() },
            "model-sequence" => next with { Sequence = 1 },
            "model-method" => next with { Request = WorkerProtocol.EncodePayload(new { method = "setUnknown", args = new object[] { sheet, "A1", 10 } }) },
            "model-sheet" => next with { Request = WorkerProtocol.EncodePayload(new { method = "setValue", args = new object[] { Guid.NewGuid().ToString(), "A1", 10 } }) },
            "model-value" => next with { Request = WorkerProtocol.EncodePayload(new { method = "setValue", args = new object[] { sheet, "A1", new { kind = "number", value = "wrong" } } }) },
            _ => next
        };
        var modelIdentity = fault == "model-round" ? envelope.Identity with { RoundId = Guid.NewGuid() } : envelope.Identity;
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(modelIdentity, "model", next));
        await WorkerProtocol.ReadAsync(input);
        // Pretend the script caught an error and succeeded; host validation must still reject the transaction.
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(envelope.Identity, "runAction", new ActionPatch(action.OperationId, action.Snapshot.BaseRevision, action.Name, [])));
        var logs = await WorkerProtocol.ReadAsync(input);
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(logs.Identity, "logs", new WorkerLogs([])));
        var end = await WorkerProtocol.ReadAsync(input);
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(end.Identity, "endRound", new EmptyPayload()));
        return;
    }
    if (fault is "failure-log-count" or "failure-log-bytes")
    {
        var logs = fault == "failure-log-count" ? Enumerable.Repeat("line", 101).ToArray() : [new string('中', 2000)];
        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(envelope.Identity, "failure", new WorkerFailure("Action", "injected") { Logs = logs }));
        return;
    }
    var good = new ActionWrite(action.Snapshot.Sheets[0].Id, 0, 0, WireValue.FromObject(1d), null);
    var bad = good with { Row = 1 };
    bad = fault switch
    {
        "duplicate" => good,
        "sheet" => bad with { SheetId = Guid.NewGuid().ToString() },
        "coordinate" => bad with { Column = 1024 },
        "type" => bad with { Value = new WireValue("number", System.Text.Json.JsonSerializer.SerializeToElement("wrong")) },
        "both" => bad with { Formula = "=1" },
        "formula" => bad with { Value = null, Formula = "=1+" },
        "error" => bad with { Value = new WireValue("error", System.Text.Json.JsonSerializer.SerializeToElement("unknown")) },
        _ => bad
    };
    var patch = new ActionPatch(action.OperationId, action.Snapshot.BaseRevision, action.Name, [good, bad]);
    patch = fault switch
    {
        "operation" => patch with { OperationId = Guid.NewGuid() },
        "base" => patch with { BaseRevision = patch.BaseRevision + 1 },
        "name" => patch with { Name = "wrong" },
        "count" => patch with { Edits = Enumerable.Repeat(good, 20001).ToArray() },
        _ => patch
    };
    var peer = fault switch
    {
        "session" => envelope.Identity with { SessionId = Guid.NewGuid() },
        "worker" => envelope.Identity with { WorkerInstanceId = Guid.NewGuid() },
        "request" => envelope.Identity with { RequestId = Guid.NewGuid() },
        "revision" => envelope.Identity with { Revision = envelope.Identity.Revision + 1 },
        "generation" => envelope.Identity with { ScriptGeneration = envelope.Identity.ScriptGeneration + 1 },
        "calculation" => envelope.Identity with { CalculationGeneration = 1 },
        _ => envelope.Identity
    };
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(peer, "runAction", patch));
    var logRequest = await WorkerProtocol.ReadAsync(input);
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(logRequest.Identity, "logs", new WorkerLogs([])));
    return;
}
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
if (mode == "isolation")
{
    string result;
    try { result = IsolationChecks.Run(WorkerProtocol.Payload<InvokeBatch>(invoke).Calls.Single()); }
    catch (Exception e) { result = System.Text.Json.JsonSerializer.Serialize(new Dictionary<string, bool> { [e.ToString()] = false }); }
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "invokeBatch", new BatchResult([new(WireValue.FromObject(result), null)])));
    return;
}
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
else if (mode == "bad-limit")
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "invokeBatch", new BatchResult(
        WorkerProtocol.Payload<InvokeBatch>(invoke).Calls.Select(_ => new InvocationResult(WireValue.FromObject(1d), null, ValueTooLong: true)).ToArray())));
else if (mode == "bad-logs")
    await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(invoke.Identity, "logs", new WorkerLogs(Enumerable.Repeat("x", 101).ToArray())));
