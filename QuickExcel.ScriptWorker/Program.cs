using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;

// stdout is exclusively a binary protocol pipe. No workbook data is written to stderr.
if (args.SequenceEqual(["--compile"])) return await CompilerWorker.Run();
if (args.Length != 0 || !Console.IsInputRedirected || !Console.IsOutputRedirected) return 2;
// No request, source, or native JavaScript is accepted before the mandatory
// Linux OS guard verifies launch restrictions and synchronizes the final filter.
try { LinuxWorkerGuard.Enter(); } catch { return 125; }
using var input = Console.OpenStandardInput();
using var output = Console.OpenStandardOutput();
JscSession? runtime = null;
var actionsMode = false;
var sharedMode = false;
var roundOpen = true;
var actionRan = false;
RequestIdentity? round = null;
var requests = new HashSet<Guid>();
var timer = Stopwatch.StartNew();
try
{
    while (true)
    {
        var request = await WorkerProtocol.ReadAsync(input);
        if (request.Kind == "beginRound")
        {
            _ = WorkerProtocol.Payload<BeginWorkerRound>(request);
            if (!sharedMode || roundOpen || runtime is null || round is null || request.Identity.SessionId != round.SessionId ||
                request.Identity.WorkerInstanceId != round.WorkerInstanceId || request.Identity.ScriptGeneration != round.ScriptGeneration || request.Identity.RoundId == round.RoundId ||
                actionsMode != (request.Identity.CalculationGeneration is null))
                throw new ProtocolException("Invalid shared-worker round.");
            round = request.Identity; requests.Clear(); requests.Add(request.Identity.RequestId);
            roundOpen = true; actionRan = false; timer.Restart(); runtime.BeginRound();
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "beginRound", new EmptyPayload()));
            continue;
        }
        if (requests.Count >= 100_000 || !requests.Add(request.Identity.RequestId) ||
            round is not null && !round.SameRound(request.Identity))
            throw new ProtocolException("Duplicate request or changed round identity.");
        if (!roundOpen && request.Kind != "shutdown") throw new ProtocolException("Worker has no active round.");
        if (round is null)
        {
            if (request.Kind != "hello") throw new ProtocolException("Handshake required.");
            var hello = WorkerProtocol.Payload<WorkerHello>(request);
            if (hello.Engine != WorkerProtocol.Engine || hello.Policy != WorkerProtocol.Policy ||
                hello.Role != "execute" || hello.Compiler is not null ||
                !hello.ApiVersions.SequenceEqual([3]) || hello.MaxFrameBytes != WorkerProtocol.MaxFrameBytes)
                throw new ProtocolException("Incompatible engine, API or execution policy.");
            round = request.Identity;
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(round, "hello", hello));
            continue;
        }
        if (request.Kind is "load" or "loadModules" or "loadActions" or "loadWorkbook" && runtime is null)
        {
            sharedMode = request.Kind == "loadWorkbook";
            actionsMode = request.Kind == "loadActions" || sharedMode && request.Identity.CalculationGeneration is null;
            if (actionsMode != (request.Identity.CalculationGeneration is null)) throw new ProtocolException("Context identity disagrees with load kind.");
            string digest;
            Func<JscSession> create;
            if (request.Kind is "loadModules" or "loadActions" or "loadWorkbook")
            {
                var bundle = WorkerProtocol.Payload<LoadModules>(request);
                bundle.Validate();
                if (sharedMode && (actionsMode
                    ? bundle.FunctionEntries.Length != 0 || !bundle.EntryModules.SequenceEqual(bundle.ActionEntries)
                    : bundle.ActionEntries.Length != 0 || !bundle.EntryModules.SequenceEqual(bundle.FunctionEntries)))
                    throw new ProtocolException("A reusable worker must load exactly one bound role.");
                digest = bundle.Sha256;
                create = () => new JscSession(bundle.Modules.ToDictionary(m => m.Id, m => m.Source, StringComparer.Ordinal), bundle.EntryModules, actionsMode,
                    sharedMode ? bundle.Modules.ToDictionary(m => m.Id, m => (bundle.FunctionEntries.Contains(m.Id) ? 1 : 0) | (bundle.ActionEntries.Contains(m.Id) ? 2 : 0)) : null,
                    bundle.Modules.ToDictionary(m => m.Id, m => m.Description, StringComparer.Ordinal));
            }
            else
            {
                var load = WorkerProtocol.Payload<LoadScript>(request);
                if (Encoding.UTF8.GetByteCount(load.Source) > 1024 * 1024 ||
                    Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(load.Source))) != load.Sha256)
                    throw new ProtocolException("Invalid source digest or source budget.");
                digest = load.Sha256;
                create = () => new JscSession(load.Source);
            }
            try { runtime = create(); }
            catch (Exception error)
            {
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                    new WorkerFailure("Load", error is InvalidOperationException ? Bounded("脚本加载失败：" + error.Message)
                        : "脚本加载失败；请检查语法、包内模块 ID、同步初始化、注册及资源限额。")));
                return 3;
            }
            if (runtime.FunctionDescriptors.Count > WorkerProtocol.MaxFunctions) throw new ProtocolException("Too many functions.");
            var descriptors = runtime.FunctionDescriptors.Select(d => new FunctionInfo(d.Name, d.ApiVersion,
                d.MinArguments, d.MaxArguments, d.ReturnType, d.Profile)).ToArray();
            if (sharedMode)
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, request.Kind,
                    new LoadedWorkbook(digest, descriptors, runtime.ActionDescriptors.Select(a => new ActionInfo(a.Name, a.Description)).ToArray())));
            else if (actionsMode)
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, request.Kind,
                    new LoadedActions(digest, runtime.ActionDescriptors.Select(a => new ActionInfo(a.Name, a.Description)).ToArray())));
            else await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, request.Kind, new LoadedScript(digest, descriptors)));
        }
        else if (request.Kind == "invokeBatch" && runtime is not null && !actionsMode && request.Identity.CalculationGeneration is not null)
        {
            var batch = WorkerProtocol.Payload<InvokeBatch>(request);
            if (batch.Calls.Length is < 1 or > WorkerProtocol.MaxBatchCalls) throw new ProtocolException("Invalid batch size.");
            // Preflight the entire batch before invoking any user code.
            foreach (var call in batch.Calls)
            {
                if (call is null || call.Name.Length is < 1 or > 256 || call.Arguments.Length > 255)
                    throw new ProtocolException("Invalid function call.");
                foreach (var argument in call.Arguments) (argument ?? throw new ProtocolException("Missing argument.")).Validate();
            }
            var results = new List<InvocationResult>();
            foreach (var call in batch.Calls)
            {
                if (timer.Elapsed > TimeSpan.FromSeconds(5)) throw new TimeoutException();
                try
                {
                    var value = runtime.Invoke(call.Name, call.Arguments.Select(a => a.ToObject()!).ToArray());
                    results.Add(value is string text && text.Length > WorkerProtocol.MaxCellTextLength
                        ? new(null, null, ValueTooLong: true) : new(WireValue.FromObject(value), null));
                }
                catch (InvalidOperationException e)
                {
                    // Interruption is fatal for the whole round; never reuse an interrupted native handle.
                    if (runtime.LastOperationInterrupted)
                    {
                        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                            new WorkerFailure("Timeout", "脚本达到单次执行时间限制。")));
                        return 5;
                    }
                    if (runtime.LastOperationInvalidated)
                    {
                        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                            new WorkerFailure("Invalidated", "脚本运行时已失效，本轮结果未提交。")));
                        return 6;
                    }
                    results.Add(new(null, Bounded(e.Message)));
                }
            }
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "invokeBatch", new BatchResult(results.ToArray())));
        }
        else if (request.Kind == "runAction" && runtime is not null && actionsMode && request.Identity.CalculationGeneration is null && !actionRan)
        {
            actionRan = true;
            var action = WorkerProtocol.Payload<RunAction>(request);
            ActionProtocol.Name(action.Name);
            ActionProtocol.ValidateSnapshot(action.Snapshot);
            VirtualFileProtocol.Validate(action.Files);
            if (action.Event is { } ev && (action.Name != "@event/" + ev.Name || ev.SessionId != request.Identity.SessionId || ev.Revision > request.Identity.Revision || ev.Sequence <= 0 || ev.Ranges is null || ev.Ranges.Length > 64) ||
                action.Event is null && action.Name.StartsWith("@event/", StringComparison.Ordinal)) throw new ProtocolException("Invalid event identity.");
            if (action.OperationId == Guid.Empty || action.Snapshot.BaseRevision != request.Identity.Revision) throw new ProtocolException("Action revision mismatch.");
            try
            {
                var modelSequence = 0;
                runtime.HostRequest = action.LiveModel ? json =>
                {
                    if (++modelSequence > 4096) throw new InvalidOperationException("模型请求超过 4096 次。");
                    using var parsed = JsonDocument.Parse(json);
                    var call = new ModelCall(action.OperationId, modelSequence, parsed.RootElement.Clone());
                    WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "model", call)).GetAwaiter().GetResult();
                    var response = WorkerProtocol.ReadAsync(input).GetAwaiter().GetResult();
                    WorkerProtocol.ValidateReply(response, request.Identity, "model");
                    var reply = WorkerProtocol.Payload<ModelReply>(response);
                    if (reply.OperationId != action.OperationId || reply.Sequence != modelSequence) throw new ProtocolException("Model reply identity mismatch.");
                    return JsonSerializer.Serialize(new { ok = reply.Ok, value = reply.Value, error = reply.Error, code = reply.Code }, WorkerProtocol.Json);
                } : null;
                var raw = runtime.RunAction(action.Name, WorkerProtocol.EncodePayload(action));
                var patch = WorkerProtocol.Payload<ActionPatch>(request with { Payload = raw });
                ActionProtocol.ValidatePatch(patch, action);
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "runAction", patch));
            }
            catch (InvalidOperationException error)
            {
                // Interrupted/invalidated contexts cannot be called again. Ordinary
                // script errors still expose the bounded output printed before failure.
                string[] failureLogs = [];
                if (!runtime.LastOperationInvalidated)
                    try { failureLogs = runtime.GetLogs().Take(WorkerProtocol.MaxLogs).Select(Bounded).ToArray(); } catch { }
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                    new WorkerFailure(runtime.LastOperationInterrupted ? "Timeout" : "Action", Bounded("操作脚本失败，本次变更未提交：" + error.Message)) { Logs = failureLogs }));
                return 7;
            }
            finally { runtime.HostRequest = null; }
        }
        else if (request.Kind == "logs" && runtime is not null)
        {
            _ = WorkerProtocol.Payload<EmptyPayload>(request);
            var logs = runtime.GetLogs().Take(WorkerProtocol.MaxLogs).Select(Bounded).ToArray();
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "logs", new WorkerLogs(logs)));
        }
        else if (request.Kind == "endRound" && sharedMode && runtime is not null)
        {
            _ = WorkerProtocol.Payload<EmptyPayload>(request);
            roundOpen = false;
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "endRound", new EmptyPayload()));
        }
        else if (request.Kind == "shutdown")
        {
            _ = WorkerProtocol.Payload<EmptyPayload>(request);
            return 0;
        }
        else throw new ProtocolException("Request is invalid in this worker state.");
    }
}
catch (EndOfStreamException) { return 0; }
catch (Exception) { return 4; }
finally { runtime?.Dispose(); }

static string Bounded(string text)
{
    if (Encoding.UTF8.GetByteCount(text) <= WorkerProtocol.MaxLogBytes) return text;
    var builder = new StringBuilder(); var bytes = 0;
    foreach (var rune in text.EnumerateRunes())
    {
        if (bytes + rune.Utf8SequenceLength > WorkerProtocol.MaxLogBytes) break;
        builder.Append(rune); bytes += rune.Utf8SequenceLength;
    }
    return builder.ToString();
}
