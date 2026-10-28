using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;

// stdout is exclusively a binary protocol pipe. No workbook data is written to stderr.
if (args.Length != 0 || !Console.IsInputRedirected || !Console.IsOutputRedirected) return 2;
using var input = Console.OpenStandardInput();
using var output = Console.OpenStandardOutput();
QuickJsSession? runtime = null;
RequestIdentity? round = null;
var requests = new HashSet<Guid>();
var timer = Stopwatch.StartNew();
try
{
    while (true)
    {
        var request = await WorkerProtocol.ReadAsync(input);
        if (requests.Count >= 100_000 || !requests.Add(request.Identity.RequestId) ||
            round is not null && !round.SameRound(request.Identity))
            throw new ProtocolException("Duplicate request or changed round identity.");
        if (round is null)
        {
            if (request.Kind != "hello") throw new ProtocolException("Handshake required.");
            var hello = WorkerProtocol.Payload<WorkerHello>(request);
            if (hello.Engine != WorkerProtocol.Engine || hello.Policy != WorkerProtocol.Policy ||
                !hello.ApiVersions.SequenceEqual([1, 2]) || hello.MaxFrameBytes != WorkerProtocol.MaxFrameBytes)
                throw new ProtocolException("Incompatible engine, API or execution policy.");
            round = request.Identity;
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(round, "hello", hello));
            continue;
        }
        if (request.Kind is "load" or "loadModules" && runtime is null)
        {
            string digest;
            Func<QuickJsSession> create;
            if (request.Kind == "loadModules")
            {
                var bundle = WorkerProtocol.Payload<LoadModules>(request);
                bundle.Validate();
                digest = bundle.Sha256;
                create = () => new QuickJsSession(bundle.Modules.ToDictionary(m => m.Id, m => m.Source, StringComparer.Ordinal), bundle.EntryModules);
            }
            else
            {
                var load = WorkerProtocol.Payload<LoadScript>(request);
                if (Encoding.UTF8.GetByteCount(load.Source) > 1024 * 1024 ||
                    Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(load.Source))) != load.Sha256)
                    throw new ProtocolException("Invalid source digest or source budget.");
                digest = load.Sha256;
                create = () => new QuickJsSession(load.Source);
            }
            try { runtime = create(); }
            catch (Exception)
            {
                await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                    new WorkerFailure("Load", "脚本加载失败；请检查语法、包内模块 ID、同步初始化、注册及资源限额。")));
                return 3;
            }
            if (runtime.FunctionDescriptors.Count > WorkerProtocol.MaxFunctions) throw new ProtocolException("Too many functions.");
            var descriptors = runtime.FunctionDescriptors.Select(d => new FunctionInfo(d.Name, d.ApiVersion,
                d.MinArguments, d.MaxArguments, d.ReturnType, d.Profile)).ToArray();
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, request.Kind, new LoadedScript(digest, descriptors)));
        }
        else if (request.Kind == "invokeBatch" && runtime is not null)
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
                    results.Add(new(WireValue.FromObject(value), null));
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
                    if (e.Message.Contains("out of memory", StringComparison.OrdinalIgnoreCase) || e.Message.Contains("call stack size", StringComparison.OrdinalIgnoreCase))
                    {
                        await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "failure",
                            new WorkerFailure("ResourceLimit", "脚本超过引擎内存或栈限制。")));
                        return 6;
                    }
                    results.Add(new(null, Bounded(e.Message)));
                }
            }
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "invokeBatch", new BatchResult(results.ToArray())));
        }
        else if (request.Kind == "logs" && runtime is not null)
        {
            _ = WorkerProtocol.Payload<EmptyPayload>(request);
            var logs = runtime.GetLogs().Take(WorkerProtocol.MaxLogs).Select(Bounded).ToArray();
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(request.Identity, "logs", new WorkerLogs(logs)));
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
