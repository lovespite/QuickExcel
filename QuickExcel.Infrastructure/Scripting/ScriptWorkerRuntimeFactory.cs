using System.ComponentModel;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using QuickExcel.Core;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

/// <summary>Leases the isolated formula worker on Windows API v2; actions use a separate worker and context.
/// No in-process fallback or automatic replay.</summary>
public sealed class ScriptWorkerRuntimeFactory(string? executablePath = null) : IScriptRuntimeFactory
{
    private readonly string _executable = Path.GetFullPath(executablePath ??
        Path.Combine(AppContext.BaseDirectory, "ScriptWorker", OperatingSystem.IsWindows() ? "QuickExcel.ScriptWorker.exe" : "QuickExcel.ScriptWorker"));
    public IScriptRuntime Create(string source) => Create(new ScriptRuntimeRequest(source, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default));
    public IScriptRuntime Create(ScriptRuntimeRequest request) => OperatingSystem.IsWindows() && request.Modules.ScriptApiVersion is 3 or 4 && request.RuntimeLifetime.CanBeCanceled
        ? SharedScriptWorkers.Acquire(_executable, request, actions: false) : ScriptWorkerRuntime.Start(_executable, request);

    /// <summary>Reclaims abandoned rounds for the current user, without executing scripts.</summary>
    public static int RecoverAbandonedSandboxes(CancellationToken cancellation = default)
    {
        if (OperatingSystem.IsLinux()) return LinuxWorkerLease.Recover(cancellation);
        if (OperatingSystem.IsMacOS()) return MacWorkerLease.Recover(cancellation);
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("此平台尚未开放脚本隔离。");
        try { return WorkerSandboxLease.Recover(cancellation); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or Win32Exception)
        { throw new ScriptRuntimeFailureException("WorkerCleanup", "脚本 worker 残留资源回收失败，将在下次启动时重试。", e); }
    }

    internal static bool SupportedArchitecture() => OperatingSystem.IsMacOS()
        ? System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture == System.Runtime.InteropServices.Architecture.Arm64
        : System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture == System.Runtime.InteropServices.Architecture.X64;
}

public sealed class ScriptWorkerRuntime : IScriptBatchRuntime
{
    private ScriptRuntimeRequest _request;
    private IWorkerProcess? _process;
    private readonly object _processGate = new();
    private readonly SemaphoreSlim _exchangeGate = new(1, 1);
    private readonly CancellationTokenSource _deadline = new();
    private CancellationTokenSource _lifetime;
    private readonly CancellationTokenSource _processLifetime = CancellationTokenSource.CreateLinkedTokenSource(ScriptExecutionPolicy.Current.CaptureLifetime());
    private CancellationTokenRegistration _registration;
    private CancellationTokenRegistration _policyRegistration;
    private bool _stopped, _disposed;
    private RequestIdentity _identity;
    private bool _actionsMode;
    private readonly bool _shared;
    internal bool IsStopped => _stopped || _disposed;
    private readonly Dictionary<string, LogicalSourceMap> _sourceMaps = new(StringComparer.Ordinal);
    private bool _hasCompiledSources;
    private string OriginalDiagnostic(string message) => System.Text.RegularExpressions.Regex.Replace(message,
        @"(?<=\n at )([A-Za-z0-9_.-]{1,128}):([1-9][0-9]{0,8}):([1-9][0-9]{0,8})", match =>
        {
            if (!_sourceMaps.TryGetValue(match.Groups[1].Value, out var map))
                return _hasCompiledSources ? match.Value + MessageText.Get(" (生成 JS 位置)") : match.Value;
            var position = map.OriginalPosition(int.Parse(match.Groups[2].Value), int.Parse(match.Groups[3].Value));
            return position is { } p ? $"{match.Groups[1].Value}:{p.Line}:{p.Column}" : match.Value + MessageText.Get(" (生成 JS 位置)");
        });

    private ScriptWorkerRuntime(ScriptRuntimeRequest request, bool actionsMode, bool shared)
    {
        _request = request;
        _actionsMode = actionsMode;
        _shared = shared;
        _identity = new(request.Stamp.SessionId, Guid.NewGuid(), Guid.NewGuid(), request.Stamp.InputRevision,
            actionsMode ? null : request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration);
        _lifetime = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, _deadline.Token, _processLifetime.Token);
    }

    public int ProcessId { get; private set; }
    internal string SandboxProfileName => _process!.ProfileName;
    internal string SandboxPayloadDirectory => _process!.PayloadDirectory;
    public Guid WorkerInstanceId => _identity.WorkerInstanceId;
    public IReadOnlyList<string> FunctionNames { get; private set; } = [];
    public IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors { get; private set; } = [];
    public IReadOnlyList<ScriptActionDescriptor> ActionDescriptors { get; private set; } = [];
    public int InvocationBatchCount { get; private set; }
    public int InvocationCount { get; private set; }

    internal static ScriptWorkerRuntime Start(string executable, ScriptRuntimeRequest request, bool actionsMode = false, bool shared = false)
    {
        ScriptExecutionPolicy.Current.RequireEnabled();
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.TimeLimit <= TimeSpan.Zero) throw new TimeoutException("重算已达到时间限制。");
        if (!ScriptWorkerRuntimeFactory.SupportedArchitecture() || (!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux() && !OperatingSystem.IsMacOS()))
            throw new PlatformNotSupportedException("此平台尚未开放脚本隔离。");
        if (!File.Exists(executable)) throw new ScriptRuntimeFailureException("WorkerMissing", "找不到脚本 worker，请重新构建或安装完整发布包。");
        if (!request.Modules.CanExecute) throw new NotSupportedException(MessageText.Get(QuickExcel.Domain.ScriptModuleSet.MigrationMessage));
        if (request.Modules.Modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value.Source)) > 1024 * 1024)
            throw new ScriptRuntimeFailureException("SourceLimit", "脚本源码超过 1 MiB UTF-8 限额。");
        if (actionsMode && !request.Modules.HasActions) throw new InvalidOperationException("没有 actions 绑定。");
        if (request.Modules.HasInitialization && !OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("初始化主入口需要 Windows 持久隔离上下文；此平台尚未支持。");
        var compiled = ScriptCompilation.PrepareAsync(request.Modules, request.Stamp, request.CancellationToken).GetAwaiter().GetResult();
        var executableSources = compiled?.Result.Outputs.ToDictionary(o => o.Id, o => o.Javascript, StringComparer.Ordinal);
        var runtime = new ScriptWorkerRuntime(request, actionsMode, shared);
        runtime._hasCompiledSources = compiled is not null;
        if (compiled is not null)
            foreach (var item in compiled.Result.Outputs)
            {
                if (item.SourceMap.Length == 0) continue;
                using var map = System.Text.Json.JsonDocument.Parse(item.SourceMap);
                runtime._sourceMaps.Add(item.Id, new LogicalSourceMap(map.RootElement.GetProperty("mappings").GetString()!,
                    request.Modules.Modules[item.Id].Source, map.RootElement.GetProperty("names").GetArrayLength()));
            }
        try
        {
            runtime._identity.Validate();
            runtime._deadline.CancelAfter(request.TimeLimit < TimeSpan.FromSeconds(5) ? request.TimeLimit : TimeSpan.FromSeconds(5));
            runtime._process = OperatingSystem.IsWindows() ? WorkerProcess.Start(executable, runtime._lifetime.Token) : OperatingSystem.IsLinux()
                ? LinuxWorkerProcess.Start(executable, runtime._lifetime.Token) : OperatingSystem.IsMacOS()
                ? MacWorkerProcess.Start(executable, runtime._lifetime.Token) : throw new PlatformNotSupportedException();
            runtime.ProcessId = runtime._process.Id;
            // This registration remains active between rounds, when the per-round registration is released.
            runtime._policyRegistration = runtime._processLifetime.Token.Register(runtime.Abort);
            runtime._registration = runtime._lifetime.Token.Register(runtime.Abort);
            _ = runtime.DrainStderrAsync();
            var hello = new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [3, 4], WorkerProtocol.MaxFrameBytes);
            var accepted = runtime.Exchange<WorkerHello, WorkerHello>("hello", hello);
            if (accepted.Engine != hello.Engine || accepted.Policy != hello.Policy ||
                accepted.Role != "execute" || accepted.Compiler is not null ||
                !accepted.ApiVersions.SequenceEqual(hello.ApiVersions) || accepted.MaxFrameBytes != hello.MaxFrameBytes)
                throw new ProtocolException("Worker handshake is incompatible.");
            if (OperatingSystem.IsLinux() && runtime._process is LinuxWorkerProcess linux) linux.VerifyBeforeSource();
            if (OperatingSystem.IsMacOS() && runtime._process is MacWorkerProcess mac) mac.VerifyBeforeSource();
            string digest;
            LoadedScript loaded;
            if (request.Modules.ScriptApiVersion is 3 or 4)
            {
                // JscSession appends the main(workbook, host) registration to every
                // entry module; no QuickExcel global wrapper is generated here.
                var permitted = request.Modules.Modules.Where(m => request.Modules.ScriptApiVersion != 4 ||
                    !request.Modules.Bindings.Any(b => b.Module == m.Key &&
                        (actionsMode ? b.Kind == "formula-functions" : b.Kind is "actions" or "events" or "initialization")));
                var bundle = new LoadModules(permitted.Select(m => new ModuleSource(m.Key, m.Value.Path,
                    executableSources is null ? m.Value.Source : executableSources[m.Key], m.Value.Description
                    )).ToArray(),
                    (actionsMode ? request.Modules.ActionContextEntries : request.Modules.EntriesFor("formula-functions")).ToArray(), "")
                { FunctionEntries = shared && !actionsMode ? request.Modules.EntriesFor("formula-functions").ToArray() : [],
                    ActionEntries = (shared || request.Modules.ScriptApiVersion == 4) && actionsMode ? request.Modules.EntriesFor("actions").ToArray() : [],
                    EventEntries = actionsMode ? request.Modules.EntriesFor("events").ToArray() : [],
                    InitializationEntries = actionsMode ? request.Modules.EntriesFor("initialization").ToArray() : [],
                    ApiVersion = request.Modules.ScriptApiVersion };
                digest = bundle.ComputeDigest();
                if (shared)
                {
                    var workbook = runtime.Exchange<LoadModules, LoadedWorkbook>("loadWorkbook", bundle with { Sha256 = digest });
                    if (actionsMode ? workbook.Functions.Length != 0 : workbook.Actions.Length != 0)
                        throw new ProtocolException("Worker registered capabilities outside its role.");
                    ValidateActions(workbook.Sha256, workbook.Actions);
                    loaded = new(workbook.Sha256, workbook.Functions);
                }
                else if (actionsMode)
                {
                    var actions = runtime.Exchange<LoadModules, LoadedActions>("loadActions", bundle with { Sha256 = digest });
                    ValidateActions(actions.Sha256, actions.Actions);
                    return runtime;
                }
                else loaded = runtime.Exchange<LoadModules, LoadedScript>("loadModules", bundle with { Sha256 = digest });
                void ValidateActions(string sha256, ActionInfo[] actions)
                {
                    if (sha256 != digest || actions.Length > 256) throw new ProtocolException("Invalid action registry.");
                    var names = new HashSet<string>(StringComparer.Ordinal);
                    foreach (var action in actions)
                    {
                        if (action is null || !names.Add(action.Name) || action.Description is null || action.Description.Length > 1000) throw new ProtocolException("Invalid action descriptor.");
                        ActionProtocol.Name(action.Name);
                    }
                    runtime.ActionDescriptors = Array.AsReadOnly(actions.Select(a => new ScriptActionDescriptor(a.Name, a.Description)).ToArray());
                }
            }
            else
            {
                digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(request.Source)));
                loaded = runtime.Exchange<LoadScript, LoadedScript>("load", new(request.Source, digest));
            }
            if (loaded.Sha256 != digest || loaded.Functions.Length > WorkerProtocol.MaxFunctions)
                throw new ProtocolException("Worker loaded a different script or an invalid function table.");
            var descriptors = loaded.Functions.Select(d => d is null ? throw new ProtocolException("Missing function.") :
                new ScriptFunctionDescriptor(d.Name, d.ApiVersion, d.MinArguments, d.MaxArguments, d.ReturnType, d.Profile)).ToArray();
            foreach (var descriptor in descriptors) descriptor.Validate();
            if (descriptors.Select(d => d.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != descriptors.Length)
                throw new ProtocolException("Duplicate function registration.");
            runtime.FunctionDescriptors = Array.AsReadOnly(descriptors);
            runtime.FunctionNames = Array.AsReadOnly(descriptors.Select(d => d.Name).ToArray());
            return runtime;
        }
        catch (Exception e)
        {
            var timedOut = runtime._deadline.IsCancellationRequested;
            runtime.Dispose();
            if (e is OperationCanceledException)
            {
                request.CancellationToken.ThrowIfCancellationRequested();
                if (timedOut) throw new TimeoutException("脚本 worker 启动超过本轮时间上限。", e);
            }
            if (e is OperationCanceledException or TimeoutException or ScriptRuntimeFailureException) throw;
            throw new ScriptRuntimeFailureException("WorkerStart", "脚本 worker 启动或握手失败。", e);
        }
    }

    internal void BeginRound(ScriptRuntimeRequest request, bool actions)
    {
        if (!_shared || IsStopped || actions != _actionsMode || request.Stamp.SessionId != _identity.SessionId || request.Stamp.ScriptGeneration != _identity.ScriptGeneration ||
            !ReferenceEquals(request.Modules, _request.Modules)) throw new InvalidOperationException("工作簿脚本上下文已失效。");
        request.CancellationToken.ThrowIfCancellationRequested(); request.RuntimeLifetime.ThrowIfCancellationRequested();
        if (request.TimeLimit <= TimeSpan.Zero) throw new TimeoutException("脚本已达到本轮时间限制。");
        _request = request; _actionsMode = actions;
        _identity = _identity with { RequestId = Guid.NewGuid(), RoundId = Guid.NewGuid(), Revision = request.Stamp.InputRevision,
            CalculationGeneration = actions ? null : request.Stamp.CalculationGeneration };
        _lifetime = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, _deadline.Token, _processLifetime.Token);
        _deadline.CancelAfter(request.TimeLimit < TimeSpan.FromSeconds(5) ? request.TimeLimit : TimeSpan.FromSeconds(5));
        _registration = _lifetime.Token.Register(Abort);
        InvocationBatchCount = InvocationCount = 0;
        _ = Exchange<BeginWorkerRound, EmptyPayload>("beginRound", new());
    }

    internal void EndRound()
    {
        if (!_shared || IsStopped) return;
        _ = Exchange<EmptyPayload, EmptyPayload>("endRound", new());
        _registration.Dispose(); _deadline.CancelAfter(Timeout.InfiniteTimeSpan); _lifetime.Dispose();
    }

    public object? Invoke(string name, IReadOnlyList<object> arguments)
    {
        var result = InvokeBatch([new FunctionCall(name, arguments.Select(WireValue.FromObject).ToArray())])[0];
        if (result.ValueTooLong) throw new FormatException("计算结果超过单元格文本长度限制。");
        if (result.Error is not null) throw new InvalidOperationException(result.Error);
        return result.Value!.ToObject(scalarOnly: true);
    }

    public IScriptInvocationBatch CreateBatch() => new InvocationBatch(this);

    private sealed class InvocationBatch(ScriptWorkerRuntime owner) : IScriptInvocationBatch
    {
        private readonly List<FunctionCall> _calls = [];
        // Covers the fixed envelope, identity (including Int64 revisions), kind and array syntax.
        private const int EnvelopeReserve = 1024;
        private int _bytes = EnvelopeReserve;
        private bool _invoked;
        public int Count => _calls.Count;

        public bool TryAdd(string name, IReadOnlyList<object> arguments)
        {
            if (_invoked) throw new InvalidOperationException("A batch cannot be reused.");
            owner._lifetime.Token.ThrowIfCancellationRequested();
            if (_calls.Count == WorkerProtocol.MaxCalculationBatchCalls) return false;
            FunctionCall call;
            int size;
            try
            {
                call = new(name, arguments.Select(WireValue.FromObject).ToArray());
                size = Encoding.UTF8.GetByteCount(WorkerProtocol.EncodePayload(call).GetRawText()) + 1;
            }
            catch (ProtocolException e)
            { throw new ScriptRuntimeFailureException("InputLimit", "单次脚本调用无法在 IPC 输入限额内编码，本轮未提交。", e); }
            if (size > WorkerProtocol.MaxFrameBytes - EnvelopeReserve)
                throw new ScriptRuntimeFailureException("InputLimit", "单次脚本调用超过 IPC 输入限额，本轮未提交。");
            if (_bytes + size > WorkerProtocol.MaxFrameBytes) return false;
            _calls.Add(call); _bytes += size;
            return true;
        }

        public IReadOnlyList<ScriptCallResult> Invoke()
        {
            if (_invoked) throw new InvalidOperationException("A batch cannot be reused.");
            _invoked = true;
            return owner.InvokeBatch(_calls).Select(r => new ScriptCallResult(r.Value?.ToObject(scalarOnly: true), r.Error, r.ValueTooLong)).ToArray();
        }
    }

    public IReadOnlyList<InvocationResult> InvokeBatch(IReadOnlyList<FunctionCall> calls)
    {
        if (_actionsMode) throw new InvalidOperationException("Action workers cannot invoke formula functions.");
        if (calls.Count is < 1 or > WorkerProtocol.MaxBatchCalls) throw new ArgumentException("Invalid batch size.");
        foreach (var call in calls)
        {
            if (call is null || call.Name.Length is < 1 or > 256 || call.Arguments.Length > 255) throw new ArgumentException("Invalid function call.");
            foreach (var argument in call.Arguments) argument.Validate();
        }
        var result = Exchange<InvokeBatch, BatchResult>("invokeBatch", new(calls.ToArray()));
        try
        {
            if (result.Results.Length != calls.Count) throw new ProtocolException("Incomplete batch result.");
            foreach (var item in result.Results)
            {
                if (item is null || (item.ValueTooLong ? item.Value is not null || item.Error is not null : (item.Value is null) == (item.Error is null)) ||
                    item.Error is not null && Encoding.UTF8.GetByteCount(item.Error) > WorkerProtocol.MaxLogBytes)
                    throw new ProtocolException("Invalid batch result.");
                item.Value?.Validate(scalarOnly: true);
            }
            InvocationBatchCount++;
            InvocationCount += calls.Count;
            return Array.AsReadOnly(result.Results.Select(item => item.Error is null ? item : item with { Error = OriginalDiagnostic(item.Error) }).ToArray());
        }
        catch (ProtocolException e) { Abort(); throw new ScriptRuntimeFailureException("Protocol", "脚本 worker 返回无效结果，本轮未提交。", e); }
    }

    public IReadOnlyList<string> GetLogs()
    {
        var logs = Exchange<EmptyPayload, WorkerLogs>("logs", new()).Messages;
        if (logs.Length > WorkerProtocol.MaxLogs || logs.Any(l => l is null || Encoding.UTF8.GetByteCount(l) > WorkerProtocol.MaxLogBytes))
        { Abort(); throw new ScriptRuntimeFailureException("Protocol", "脚本 worker 日志超过限额。"); }
        return Array.AsReadOnly(logs);
    }

    private Func<ModelCall, ModelReply>? _model;
    private Guid _modelOperation;
    public ActionPatch RunAction(RunAction request, Func<ModelCall, ModelReply>? model = null)
    {
        if (!_actionsMode) throw new InvalidOperationException("Formula workers cannot run actions.");
        ActionProtocol.Name(request.Name);
        ActionProtocol.ValidateSnapshot(request.Snapshot);
        if (request.OperationId == Guid.Empty || request.Snapshot.BaseRevision != _request.Stamp.InputRevision) throw new ProtocolException("Action input identity mismatch.");
        _model = model; _modelOperation = request.OperationId;
        try { var patch = Exchange<RunAction, ActionPatch>("runAction", request); ActionProtocol.ValidatePatch(patch, request); return patch; }
        catch (ProtocolException e) { Abort(); throw new ScriptRuntimeFailureException("Protocol", "动作变更校验失败，本次未提交。", e); }
        finally { _model = null; _modelOperation = Guid.Empty; }
    }

    private TReply Exchange<TRequest, TReply>(string kind, TRequest payload)
    {
        ScriptExecutionPolicy.Current.RequireEnabled();
        // One in-flight request, zero waiting requests. Reject overload instead of growing a queue.
        if (!_exchangeGate.Wait(0)) throw new ScriptRuntimeFailureException("Busy", "脚本 worker 已有请求正在执行。");
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _lifetime.Token.ThrowIfCancellationRequested();
            if (_stopped) throw new ScriptRuntimeFailureException("WorkerStopped", "脚本 worker 已停止；请显式重新计算。");
            var identity = _identity with { RequestId = Guid.NewGuid() };
            var reply = RoundTripAsync(WorkerProtocol.Message(identity, kind, payload)).GetAwaiter().GetResult();
            WorkerProtocol.ValidateReply(reply, identity, kind);
            if (reply.Kind == "failure")
            {
                var error = WorkerProtocol.Payload<WorkerFailure>(reply);
                if (error.Code is not ("Load" or "Timeout" or "ResourceLimit" or "Action"
                    or "Invalidated"
                    ) || Encoding.UTF8.GetByteCount(error.Message) > 4000 || error.Logs is null || error.Logs.Length > WorkerProtocol.MaxLogs ||
                    error.Logs.Any(log => log is null || Encoding.UTF8.GetByteCount(log) > WorkerProtocol.MaxLogBytes))
                    throw new ProtocolException("Invalid worker failure.");
                Abort();
                if (error.Code == "Timeout") throw new TimeoutException("脚本达到时间限制，本轮结果未提交。");
                throw new ScriptRuntimeFailureException(error.Code, OriginalDiagnostic(error.Message.Length <= 4000 ? error.Message : "脚本 worker 执行失败。"))
                { Logs = Array.AsReadOnly(error.Logs) };
            }
            return WorkerProtocol.Payload<TReply>(reply);
        }
        catch (Exception e) when (e is OperationCanceledException or IOException or Win32Exception)
        {
            Abort();
            _request.CancellationToken.ThrowIfCancellationRequested();
            if (_deadline.IsCancellationRequested) throw new TimeoutException("脚本 worker 达到本轮时间上限，已终止；结果未提交。", e);
            throw new ScriptRuntimeFailureException(e is ProtocolException ? "Protocol" : "WorkerExit", "脚本 worker 退出或通信中断，本轮结果未提交。", e);
        }
        catch
        {
            // A host model callback can fail while the worker is waiting for its reply
            // (for example, a provisional UDF calculation timed out). That exchange
            // cannot be resumed or followed by endRound; preserve the original failure.
            Abort();
            throw;
        }
        finally { _exchangeGate.Release(); }
    }

    private async Task<WorkerEnvelope> RoundTripAsync(WorkerEnvelope request)
    {
        await WorkerProtocol.WriteAsync(_process!.Input, request, _lifetime.Token).ConfigureAwait(false);
        var sequence = 0; long bytes = 0;
        while (true)
        {
            var reply = await WorkerProtocol.ReadAsync(_process.Output, _lifetime.Token).ConfigureAwait(false);
            if (reply.Kind != "model") return reply;
            WorkerProtocol.ValidateReply(reply, request.Identity, "model");
            if (request.Kind != "runAction" || _model is null || ++sequence > 4096) throw new ProtocolException("Unexpected model request.");
            var call = WorkerProtocol.Payload<ModelCall>(reply);
            if (call.OperationId != _modelOperation || call.Sequence != sequence) throw new ProtocolException("Model request identity mismatch.");
            bytes += Encoding.UTF8.GetByteCount(reply.Payload.GetRawText());
            if (bytes > WorkerProtocol.MaxFrameBytes) throw new ProtocolException("Model round exceeds its byte budget.");
            var response = WorkerProtocol.Message(request.Identity, "model", _model(call));
            bytes += Encoding.UTF8.GetByteCount(response.Payload.GetRawText());
            if (bytes > WorkerProtocol.MaxFrameBytes) throw new ProtocolException("Model round exceeds its byte budget.");
            await WorkerProtocol.WriteAsync(_process.Input, response, _lifetime.Token).ConfigureAwait(false);
        }
    }

    private async Task DrainStderrAsync()
    {
        var buffer = new byte[1024]; var total = 0;
        try
        {
            int count;
            while ((count = await _process!.Error.ReadAsync(buffer, _processLifetime.Token).ConfigureAwait(false)) > 0)
                if ((total += count) > 16 * 1024) { Abort(); return; }
        }
        catch (Exception e) when (e is IOException or OperationCanceledException or ObjectDisposedException or InvalidOperationException) { }
    }

    internal void Abort()
    {
        lock (_processGate)
        {
            if (_stopped) return;
            _stopped = true;
            _processLifetime.Cancel();
            _process?.Kill(); // KILL_ON_JOB_CLOSE also cleans up if the desktop process dies.
        }
    }

    public void Dispose()
    {
        Abort();
        _exchangeGate.Wait();
        try
        {
            if (_disposed) return;
            _disposed = true;
            _registration.Dispose();
            _policyRegistration.Dispose();
            try { _process?.Dispose(); }
            catch (IOException e) { throw new ScriptRuntimeFailureException("WorkerCleanup", "脚本 worker 隔离资源清理失败。", e); }
            finally { _lifetime.Dispose(); _deadline.Dispose(); _processLifetime.Dispose(); }
        }
        finally { _exchangeGate.Release(); }
    }
}
