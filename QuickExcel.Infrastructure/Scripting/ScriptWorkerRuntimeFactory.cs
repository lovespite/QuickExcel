using System.ComponentModel;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using QuickExcel.Core;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

/// <summary>Creates a dedicated worker for one authorized round, with no fallback or automatic replay.</summary>
public sealed class ScriptWorkerRuntimeFactory(string? executablePath = null) : IScriptRuntimeFactory
{
    private readonly string _executable = Path.GetFullPath(executablePath ??
        Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"));
    public IScriptRuntime Create(string source) => Create(new ScriptRuntimeRequest(source, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default));
    public IScriptRuntime Create(ScriptRuntimeRequest request) => ScriptWorkerRuntime.Start(_executable, request);
}

public sealed class ScriptWorkerRuntime : IScriptRuntime
{
    private readonly ScriptRuntimeRequest _request;
    private readonly Process _process = new();
    private readonly object _processGate = new();
    private readonly SemaphoreSlim _exchangeGate = new(1, 1);
    private readonly CancellationTokenSource _deadline = new();
    private readonly CancellationTokenSource _lifetime;
    private CancellationTokenRegistration _registration;
    private WorkerJob? _job;
    private bool _started, _stopped, _disposed;
    private readonly RequestIdentity _identity;

    private ScriptWorkerRuntime(ScriptRuntimeRequest request)
    {
        _request = request;
        _identity = new(request.Stamp.SessionId, Guid.NewGuid(), Guid.NewGuid(), request.Stamp.InputRevision,
            request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration);
        _lifetime = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, _deadline.Token);
    }

    public int ProcessId { get; private set; }
    public Guid WorkerInstanceId => _identity.WorkerInstanceId;
    public IReadOnlyList<string> FunctionNames { get; private set; } = [];
    public IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors { get; private set; } = [];

    internal static ScriptWorkerRuntime Start(string executable, ScriptRuntimeRequest request)
    {
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.TimeLimit <= TimeSpan.Zero) throw new TimeoutException("重算已达到时间限制。");
        if (!OperatingSystem.IsWindows() || !Environment.Is64BitProcess) throw new PlatformNotSupportedException("Worker requires Windows x64.");
        if (!File.Exists(executable)) throw new ScriptRuntimeFailureException("WorkerMissing", "找不到脚本 worker，请重新构建或安装完整发布包。");
        if (!request.Modules.CanExecute) throw new NotSupportedException("多模块执行需要显式声明脚本 API v2。");
        if (request.Modules.Modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value.Source)) > 1024 * 1024)
            throw new ScriptRuntimeFailureException("SourceLimit", "脚本源码超过 1 MiB UTF-8 限额。");
        var runtime = new ScriptWorkerRuntime(request);
        try
        {
            runtime._identity.Validate();
            runtime._deadline.CancelAfter(request.TimeLimit < TimeSpan.FromSeconds(5) ? request.TimeLimit : TimeSpan.FromSeconds(5));
            runtime._process.StartInfo = new(executable)
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                WorkingDirectory = Path.GetDirectoryName(executable)!
            };
            // OS limits must be installed before sending any workbook source.
            if (!runtime._process.Start()) throw new IOException("Worker process did not start.");
            runtime._started = true;
            runtime.ProcessId = runtime._process.Id;
            runtime._job = new WorkerJob(runtime._process);
            runtime._registration = runtime._lifetime.Token.Register(runtime.Abort);
            _ = runtime.DrainStderrAsync();
            var hello = new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [1, 2], WorkerProtocol.MaxFrameBytes);
            var accepted = runtime.Exchange<WorkerHello, WorkerHello>("hello", hello);
            if (accepted.Engine != hello.Engine || accepted.Policy != hello.Policy ||
                !accepted.ApiVersions.SequenceEqual(hello.ApiVersions) || accepted.MaxFrameBytes != hello.MaxFrameBytes)
                throw new ProtocolException("Worker handshake is incompatible.");
            string digest;
            LoadedScript loaded;
            if (request.Modules.ScriptApiVersion == 2)
            {
                var bundle = new LoadModules(request.Modules.Modules.Select(m => new ModuleSource(m.Key, m.Value.Path, m.Value.Source)).ToArray(),
                    request.Modules.EntryModules.ToArray(), "");
                digest = bundle.ComputeDigest();
                loaded = runtime.Exchange<LoadModules, LoadedScript>("loadModules", bundle with { Sha256 = digest });
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
            runtime.Dispose();
            if (e is OperationCanceledException or TimeoutException or ScriptRuntimeFailureException) throw;
            throw new ScriptRuntimeFailureException("WorkerStart", "脚本 worker 启动或握手失败。", e);
        }
    }

    public object? Invoke(string name, IReadOnlyList<object> arguments)
    {
        var result = InvokeBatch([new FunctionCall(name, arguments.Select(WireValue.FromObject).ToArray())])[0];
        if (result.Error is not null) throw new InvalidOperationException(result.Error);
        return result.Value!.ToObject(scalarOnly: true);
    }

    public IReadOnlyList<InvocationResult> InvokeBatch(IReadOnlyList<FunctionCall> calls)
    {
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
                if (item is null || (item.Value is null) == (item.Error is null) ||
                    item.Error is not null && Encoding.UTF8.GetByteCount(item.Error) > WorkerProtocol.MaxLogBytes)
                    throw new ProtocolException("Invalid batch result.");
                item.Value?.Validate(scalarOnly: true);
            }
            return Array.AsReadOnly(result.Results);
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

    private TReply Exchange<TRequest, TReply>(string kind, TRequest payload)
    {
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
                if (error.Code is not ("Load" or "Timeout" or "ResourceLimit") || Encoding.UTF8.GetByteCount(error.Message) > 4000)
                    throw new ProtocolException("Invalid worker failure.");
                Abort();
                if (error.Code == "Timeout") throw new TimeoutException("脚本达到时间限制，本轮结果未提交。");
                throw new ScriptRuntimeFailureException(error.Code, error.Message.Length <= 4000 ? error.Message : "脚本 worker 执行失败。");
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
        finally { _exchangeGate.Release(); }
    }

    private async Task<WorkerEnvelope> RoundTripAsync(WorkerEnvelope request)
    {
        await WorkerProtocol.WriteAsync(_process.StandardInput.BaseStream, request, _lifetime.Token).ConfigureAwait(false);
        return await WorkerProtocol.ReadAsync(_process.StandardOutput.BaseStream, _lifetime.Token).ConfigureAwait(false);
    }

    private async Task DrainStderrAsync()
    {
        var buffer = new char[1024]; var total = 0;
        try
        {
            int count;
            while ((count = await _process.StandardError.ReadAsync(buffer, _lifetime.Token).ConfigureAwait(false)) > 0)
                if ((total += count) > 16 * 1024) { Abort(); return; }
        }
        catch (Exception e) when (e is IOException or OperationCanceledException or ObjectDisposedException or InvalidOperationException) { }
    }

    private void Abort()
    {
        lock (_processGate)
        {
            if (_stopped) return;
            _stopped = true;
            _job?.Dispose(); // KILL_ON_JOB_CLOSE also cleans up if the desktop process dies.
            if (_started)
                try { if (!_process.HasExited) _process.Kill(entireProcessTree: true); }
                catch (Exception e) when (e is Win32Exception or InvalidOperationException) { }
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
            if (_started) _process.WaitForExit(250);
            _process.Dispose();
            _lifetime.Dispose(); _deadline.Dispose();
        }
        finally { _exchangeGate.Release(); }
    }
}
