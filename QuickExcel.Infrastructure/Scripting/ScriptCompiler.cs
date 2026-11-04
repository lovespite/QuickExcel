using System.ComponentModel;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

/// <summary>One isolated compiler globally, zero queued requests. Each instance owns a bounded session cache.</summary>
public sealed class ScriptCompiler(string? executablePath = null)
{
    private static readonly SemaphoreSlim Capacity = new(1, 1);
    private readonly string _executable = Path.GetFullPath(executablePath ?? Path.Combine(AppContext.BaseDirectory, "ScriptWorker", OperatingSystem.IsWindows() ? "QuickExcel.ScriptWorker.exe" : "QuickExcel.ScriptWorker"));
    private readonly object _cacheGate = new();
    // Cache serialized immutable derived data so callers cannot mutate an accepted entry.
    private readonly Dictionary<string, byte[]> _cache = new(StringComparer.Ordinal);
    private int _cacheBytes;
    public int ProcessLaunchCount { get; private set; }

    public async Task<CompiledReply> CompileAsync(CompileRequest input, CalculationStamp stamp, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // Snapshot caller-owned arrays before validating, hashing or starting asynchronous work.
        var request = input with { Modules = input.Modules.ToArray(), Entries = input.Entries.ToArray() };
        request.Validate();
        lock (_cacheGate)
        {
            if (_cache.TryGetValue(request.Sha256, out var cached))
            {
                var hit = JsonSerializer.Deserialize<CompiledReply>(cached, WorkerProtocol.Json)!;
                CompilerProtocol.ValidateReply(hit, request);
                cancellationToken.ThrowIfCancellationRequested();
                return hit;
            }
        }
        if (!await Capacity.WaitAsync(0, cancellationToken).ConfigureAwait(false))
            throw new ScriptRuntimeFailureException("CompilerBusy", "已有 TypeScript 编译正在进行，请稍后重试。");
        try
        {
            if (!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux()) throw new PlatformNotSupportedException("此平台尚未验证隔离 TypeScript 编译角色。");
            if (!File.Exists(_executable)) throw new ScriptRuntimeFailureException("WorkerMissing", "找不到完整的脚本 worker。");
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(6));
            using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
            IWorkerProcess? process = null;
            try
            {
                process = await Task.Run<IWorkerProcess>(() => OperatingSystem.IsWindows() ? WorkerProcess.Start(_executable, lifetime.Token, compilation: true)
                    : OperatingSystem.IsLinux() ? LinuxWorkerProcess.Start(_executable, lifetime.Token, compilation: true)
                    : throw new PlatformNotSupportedException(), lifetime.Token).ConfigureAwait(false);
                ProcessLaunchCount++;
                using var stop = lifetime.Token.Register(process.Kill);
                var stderr = DrainError(process, lifetime.Token);
                var identity = new RequestIdentity(stamp.SessionId, Guid.NewGuid(), Guid.NewGuid(), stamp.InputRevision, null, stamp.ScriptGeneration);
                identity.Validate();
                var hello = CompilerProtocol.Hello;
                await WorkerProtocol.WriteAsync(process.Input, WorkerProtocol.Message(identity, "hello", hello), lifetime.Token).ConfigureAwait(false);
                var response = await WorkerProtocol.ReadAsync(process.Output, lifetime.Token).ConfigureAwait(false);
                WorkerProtocol.ValidateReply(response, identity, "hello");
                var accepted = WorkerProtocol.Payload<WorkerHello>(response);
                if (accepted.Engine != hello.Engine || accepted.Policy != hello.Policy || accepted.Role != "compile" ||
                    accepted.Compiler != hello.Compiler || !accepted.ApiVersions.SequenceEqual([3]) || accepted.MaxFrameBytes != hello.MaxFrameBytes)
                    throw new ProtocolException("Compiler handshake mismatch; no source was sent.");
                if (OperatingSystem.IsLinux() && process is LinuxWorkerProcess linux) linux.VerifyBeforeSource();
                identity = identity with { RequestId = Guid.NewGuid() };
                await WorkerProtocol.WriteAsync(process.Input, WorkerProtocol.Message(identity, "compile", request), lifetime.Token).ConfigureAwait(false);
                response = await WorkerProtocol.ReadAsync(process.Output, lifetime.Token).ConfigureAwait(false);
                WorkerProtocol.ValidateReply(response, identity, "compile");
                var reply = WorkerProtocol.Payload<CompiledReply>(response);
                CompilerProtocol.ValidateReply(reply, request);
                lifetime.Token.ThrowIfCancellationRequested();
                if (reply.Result.Success)
                {
                    var bytes = JsonSerializer.SerializeToUtf8Bytes(reply, WorkerProtocol.Json);
                    lock (_cacheGate)
                    {
                        lifetime.Token.ThrowIfCancellationRequested();
                        if (_cache.Count >= 4 || _cacheBytes + bytes.Length > 16 * 1024 * 1024) { _cache.Clear(); _cacheBytes = 0; }
                        _cache[request.Sha256] = bytes;
                        _cacheBytes += bytes.Length;
                    }
                }
                return reply;
            }
            catch (Exception e) when (e is IOException or Win32Exception or OperationCanceledException or JsonException or InvalidOperationException or KeyNotFoundException or FormatException)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (deadline.IsCancellationRequested) throw new TimeoutException("TypeScript 编译达到独立时间限制，未接受任何产物。", e);
                throw new ScriptRuntimeFailureException("Compile", "TypeScript 编译 worker 失败或产物校验失败。", e);
            }
            finally { process?.Dispose(); }
        }
        finally { Capacity.Release(); }
    }

    public void ClearCache() { lock (_cacheGate) { _cache.Clear(); _cacheBytes = 0; } }

    private static async Task DrainError(IWorkerProcess process, CancellationToken token)
    {
        var buffer = new byte[1024]; var total = 0;
        try
        {
            int count;
            while ((count = await process.Error.ReadAsync(buffer, token).ConfigureAwait(false)) != 0)
                if ((total += count) > 16384) { process.Kill(); return; }
        }
        catch (Exception e) when (e is IOException or OperationCanceledException or ObjectDisposedException) { }
    }
}
