using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>Separate reusable API-v3 workers for formulas and actions in each authorized Windows session.
/// Linux retains its existing OS-enforced five-second process lifetime.</summary>
internal static class SharedScriptWorkers
{
    private static readonly object Gate = new();
    private static readonly Dictionary<(Guid, CancellationToken, bool Actions), Entry> Entries = [];
    private sealed class Entry(string executable)
    {
        internal readonly string Executable = executable;
        internal readonly SemaphoreSlim Busy = new(1, 1);
        internal ScriptWorkerRuntime? Runtime;
        internal CancellationTokenRegistration Registration;
    }

    internal static ScriptWorkerLease Acquire(string executable, ScriptRuntimeRequest request, bool actions)
    {
        if (!OperatingSystem.IsWindows() || request.Modules.ScriptApiVersion is not (3 or 4) || !request.RuntimeLifetime.CanBeCanceled)
        {
            var isolated = ScriptWorkerRuntime.Start(executable, request, actions);
            return new(isolated, isolated.Dispose);
        }
        request.RuntimeLifetime.ThrowIfCancellationRequested();
        var key = (request.Stamp.SessionId, request.RuntimeLifetime, actions);
        Entry entry;
        lock (Gate)
        {
            if (!Entries.TryGetValue(key, out entry!))
            {
                entry = new(executable); Entries.Add(key, entry);
                var captured = entry;
                entry.Registration = request.RuntimeLifetime.Register(() =>
                {
                    lock (Gate) Entries.Remove(key);
                    captured.Runtime?.Abort();
                    // Keep the ownership lease until verified exit and cleanup finish.
                    // Cancellation may occur while the document's own lock is held.
                    _ = Task.Run(() =>
                    {
                        captured.Busy.Wait();
                        try { captured.Runtime?.Dispose(); }
                        catch (Exception error) { System.Diagnostics.Trace.TraceError("Script worker cleanup: {0}", error.Message); }
                        finally { captured.Busy.Release(); captured.Registration.Dispose(); }
                    });
                });
            }
        }
        if (!string.Equals(entry.Executable, executable, StringComparison.OrdinalIgnoreCase))
            throw new ScriptRuntimeFailureException("WorkerStart", "工作簿不能在不同 worker 实现之间共享上下文。");
        if (!entry.Busy.Wait(0)) throw new ScriptRuntimeFailureException("Busy", "此工作簿的同类脚本正在执行。");
        try
        {
            request.RuntimeLifetime.ThrowIfCancellationRequested();
            if (entry.Runtime is { IsStopped: true })
            {
                if (request.Modules.HasInitialization) throw new ScriptRuntimeFailureException("WorkerStopped", "初始化上下文已失效，请重新应用脚本。");
                entry.Runtime.Dispose(); entry.Runtime = null;
            }
            if (entry.Runtime is null) entry.Runtime = ScriptWorkerRuntime.Start(executable, request, actions, shared: true);
            else entry.Runtime.BeginRound(request, actions);
            request.RuntimeLifetime.ThrowIfCancellationRequested();
            return new(entry.Runtime, () =>
            {
                try { entry.Runtime.EndRound(); }
                catch { entry.Runtime.Abort(); throw; }
                finally { entry.Busy.Release(); }
            });
        }
        catch { entry.Runtime?.Abort(); entry.Busy.Release(); throw; }
    }
}

internal sealed class ScriptWorkerLease(ScriptWorkerRuntime runtime, Action release) : IScriptBatchRuntime
{
    internal ScriptWorkerRuntime Runtime => runtime;
    public IReadOnlyList<string> FunctionNames => runtime.FunctionNames;
    public IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors => runtime.FunctionDescriptors;
    public object? Invoke(string name, IReadOnlyList<object> arguments) => runtime.Invoke(name, arguments);
    public IScriptInvocationBatch CreateBatch() => runtime.CreateBatch();
    public IReadOnlyList<string> GetLogs() => runtime.GetLogs();
    private Action? _release = release;
    public void Dispose() => Interlocked.Exchange(ref _release, null)?.Invoke();
}
