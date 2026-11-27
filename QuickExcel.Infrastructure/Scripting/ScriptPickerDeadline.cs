using System.Diagnostics;

namespace QuickExcel.Runtime;

/// <summary>Pause only while the trusted UI owns a native selection. Revocation remains live.</summary>
internal sealed class ScriptPickerDeadline(CancellationTokenSource action, ScriptWorkerRuntime runtime,
    Stopwatch elapsed, TimeSpan limit) : IDisposable
{
    private readonly object _gate = new();
    private bool _disposed;
    private sealed class Pause(ScriptPickerDeadline owner) : IDisposable
    {
        public void Dispose() => owner.Resume();
    }

    public IDisposable Begin()
    {
        lock (_gate)
        {
            if (_disposed) throw new OperationCanceledException();
            action.Token.ThrowIfCancellationRequested();
            runtime.RoundCancellation.ThrowIfCancellationRequested();
            if (!elapsed.IsRunning || elapsed.Elapsed >= limit) throw new OperationCanceledException();
            if (SharedScriptWorkers.SupportedPlatform)
            {
                elapsed.Stop();
                action.CancelAfter(Timeout.InfiniteTimeSpan);
                runtime.SetRoundTimeout(Timeout.InfiniteTimeSpan);
            }
            return new Pause(this);
        }
    }

    private void Resume()
    {
        lock (_gate)
        {
            if (_disposed || elapsed.IsRunning) return;
            var remaining = limit - elapsed.Elapsed;
            if (remaining < TimeSpan.Zero) remaining = TimeSpan.Zero;
            elapsed.Start();
            action.CancelAfter(remaining);
            runtime.SetRoundTimeout(remaining);
        }
    }

    public void Dispose() { lock (_gate) _disposed = true; }
}

