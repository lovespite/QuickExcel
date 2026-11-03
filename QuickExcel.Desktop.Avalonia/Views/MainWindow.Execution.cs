namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private ExecutionOverlay _execution = null!;
    private ExecutionOverlay? _scriptExecution;
    private bool NonExecutionBusy => _workspace.IsBusy && !_execution.IsRunning;

    private IDisposable BeginExecution(string title, CancellationTokenSource cancellation)
    {
        var main = _execution.Begin(title, cancellation);
        try { return new ExecutionScope(main, _scriptExecution?.Begin(title, cancellation)); }
        catch { main.Dispose(); throw; }
    }

    private sealed class ExecutionScope(IDisposable main, IDisposable? scripts) : IDisposable
    {
        public void Dispose() { scripts?.Dispose(); main.Dispose(); }
    }
}
