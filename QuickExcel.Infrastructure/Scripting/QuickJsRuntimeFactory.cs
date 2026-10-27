using QuickExcel.Scripting;

using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>Creates an isolated native context for one authorized calculation. No runtime is reused.</summary>
public sealed class QuickJsRuntimeFactory : IScriptRuntimeFactory
{
    public IScriptRuntime Create(string source) => new Runtime(source);

    private sealed class Runtime(string source) : IScriptRuntime
    {
        private readonly QuickJsSession _session = new(source);
        public IReadOnlyList<string> FunctionNames => _session.FunctionNames;
        public object? Invoke(string name, IReadOnlyList<object> arguments) => _session.Invoke(name, arguments);
        public IReadOnlyList<string> GetLogs() => _session.GetLogs();
        public void Dispose() => _session.Dispose();
    }
}

