using QuickExcel.Scripting;

using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>Creates an isolated native context for one authorized calculation. No runtime is reused.</summary>
public sealed class QuickJsRuntimeFactory : IScriptRuntimeFactory
{
    public IScriptRuntime Create(string source) => new Runtime(new QuickJsSession(source));
    public IScriptRuntime Create(ScriptRuntimeRequest request) => request.Modules.ScriptApiVersion == 2
        ? new Runtime(new QuickJsSession(request.Modules.Modules.ToDictionary(m => m.Key, m => m.Value.Source, StringComparer.Ordinal), request.Modules.EntryModules))
        : request.Modules.CanExecuteLegacy ? Create(request.Source) : throw new NotSupportedException("多模块执行需要显式声明脚本 API v2。");

    private sealed class Runtime(QuickJsSession session) : IScriptRuntime
    {
        private readonly QuickJsSession _session = session;
        public IReadOnlyList<string> FunctionNames => _session.FunctionNames;
        public IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors => _session.FunctionDescriptors
            .Select(d => new ScriptFunctionDescriptor(d.Name, d.ApiVersion, d.MinArguments, d.MaxArguments, d.ReturnType, d.Profile)).ToArray();
        public object? Invoke(string name, IReadOnlyList<object> arguments) => _session.Invoke(name, arguments);
        public IReadOnlyList<string> GetLogs() => _session.GetLogs();
        public void Dispose() => _session.Dispose();
    }
}

