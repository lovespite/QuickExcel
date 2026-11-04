using QuickExcel.Scripting;

using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>Trusted test-only in-process adapter. Production execution exclusively uses ScriptWorkerRuntimeFactory.</summary>
internal sealed class JscRuntimeFactory : IScriptRuntimeFactory
{
    public IScriptRuntime Create(string source) => new Runtime(new JscSession(source));
    public IScriptRuntime Create(ScriptRuntimeRequest request) => request.Modules.HasTypeScript
        ? throw new NotSupportedException("TypeScript 必须通过隔离编译和执行 worker。") : request.Modules.ScriptApiVersion == 3
        ? new Runtime(new JscSession(request.Modules.Modules.ToDictionary(m => m.Key, m => m.Value.Source, StringComparer.Ordinal), request.Modules.EntriesFor("formula-functions")))
        : request.Modules.CanExecuteSingleSource ? Create(request.Source) : throw new NotSupportedException("多模块执行需要显式声明脚本 API v2。");

    private sealed class Runtime(JscSession session) : IScriptRuntime
    {
        private readonly JscSession _session = session;
        public IReadOnlyList<string> FunctionNames => _session.FunctionNames;
        public IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors => _session.FunctionDescriptors
            .Select(d => new ScriptFunctionDescriptor(d.Name, d.ApiVersion, d.MinArguments, d.MaxArguments, d.ReturnType, d.Profile)).ToArray();
        public object? Invoke(string name, IReadOnlyList<object> arguments) => _session.Invoke(name, arguments);
        public IReadOnlyList<string> GetLogs() => _session.GetLogs();
        public void Dispose() => _session.Dispose();
    }
}

