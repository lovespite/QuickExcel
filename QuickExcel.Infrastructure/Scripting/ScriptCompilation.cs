using System.Collections.Immutable;
using System.Runtime.CompilerServices;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

public sealed class ScriptCompilationException(IReadOnlyList<CompilerDiagnostic> diagnostics)
    : InvalidOperationException("TypeScript 检查未通过；当前草稿未应用，也未执行。")
{
    public IReadOnlyList<CompilerDiagnostic> Diagnostics { get; } = diagnostics;
}

/// <summary>Derived artifacts belong to the immutable source set, never to workbook storage or signatures.</summary>
public static class ScriptCompilation
{
    private static readonly ConditionalWeakTable<ImmutableDictionary<string, ScriptModule>, ScriptCompiler> Compilers = new();

    public static async Task<CompiledReply?> PrepareAsync(ScriptModuleSet modules, CalculationStamp stamp, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!modules.HasTypeScript) return null;
        if (ScriptModuleSet.TypeScriptOptions != CompilerProtocol.Options) throw new InvalidOperationException("Compilation option identity mismatch.");
        var request = new CompileRequest(modules.Modules.Select(m => new CompilerModule(m.Key,
            m.Value.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript", m.Value.Source)).ToArray(), modules.EntryModules.ToArray(), "");
        request = request with { Sha256 = request.ComputeDigest() };
        var result = await Compilers.GetValue(modules.Modules, _ => new()).CompileAsync(request, stamp, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (!result.Result.Success) throw new ScriptCompilationException(Array.AsReadOnly(result.Result.Diagnostics));
        return result;
    }
}
