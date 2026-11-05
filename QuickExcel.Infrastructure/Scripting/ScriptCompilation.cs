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

/// <summary>Compiles immutable sources or reuses authenticated package artifacts. Never grants execution consent.</summary>
public static class ScriptCompilation
{
    private static readonly ConditionalWeakTable<ImmutableDictionary<string, ScriptModule>, ScriptCompiler> Compilers = new();
    private static string _sourceMapMode = "full";
    /// <summary>Host-local preference. Captured in each request and its cache digest; never source semantics or consent.</summary>
    public static string SourceMapMode
    {
        get => Volatile.Read(ref _sourceMapMode);
        set
        {
            if (value is not ("none" or "compact" or "full")) throw new ArgumentOutOfRangeException(nameof(value));
            Volatile.Write(ref _sourceMapMode, value);
        }
    }

    public static async Task<CompiledReply?> PrepareAsync(ScriptModuleSet modules, CalculationStamp stamp, CancellationToken cancellationToken)
    {
        ScriptExecutionPolicy.Current.RequireEnabled();
        cancellationToken.ThrowIfCancellationRequested();
        if (!modules.HasTypeScript && modules.Signature is null) return null;
        var request = Request(modules);
        if (modules.Signature is not null)
        {
            if (!QuickExcel.Security.ScriptSignatures.VerifyContent(modules))
                throw new InvalidOperationException(MessageText.Get("脚本签名无效，请审阅源码并重新签名。"));
            if (ScriptCompilationCache.TryRead(modules, request, out var cached))
            {
                cancellationToken.ThrowIfCancellationRequested();
                return cached;
            }
        }
        var result = await CompileAsync(modules, request, stamp, cancellationToken).ConfigureAwait(false);
        if (modules.Signature is not null)
        {
            ScriptCompilationCache.StoreLocal(modules, request, result, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
        }
        return result;
    }

    internal static CompileRequest Request(ScriptModuleSet modules)
    {
        if (ScriptModuleSet.TypeScriptOptions != CompilerProtocol.Options) throw new InvalidOperationException("Compilation option identity mismatch.");
        var request = new CompileRequest(modules.Modules.Select(m => new CompilerModule(m.Key,
            m.Value.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript", m.Value.Source)).ToArray(), modules.EntryModules.ToArray(), "")
            { ApiVersion = modules.ScriptApiVersion, SourceMapMode = SourceMapMode };
        return request with { Sha256 = request.ComputeDigest() };
    }

    internal static async Task<CompiledReply> CompileAsync(ScriptModuleSet modules, CompileRequest request,
        CalculationStamp stamp, CancellationToken cancellationToken)
    {
        var result = await Compilers.GetValue(modules.Modules, _ => new()).CompileAsync(request, stamp, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (!result.Result.Success) throw new ScriptCompilationException(Array.AsReadOnly(result.Result.Diagnostics));
        return result;
    }

    public static bool IsCacheDirty(ScriptModuleSet modules) => !ScriptCompilationCache.TryRead(modules, Request(modules), out _);
    public static void InvalidateCache(ScriptModuleSet modules)
    {
        ScriptCompilationCache.Invalidate(modules);
        if (Compilers.TryGetValue(modules.Modules, out var compiler)) compiler.ClearCache();
    }
    internal static int ProcessLaunchCount(ScriptModuleSet modules) => Compilers.TryGetValue(modules.Modules, out var compiler) ? compiler.ProcessLaunchCount : 0;
}
