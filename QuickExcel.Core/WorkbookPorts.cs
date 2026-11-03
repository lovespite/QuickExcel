using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public interface IWorkbookStorage
{
    WorkbookSnapshot Open(string path);
    void Save(string path, WorkbookSnapshot snapshot);
    void Export(string path, WorkbookSnapshot snapshot);
}

public interface IWorkbookStorageOptions : IWorkbookStorage
{
    void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported);
    void Export(string path, WorkbookSnapshot snapshot, bool allowStaleCaches, bool discardUnsupported);
}

/// <summary>A null digest means the destination must still be absent.</summary>
public sealed record FileVersion(string? Sha256);
public sealed class WorkbookFileConflictException() : IOException("文件已被其他程序修改、替换或删除。请重新打开，或另存到其他文件；当前编辑仍保留。");
public interface IVersionedWorkbookStorage : IWorkbookStorageOptions
{
    FileVersion SaveChecked(string path, WorkbookSnapshot snapshot, bool discardUnsupported, FileVersion expected);
}

public interface IWorkbookCalculator
{
    CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default);
}

public interface IScriptRuntimeFactory
{
    IScriptRuntime Create(string source);
    IScriptRuntime Create(ScriptRuntimeRequest request) => request.Modules.CanExecuteLegacy
        ? Create(request.Source) : throw new NotSupportedException("This runtime does not support package modules.");
}

public sealed record ScriptRuntimeRequest(ScriptModuleSet Modules, CalculationStamp Stamp, TimeSpan TimeLimit, CancellationToken CancellationToken)
{
    public ScriptRuntimeRequest(string source, CalculationStamp stamp, TimeSpan timeLimit, CancellationToken cancellationToken)
        : this(ScriptModuleSet.Single(source), stamp, timeLimit, cancellationToken) { }
    public string Source => Modules.PrimarySource;
}

/// <summary>Transport/lifecycle faults abort the entire calculation, rather than becoming a single cell error.</summary>
public sealed class ScriptRuntimeFailureException(string code, string message, Exception? inner = null) : Exception(message, inner)
{
    public string Code { get; } = code;
    public IReadOnlyList<string> Logs { get; init; } = [];
}

public interface IScriptRuntime : IDisposable
{
    IReadOnlyList<string> FunctionNames { get; }
    IReadOnlyList<ScriptFunctionDescriptor> FunctionDescriptors => FunctionNames.Select(n => new ScriptFunctionDescriptor(n)).ToArray();
    object? Invoke(string name, IReadOnlyList<object> arguments);
    IReadOnlyList<string> GetLogs();
}

/// <summary>Optional ordered batching. Admission must be bounded and must not execute user code.</summary>
public interface IScriptBatchRuntime : IScriptRuntime
{
    IScriptInvocationBatch CreateBatch();
}

public interface IScriptInvocationBatch
{
    int Count { get; }
    // False means flush this batch first. An individually oversized call must throw.
    bool TryAdd(string name, IReadOnlyList<object> arguments);
    IReadOnlyList<ScriptCallResult> Invoke();
}

public sealed record ScriptCallResult(object? Value, string? Error, bool ValueTooLong = false);

public readonly record struct CalculationStamp(Guid SessionId, long InputRevision, long CalculationGeneration, long ScriptGeneration);
public sealed record CalculationRequest(CalculationStamp Stamp, WorkbookSnapshot Snapshot, bool ScriptsEnabled)
{
    public bool ForceFull { get; init; }
    public CancellationToken CancellationToken { get; init; }
}
public sealed record CalculationStatistics(int FormulaCount, int EvaluatedCount, int ReusedCount, bool FullRecalculation, TimeSpan Elapsed);

public sealed class CalculationResult
{
    public CalculationResult(CalculationStamp stamp, IEnumerable<KeyValuePair<CellAddress, CellCalculation>> cells,
        IEnumerable<string>? diagnostics = null)
    {
        Stamp = stamp;
        Cells = cells.ToImmutableDictionary();
        Diagnostics = (diagnostics ?? []).Take(100).Select(message => message.Length > 4000 ? message[..4000] : message).ToImmutableArray();
    }

    public CalculationStatistics? Statistics { get; init; }
    public CalculationStamp Stamp { get; }
    public ImmutableDictionary<CellAddress, CellCalculation> Cells { get; }
    public ImmutableArray<string> Diagnostics { get; }
}

/// <summary>Instances are supplied by the desktop/test composition root. Core never creates a native runtime.</summary>
public sealed record WorkbookServices(IWorkbookStorage Storage, IWorkbookCalculator Calculator, IScriptActionExecutor? Actions = null);

public sealed record PortableExportReport(int RetainedFormulas, int FrozenFormulas, int MissingResults, int StaleResults);
public interface IPortableWorkbookStorage
{
    PortableExportReport InspectPortableExport(WorkbookSnapshot snapshot);
    void ExportPortable(string path, WorkbookSnapshot snapshot, bool allowStaleCaches = false, bool discardUnsupported = false);
}
