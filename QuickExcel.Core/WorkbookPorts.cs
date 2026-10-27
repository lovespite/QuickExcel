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

public interface IWorkbookCalculator
{
    CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default);
}

public interface IScriptRuntimeFactory
{
    IScriptRuntime Create(string source);
}

public interface IScriptRuntime : IDisposable
{
    IReadOnlyList<string> FunctionNames { get; }
    object? Invoke(string name, IReadOnlyList<object> arguments);
    IReadOnlyList<string> GetLogs();
}

public readonly record struct CalculationStamp(Guid SessionId, long InputRevision, long CalculationGeneration, long ScriptGeneration);
public sealed record CalculationRequest(CalculationStamp Stamp, WorkbookSnapshot Snapshot, bool ScriptsEnabled);

public sealed class CalculationResult
{
    public CalculationResult(CalculationStamp stamp, IEnumerable<KeyValuePair<CellAddress, CellCalculation>> cells,
        IEnumerable<string>? diagnostics = null)
    {
        Stamp = stamp;
        Cells = cells.ToImmutableDictionary();
        Diagnostics = (diagnostics ?? []).Take(100).Select(message => message.Length > 4000 ? message[..4000] : message).ToImmutableArray();
    }

    public CalculationStamp Stamp { get; }
    public ImmutableDictionary<CellAddress, CellCalculation> Cells { get; }
    public ImmutableArray<string> Diagnostics { get; }
}

/// <summary>Instances are supplied by the desktop/test composition root. Core never creates a native runtime.</summary>
public sealed record WorkbookServices(IWorkbookStorage Storage, IWorkbookCalculator Calculator);
