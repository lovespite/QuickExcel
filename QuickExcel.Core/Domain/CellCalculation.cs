using System.Collections.ObjectModel;

namespace QuickExcel.Domain;

public enum CalculationState { Fresh, Dirty, Running, Stale, Blocked, Error }

/// <summary>An immutable result/cache description, separate from input and from submission authorization.</summary>
public sealed class CellCalculation
{
    public CellCalculation(CellValue? value, CalculationState state, long inputRevision,
        long calculationGeneration, long scriptGeneration, IEnumerable<string>? diagnostics = null)
    {
        if (!Enum.IsDefined(state)) throw new ArgumentOutOfRangeException(nameof(state));
        ArgumentOutOfRangeException.ThrowIfNegative(inputRevision);
        ArgumentOutOfRangeException.ThrowIfNegative(calculationGeneration);
        ArgumentOutOfRangeException.ThrowIfNegative(scriptGeneration);
        if (state == CalculationState.Fresh && (value is null || value.Value.Kind == CellValueKind.Error))
            throw new ArgumentException("A fresh result requires a non-error value.", nameof(value));
        if (state == CalculationState.Error && (value is null || value.Value.Kind != CellValueKind.Error))
            throw new ArgumentException("An error result requires an error value.", nameof(value));

        var messages = diagnostics?.ToArray() ?? [];
        if (messages.Any(message => message is null))
            throw new ArgumentException("Diagnostics cannot contain null entries.", nameof(diagnostics));
        Value = value;
        State = state;
        InputRevision = inputRevision;
        CalculationGeneration = calculationGeneration;
        ScriptGeneration = scriptGeneration;
        Diagnostics = Array.AsReadOnly(messages);
    }

    /// <summary>Null means no cached result; CellValue.Blank is a present blank result.</summary>
    public CellValue? Value { get; }
    public CalculationState State { get; }
    public long InputRevision { get; }
    public long CalculationGeneration { get; }
    public long ScriptGeneration { get; }
    public ReadOnlyCollection<string> Diagnostics { get; }
}
