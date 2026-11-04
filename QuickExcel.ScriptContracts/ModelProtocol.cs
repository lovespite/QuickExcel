using System.Text.Json;

namespace QuickExcel.ScriptContracts;

public sealed record ModelCall(Guid OperationId, int Sequence, JsonElement Request);
public sealed record ModelReply(Guid OperationId, int Sequence, bool Ok, JsonElement Value, string? Error)
{
    /// <summary>Stable script-visible error code; null when the call succeeded.</summary>
    public string? Code { get; init; }
}
