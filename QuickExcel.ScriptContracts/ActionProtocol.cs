using System.Text.Json;

namespace QuickExcel.ScriptContracts;

public sealed record ActionInfo(string Name, string Description);
public sealed record LoadedActions(string Sha256, ActionInfo[] Actions)
{
    public string[] NetworkHosts { get; init; } = [];
}
public sealed record ActionCell(int Row, int Column, WireValue Value, string? Formula, string State, bool HasValue);
public sealed record ActionSheet(string Id, string Name, ActionCell[] Cells);
public sealed record ActionSnapshot(long BaseRevision, string DateSystem, ActionSheet[] Sheets);
public sealed record RunAction(Guid OperationId, string Name, ActionSnapshot Snapshot)
{
    public VirtualEntry[] Files { get; init; } = [];
    public ActionEvent? Event { get; init; }
    public bool LiveModel { get; init; }
    public bool AllowNetwork { get; init; }
    public int MaxNetworkBodyBytes { get; init; } = ActionProtocol.DefaultNetworkBodyBytes;
}
public sealed record ActionWrite(string SheetId, int Row, int Column, WireValue? Value, string? Formula);
public sealed record ActionPatch(Guid OperationId, long BaseRevision, string Name, ActionWrite[] Edits)
{
    public VirtualEdit[] FileEdits { get; init; } = [];
}

public static class ActionProtocol
{
    public const int MaxEdits = 20_000;
    public const int MaxSnapshotCells = 100_000;
    public const int MinNetworkBodyBytes = 1024 * 1024;
    public const int DefaultNetworkBodyBytes = 8 * 1024 * 1024;
    // Base64 plus envelope/header metadata must remain below the unchanged
    // 16 MiB frame limit. The round's cumulative byte budget is separate.
    public const int MaxNetworkBodyBytes = 10 * 1024 * 1024;
    public static void ValidateNetworkBodyLimit(int bytes)
    {
        if (bytes is < MinNetworkBodyBytes or > MaxNetworkBodyBytes)
            throw new ProtocolException("Network body limit must be between 1 and 10 MiB.");
    }
    public static void Name(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl)) throw new ProtocolException("Invalid action name.");
    }
    public static void Scalar(WireValue value)
    {
        if (value is null) throw new ProtocolException("Missing action value.");
        if (value.Kind == "error")
        {
            if (value.Value.ValueKind != JsonValueKind.String || value.Value.GetString() is not ("Null" or "DivisionByZero" or "Value" or "Reference" or "Name" or "Number" or "NotAvailable" or "Cycle" or "Script"))
                throw new ProtocolException("Invalid action error value.");
        }
        else value.Validate(scalarOnly: true);
        if (value.Kind == "text" && value.Value.GetString()!.Length > 32767) throw new ProtocolException("Action text exceeds cell capacity.");
    }
    public static void Position(int row, int column)
    {
        if (row is < 0 or >= 100_000 || column is < 0 or >= 1024) throw new ProtocolException("Action address out of bounds.");
    }
    public static void ValidateSnapshot(ActionSnapshot snapshot)
    {
        if (snapshot is null || snapshot.BaseRevision < 0 || snapshot.DateSystem is not ("Windows1900" or "Mac1904") || snapshot.Sheets is null || snapshot.Sheets.Length is < 1 or > 64)
            throw new ProtocolException("Invalid action snapshot.");
        var ids = new HashSet<string>(); var total = 0;
        foreach (var sheet in snapshot.Sheets)
        {
            if (sheet is null || !Guid.TryParseExact(sheet.Id, "D", out var id) || id == Guid.Empty || !ids.Add(sheet.Id) ||
                string.IsNullOrWhiteSpace(sheet.Name) || sheet.Name.Length > 31 || sheet.Cells is null || (total += sheet.Cells.Length) > MaxSnapshotCells)
                throw new ProtocolException("Invalid snapshot sheet or capacity.");
            var positions = new HashSet<(int, int)>();
            foreach (var cell in sheet.Cells)
            {
                if (cell is null || !positions.Add((cell.Row, cell.Column)) || cell.State is not ("literal" or "Dirty" or "Running" or "Stale" or "Fresh" or "Error" or "Blocked") ||
                    cell.Formula is { } formula && (formula.Length > 32767 || !formula.StartsWith('='))) throw new ProtocolException("Invalid snapshot cell.");
                Position(cell.Row, cell.Column); Scalar(cell.Value);
            }
        }
    }
    public static void ValidatePatch(ActionPatch patch, RunAction request)
    {
        if (patch is null || patch.OperationId != request.OperationId || patch.BaseRevision != request.Snapshot.BaseRevision || patch.Name != request.Name ||
            patch.Edits is null || patch.Edits.Length > MaxEdits) throw new ProtocolException("Action patch identity or capacity mismatch.");
        VirtualFileProtocol.ValidateEdits(patch.FileEdits);
        var ids = request.Snapshot.Sheets.Select(s => s.Id).ToHashSet(StringComparer.Ordinal);
        var positions = new HashSet<(string, int, int)>();
        foreach (var edit in patch.Edits)
        {
            if (edit is null || !ids.Contains(edit.SheetId) || !positions.Add((edit.SheetId, edit.Row, edit.Column)) ||
                (edit.Value is null) == (edit.Formula is null)) throw new ProtocolException("Invalid, duplicate or unknown action target.");
            Position(edit.Row, edit.Column);
            if (edit.Value is not null) Scalar(edit.Value);
            if (edit.Formula is { } f && (f.Length is < 2 or > 8192 || !f.StartsWith('='))) throw new ProtocolException("Invalid action formula.");
        }
    }
}
