namespace QuickExcel.ScriptContracts;

public sealed record VirtualEntry(string Path, string? Base64, bool IsDirectory);
public sealed record VirtualEdit(string Operation, string Path, string? Base64);
public sealed record ActionEvent(string Name, Guid SessionId, long Sequence, long Revision, string? SheetId, string? SheetName, string[] Ranges, bool IsUndoOrRedo);

public static class VirtualFileProtocol
{
    public const int MaxEntries = 4096, MaxBytes = 8 * 1024 * 1024;
    public static void Path(string path)
    {
        if (path is null || path.Length > 248 || !path.StartsWith("//", StringComparison.Ordinal) ||
            path.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 || path.Any(char.IsControl)) throw new ProtocolException("Invalid virtual path.");
        var parts = path[2..].Split('/');
        if (parts.Length < 2 || parts[0] is not ("xlsq" or "temp" or "mem") ||
            parts.Skip(1).Any(s => s.Length == 0 || s is "." or ".." || s.EndsWith('.') || s.EndsWith(' '))) throw new ProtocolException("Invalid virtual path.");
    }
    public static byte[] Bytes(string? base64)
    {
        if (base64 is null || base64.Length > (MaxBytes + 2) / 3 * 4) throw new ProtocolException("Invalid file capacity.");
        try { return Convert.FromBase64String(base64); }
        catch (FormatException error) { throw new ProtocolException("Invalid file encoding.", error); }
    }
    public static void Validate(VirtualEntry[] files)
    {
        if (files is null || files.Length > MaxEntries) throw new ProtocolException("Invalid file count.");
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase); long bytes = 0;
        foreach (var file in files)
        {
            if (file is null) throw new ProtocolException("Missing file.");
            Path(file.Path);
            if (!paths.Add(file.Path) || file.IsDirectory && file.Base64 is not null ||
                (bytes += file.IsDirectory ? 0 : Bytes(file.Base64).Length) > MaxBytes) throw new ProtocolException("Invalid files or quota.");
        }
    }
    public static void ValidateEdits(VirtualEdit[] edits)
    {
        if (edits is null || edits.Length > MaxEntries) throw new ProtocolException("Invalid file edit count.");
        long bytes = 0;
        foreach (var edit in edits)
        {
            if (edit is null) throw new ProtocolException("Missing file edit.");
            Path(edit.Path);
            if (edit.Operation is not ("write" or "mkdir" or "remove") || edit.Operation != "write" && edit.Base64 is not null ||
                (bytes += edit.Operation == "write" ? Bytes(edit.Base64).Length : 0) > MaxBytes) throw new ProtocolException("Invalid file edit or quota.");
        }
    }
}
