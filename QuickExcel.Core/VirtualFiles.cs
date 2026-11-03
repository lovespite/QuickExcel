using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Logical names only. Never pass these paths to an OS filesystem API.</summary>
public static class VirtualPath
{
    public static string Validate(string path, bool allowRoot = false)
    {
        if (path is null || path.Length > 248 || !path.StartsWith("//", StringComparison.Ordinal) ||
            path.IndexOfAny(['\\', ':', '%', '?', '#', '\0']) >= 0 || path.Any(char.IsControl))
            throw new ArgumentException("虚拟路径必须使用 //xlsq/、//temp/ 或 //mem/ 和 / 分隔符。");
        var segments = path[2..].Split('/');
        if (segments.Length < 2 || segments[0] is not ("xlsq" or "temp" or "mem")) throw new ArgumentException("未知虚拟文件卷。");
        if (allowRoot && segments.Length == 2 && segments[1] == "") return path;
        if (segments.Skip(1).Any(s => s.Length == 0 || s is "." or ".." || s.EndsWith('.') || s.EndsWith(' ')))
            throw new ArgumentException("虚拟路径包含空段、相对段或歧义段。");
        return path;
    }
}

public sealed record VirtualFileEntry(string Path, ImmutableArray<byte> Data, bool IsDirectory = false);
public sealed record VirtualFileEdit(string Operation, string Path, ImmutableArray<byte> Data = default);

/// <summary>Immutable transaction data; temp and mem are memory backed, with no host path mapping.</summary>
public sealed class VirtualFileSet
{
    public const int MaxEntries = 4096;
    public const int MaxBytes = 8 * 1024 * 1024;
    public static VirtualFileSet Empty { get; } = new([]);
    public ImmutableDictionary<string, VirtualFileEntry> Entries { get; }
    public long Bytes => Entries.Values.Sum(e => (long)e.Data.Length + e.Path.Length * 2L + 64);

    public VirtualFileSet(IEnumerable<VirtualFileEntry> entries)
    {
        var map = ImmutableDictionary.CreateBuilder<string, VirtualFileEntry>(StringComparer.Ordinal);
        var spelling = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long size = 0;
        foreach (var entry in entries)
        {
            if (entry is null) throw new ArgumentException("缺少虚拟条目。");
            VirtualPath.Validate(entry.Path);
            if (entry.Data.IsDefault || entry.IsDirectory && !entry.Data.IsEmpty || !spelling.Add(entry.Path) || map.Count >= MaxEntries ||
                (size += entry.Data.Length) > MaxBytes) throw new ArgumentException("虚拟文件重复、目录有内容或超过 4096 条目 / 8 MiB 容量。");
            map.Add(entry.Path, entry);
        }
        foreach (var path in map.Keys)
        {
            for (var end = path.LastIndexOf('/'); end > path.IndexOf('/', 2); end = path.LastIndexOf('/', end - 1))
            {
                var parent = path[..end];
                if (map.TryGetValue(parent, out var entry) && !entry.IsDirectory) throw new ArgumentException("父路径不是目录。");
                if (spelling.TryGetValue(parent, out var actual) && actual != parent) throw new ArgumentException("目录大小写歧义。");
                if (!spelling.Contains(parent)) spelling.Add(parent);
            }
        }
        Entries = map.ToImmutable();
    }

    public VirtualFileSet Apply(IEnumerable<VirtualFileEdit> edits)
    {
        var map = Entries.ToBuilder(); var count = 0; long written = 0;
        foreach (var edit in edits)
        {
            if (edit is null || ++count > MaxEntries) throw new ArgumentException("虚拟文件操作超限。");
            VirtualPath.Validate(edit.Path);
            switch (edit.Operation)
            {
                case "write":
                    if (edit.Data.IsDefault || map.TryGetValue(edit.Path, out var old) && old.IsDirectory) throw new ArgumentException("文件数据或类型无效。");
                    if ((written += edit.Data.Length) > MaxBytes) throw new ArgumentException("本次累计文件写入超过 8 MiB。");
                    map[edit.Path] = new(edit.Path, edit.Data); break;
                case "mkdir":
                    if (!edit.Data.IsDefaultOrEmpty || map.TryGetValue(edit.Path, out var existing) && !existing.IsDirectory) throw new ArgumentException("目录类型无效。");
                    map[edit.Path] = new(edit.Path, [], true); break;
                case "remove":
                    if (!edit.Data.IsDefaultOrEmpty || map.Keys.Any(p => p.StartsWith(edit.Path + "/", StringComparison.Ordinal))) throw new ArgumentException("目录非空；请先删除子条目。");
                    if (!map.Remove(edit.Path)) throw new ArgumentException("虚拟文件不存在。");
                    break;
                default: throw new ArgumentException("未知虚拟文件操作。");
            }
            // Validate each intermediate state too; a later delete cannot hide a quota/path violation.
            _ = new VirtualFileSet(map.Values);
        }
        return new(map.Values);
    }

    public VirtualFileSet Volume(string prefix) => new(Entries.Values.Where(e => e.Path.StartsWith(prefix, StringComparison.Ordinal)));
    public bool ContentEquals(VirtualFileSet other) => Entries.Count == other.Entries.Count && Entries.All(p =>
        other.Entries.TryGetValue(p.Key, out var value) && p.Value.IsDirectory == value.IsDirectory && p.Value.Data.SequenceEqual(value.Data));
}

public interface IWorkbookVirtualFiles
{
    VirtualFileSet Capture(WorkbookSnapshot snapshot);
    WorkbookSnapshot Apply(WorkbookSnapshot snapshot, IReadOnlyList<VirtualFileEdit> edits);
}
