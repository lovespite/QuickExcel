using System.Collections.Immutable;
using System.Text;

namespace QuickExcel.Core;

/// <summary>Immutable, ordinal string storage. Quotas count UTF-8 key + value bytes.</summary>
public sealed class ScriptStorage
{
    public const int MaxEntryBytes = 1024 * 1024;
    public const long MaxSessionBytes = 100L * 1024 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    public static ScriptStorage Empty { get; } = new(ImmutableSortedDictionary.Create<string, string>(StringComparer.Ordinal), 0);
    private ScriptStorage(ImmutableSortedDictionary<string, string> entries, long bytes) { Entries = entries; Bytes = bytes; }
    public ImmutableSortedDictionary<string, string> Entries { get; }
    public long Bytes { get; }
    public string? Get(string key) { ValidateKey(key); return Entries.GetValueOrDefault(key); }
    public static void ValidateKey(string key)
    {
        ArgumentNullException.ThrowIfNull(key);
        if (Utf8.GetByteCount(key) > MaxEntryBytes) throw new ArgumentException("存储键超过 1 MiB。");
    }
    private static long Size(string key, string value) => (long)Utf8.GetByteCount(key) + Utf8.GetByteCount(value);
    public ScriptStorage Set(string key, string value, long? maxBytes = null)
    {
        ValidateKey(key); ArgumentNullException.ThrowIfNull(value);
        var size = Size(key, value);
        if (size > MaxEntryBytes) throw new ScriptStorageQuotaException("存储单条键和值的 UTF-8 总量超过 1 MiB。");
        var old = Entries.GetValueOrDefault(key);
        var bytes = checked(Bytes + size - (old is null ? 0 : Size(key, old)));
        if (maxBytes is { } maximum && bytes > maximum) throw new ScriptStorageQuotaException("sessionStorage 总量超过 100 MiB。");
        return old == value ? this : new(Entries.SetItem(key, value), bytes);
    }
    public ScriptStorage Remove(string key)
    {
        ValidateKey(key);
        return Entries.TryGetValue(key, out var value) ? new(Entries.Remove(key), Bytes - Size(key, value)) : this;
    }
    public bool ContentEquals(ScriptStorage other) => ReferenceEquals(this, other) || Entries.SequenceEqual(other.Entries);
    // Charge changed entries retained by undo/redo, not the entire store on every edit.
    internal long HistoryBytesFrom(ScriptStorage before) => ReferenceEquals(this, before) ? 0 :
        256L + Entries.Where(p => before.Entries.GetValueOrDefault(p.Key) != p.Value)
            .Sum(p => 128L + 2L * (p.Key.Length + p.Value.Length)) +
        before.Entries.Where(p => Entries.GetValueOrDefault(p.Key) != p.Value)
            .Sum(p => 128L + 2L * (p.Key.Length + p.Value.Length));
}

public sealed class ScriptStorageQuotaException(string message) : InvalidOperationException(message);

public sealed partial class DocumentSession
{
    private ScriptStorage _sessionStorage = ScriptStorage.Empty;
    public ScriptStorage SessionStorage { get { lock (_gate) { ThrowIfClosed(); return _sessionStorage; } } }
}
