using System.Security.Cryptography;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Presentation.Web;

public sealed record DraftChunk(string Module, long BaseVersion, long Version, string Transfer, int Index, int Count, int TotalBytes, string Sha256, string Data);
public sealed record ScriptDraft(string Path, string Committed, string Text, long Version = 1, long ObservedVersion = 1, ScriptLanguage Language = ScriptLanguage.JavaScript);

/// <summary>Confirmed editor drafts survive a page failure. Only a matching flush barrier permits a document command.</summary>
public sealed class ScriptDraftSession
{
    public const int MaximumSourceBytes = 1024 * 1024, ChunkBytes = 48 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private readonly Dictionary<string, ScriptDraft> _models;
    private readonly TimeProvider _clock;
    private TransferState? _transfer;
    private bool _flushConfirmed;
    private sealed class TransferState(DraftChunk first, long started)
    {
        public readonly DraftChunk First = first;
        public readonly long Started = started;
        public readonly byte[] Bytes = new byte[first.TotalBytes];
        public int Next, Written;
    }
    public IReadOnlyDictionary<string, ScriptDraft> Models => _models;
    public bool IsReadOnly { get; }
    public int ApiVersion { get; }
    public bool Failed { get; private set; }
    public bool HasPending => _transfer is not null || _models.Values.Any(m => m.Version != m.ObservedVersion);
    public bool IsDirty => HasPending || _models.Values.Any(m => m.Text != m.Committed);

    public ScriptDraftSession(ScriptModuleSet modules, bool readOnly = false, TimeProvider? clock = null)
    {
        IsReadOnly = readOnly; ApiVersion = modules.ScriptApiVersion; _clock = clock ?? TimeProvider.System;
        _models = modules.Modules.ToDictionary(p => p.Key, p => new ScriptDraft(p.Value.Path, p.Value.Source, p.Value.Source, Language: p.Value.Language), StringComparer.Ordinal);
        if (_models.Values.Sum(m => (long)Utf8.GetByteCount(m.Text)) > MaximumSourceBytes) throw new InvalidDataException(UiText.Get("脚本总量不能超过 1 MiB。"));
    }
    private ScriptDraft Require(string module)
    {
        if (Failed) throw new InvalidOperationException(UiText.Get("脚本页面已失效，请先恢复已确认草稿。"));
        return _models.TryGetValue(module, out var model) ? model : throw new InvalidDataException(UiText.Get("未知模块。"));
    }
    public void Observe(string module, long version)
    {
        var model = Require(module);
        if (IsReadOnly || version <= model.ObservedVersion || version > 9_007_199_254_740_991) throw new InvalidDataException(UiText.Get("草稿版本重放或只读编辑。"));
        _models[module] = model with { ObservedVersion = version };
        _flushConfirmed = false;
    }
    public bool Receive(DraftChunk chunk)
    {
        _flushConfirmed = false;
        try
        {
            var model = Require(chunk.Module);
            if (IsReadOnly || chunk.BaseVersion != model.Version || chunk.Version <= model.Version || chunk.Version > model.ObservedVersion ||
                !WebMessage.ValidId(chunk.Transfer) || chunk.TotalBytes is < 0 or > MaximumSourceBytes || chunk.Count is < 1 or > 32 ||
                chunk.Index < 0 || chunk.Index >= chunk.Count || chunk.Sha256 is not { Length: 64 } || chunk.Sha256.Any(c => !Uri.IsHexDigit(c)) ||
                chunk.Data is null || chunk.Data.Length > (ChunkBytes + 2) / 3 * 4)
                throw new InvalidDataException(UiText.Get("草稿分块身份、版本或预算无效。"));
            if (_transfer is null)
            {
                if (chunk.Index != 0) throw new InvalidDataException(UiText.Get("草稿分块必须从 0 开始。"));
                _transfer = new(chunk, _clock.GetTimestamp());
            }
            var transfer = _transfer; var first = transfer.First;
            if (_clock.GetElapsedTime(transfer.Started) > TimeSpan.FromSeconds(10) || chunk.Index != transfer.Next ||
                chunk.Module != first.Module || chunk.Transfer != first.Transfer || chunk.Version != first.Version || chunk.BaseVersion != first.BaseVersion ||
                chunk.Count != first.Count || chunk.TotalBytes != first.TotalBytes || chunk.Sha256 != first.Sha256)
                throw new InvalidDataException(UiText.Get("草稿分块重复、乱序、混入旧传输或已超时。"));
            var bytes = Convert.FromBase64String(chunk.Data);
            if (bytes.Length > ChunkBytes || transfer.Written > transfer.Bytes.Length - bytes.Length) throw new InvalidDataException(UiText.Get("草稿分块长度超限。"));
            bytes.CopyTo(transfer.Bytes, transfer.Written); transfer.Written += bytes.Length; transfer.Next++;
            if (transfer.Next != first.Count) return false;
            if (transfer.Written != first.TotalBytes || !Convert.ToHexString(SHA256.HashData(transfer.Bytes)).Equals(first.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(UiText.Get("草稿整体长度或 SHA-256 不匹配。"));
            var text = Utf8.GetString(transfer.Bytes);
            if (_models.Where(p => p.Key != chunk.Module).Sum(p => (long)Utf8.GetByteCount(p.Value.Text)) + first.TotalBytes > MaximumSourceBytes)
                throw new InvalidDataException(UiText.Get("所有模块草稿合计超过 1 MiB。"));
            _models[chunk.Module] = model with { Text = text, Version = chunk.Version };
            _transfer = null; return true;
        }
        catch { _transfer = null; throw; }
    }
    public void ConfirmFlush(IReadOnlyDictionary<string, long> versions)
    {
        if (Failed || HasPending || versions.Count != _models.Count || _models.Any(p => !versions.TryGetValue(p.Key, out var version) || version != p.Value.Version))
            throw new InvalidOperationException(UiText.Get("页面尚未确认全部最新草稿；本次没有保存或运行旧源码。"));
        _flushConfirmed = true;
    }
    public ScriptModuleSet BuildModules(ScriptModuleSet current)
    {
        if (Failed || HasPending || !_flushConfirmed) throw new InvalidOperationException(UiText.Get("草稿尚未完成版本确认。"));
        if (current.Modules.Count != _models.Count || current.Modules.Any(p => !_models.TryGetValue(p.Key, out var m) || m.Path != p.Value.Path || m.Language != p.Value.Language || m.Committed != p.Value.Source))
            throw new InvalidOperationException(UiText.Get("模块在同步期间已改变，请重新加载已确认草稿后重试。"));
        _flushConfirmed = false;
        if (current.Modules.All(p => p.Value.Source == _models[p.Key].Text)) return current;
        return new(current.Modules.Select(p => KeyValuePair.Create(p.Key, p.Value with { Source = _models[p.Key].Text })), current.EntryModules,
            current.Revision, current.BindingRevision, current.ScriptApiVersion, current.Bindings, current.Signature);
    }
    public void MarkCommitted(ScriptModuleSet modules)
    {
        foreach (var (id, model) in _models.ToArray())
        {
            if (!modules.Modules.TryGetValue(id, out var module) || module.Source != model.Text) throw new InvalidOperationException(UiText.Get("提交结果与已确认草稿不一致。"));
            _models[id] = model with { Committed = model.Text };
        }
    }
    public void MarkFailed() { Failed = true; _transfer = null; _flushConfirmed = false; }
    public void RecoverConfirmed()
    {
        // Explicit UI recovery acknowledges that unconfirmed browser text cannot be recovered.
        Failed = false; _transfer = null; _flushConfirmed = false;
        foreach (var (id, model) in _models.ToArray()) _models[id] = model with { Version = 1, ObservedVersion = 1 };
    }
    public static IEnumerable<DraftChunk> Split(string module, string text, long baseVersion, long version, string? transfer = null)
    {
        var bytes = Utf8.GetBytes(text); if (bytes.Length > MaximumSourceBytes) throw new InvalidDataException(UiText.Get("源码超过 1 MiB。"));
        var hash = Convert.ToHexString(SHA256.HashData(bytes)); var id = transfer ?? Guid.NewGuid().ToString("N"); var count = Math.Max(1, (bytes.Length + ChunkBytes - 1) / ChunkBytes);
        for (var i = 0; i < count; i++) yield return new(module, baseVersion, version, id, i, count, bytes.Length, hash,
            Convert.ToBase64String(bytes.AsSpan(i * ChunkBytes, Math.Min(ChunkBytes, bytes.Length - i * ChunkBytes))));
    }
}
