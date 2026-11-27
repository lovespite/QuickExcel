using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.ScriptContracts;
using static QuickExcel.Runtime.SelectedFileSystem;

namespace QuickExcel.Runtime;

/// <summary>One Action owns this broker and all of its non-transferable grants, streams and cursors.</summary>
internal sealed partial class ScriptUserFiles(ScriptActionRequest request, Func<IDisposable> beginInteraction,
    CancellationToken cancellation) : IDisposable
{
    private const int Block = 64 * 1024, BufferedLimit = 4 * 1024 * 1024, DraftLimit = 16 * 1024 * 1024;
    private sealed class Entry(Target target, Stamp? stamp, bool directory, string access, string mode, bool delete)
    {
        internal readonly Target Target = target;
        internal Stamp? Stamp = stamp;
        internal readonly bool Directory = directory, Delete = delete;
        internal readonly string Access = access;
        internal string Mode = mode;
        internal bool CanRead => Access.Contains('r');
        internal bool CanWrite => Access.Contains('w');
    }

    private sealed class OpenFile(Entry entry, Stream data, bool dirty, bool buffered) : IDisposable
    {
        internal readonly Entry Entry = entry;
        internal readonly Stream Data = data;
        internal readonly bool Buffered = buffered;
        internal bool Dirty = dirty, Failed;
        internal WorkbookDocument? Workbook;
        internal ScriptModelAdapter? WorkbookModel;
        public void Dispose() { Workbook?.Dispose(); Data.Dispose(); }
    }

    private readonly Dictionary<string, Entry> _entries = [];
    private readonly Dictionary<string, OpenFile> _streams = [];
    private readonly Dictionary<string, IEnumerator<JsonObject>> _cursors = [];
    private string? _workbook;
    private Identity? _workbookId;
    private int _pickers, _visited;

    internal ModelReply Call(ModelCall call)
    {
        try
        {
            cancellation.ThrowIfCancellationRequested();
            var packet = call.Request.GetProperty("args")[0];
            var method = packet.GetProperty("fs").GetString()!;
            var arguments = packet.GetProperty("args");
            if (method == "workbookModel") return WorkbookModelCall(call, arguments);
            var data = Dispatch(method, arguments);
            using var json = JsonDocument.Parse(data?.ToJsonString() ?? "null");
            return new(call.OperationId, call.Sequence, true, json.RootElement.Clone(), null);
        }
        catch (Exception error) when (error is not (OperationCanceledException or OutOfMemoryException))
        {
            var (code, message) = error switch
            {
                UserFileError e => (e.Code, e.Message),
                InvalidDataException => ("FileError", "Invalid or unsupported workbook contents."),
                QuickExcel.Storage.WorkbookCompatibilityException => ("UnsupportedApi", "Saving the edits would discard unsupported workbook content."),
                UnauthorizedAccessException => ("PermissionDenied", "File access denied."),
                IOException => ("FileError", "Local file operation failed."),
                NotSupportedException => ("UnsupportedApi", "This operation is unavailable on the selected storage."),
                _ => ("InvalidArgument", "Invalid file request.")
            };
            return new(call.OperationId, call.Sequence, false, WorkerProtocol.EncodePayload<string?>(null), message) { Code = code };
        }
    }

    private JsonNode? Dispatch(string method, JsonElement a)
    {
        if (a.ValueKind != JsonValueKind.Array || a.GetArrayLength() > 5) throw UserFileError.Invalid();
        if (method is "selectFiles" or "selectDirectory") return Select(a[0], method == "selectDirectory");
        if (method == "beginSaveAs") return BeginSaveAs();
        var id = a[0].GetString()!;
        if (method == "next") return Next(id);
        if (method == "endEntries") { if (_cursors.Remove(id, out var c)) c.Dispose(); return null; }
        if (method == "abort") { if (_streams.Remove(id, out var s)) s.Dispose(); return null; }
        if (method is "read" or "write" or "close") return StreamCall(method, id, a);
        if (method == "saveAs") { SaveAs(id, a); return null; }
        if (!_entries.TryGetValue(id, out var entry)) throw new UserFileError("InvalidReference", "File handle expired.");
        if (method == "release") { _entries.Remove(id); return null; }
        using var chain = new Chain(entry.Target.Path, entry.Target.Parents);
        Verify(entry, chain.Observe());
        switch (method)
        {
            case "stat": return Info(entry.Target.Path, entry.Stamp, entry.Directory);
            case "open" when !entry.Directory: return OpenStream(entry, chain, a[1].GetBoolean());
            case "openWorkbook" when !entry.Directory: return OpenStream(entry, chain, buffered: false, workbook: true);
            case "getFile" when entry.Directory: return Child(entry, a[1].GetString()!, a[2], false);
            case "getDirectory" when entry.Directory: return Child(entry, a[1].GetString()!, a[2], true);
            case "entries" when entry.Directory:
                if (!entry.CanRead) throw UserFileError.Denied();
                if (_cursors.Count >= 8) throw UserFileError.Limit();
                var cursor = Id();
                _cursors.Add(cursor, Enumerate(entry).GetEnumerator());
                return cursor;
            case "remove" when entry.Directory:
                if (!entry.Delete) throw UserFileError.Denied();
                var child = ChildTarget(entry, a[1].GetString()!);
                int removed = 0;
                try { RemoveTree(child, a[2].GetBoolean(), ref removed); }
                catch (Exception ex) when (removed > 0 && ex is not OperationCanceledException)
                { throw new UserFileError("FileError", $"Deletion stopped after {removed} entries."); }
                return null;
            default: throw UserFileError.Invalid();
        }
    }

    private static string Id() => Guid.NewGuid().ToString("N");
    private static string Leaf(string path) => Path.GetFileName(path) is { Length: > 0 } name ? name : "Volume";
    private static string? String(JsonElement opts, string name) => opts.TryGetProperty(name, out var v) ? v.GetString() ?? throw UserFileError.Invalid() : null;
    private static bool Flag(JsonElement opts, string name) => opts.TryGetProperty(name, out var v) && v.GetBoolean();
    private static void Keys(JsonElement opts, params string[] allowed)
    {
        if (opts.ValueKind != JsonValueKind.Object || opts.EnumerateObject().Any(p => !allowed.Contains(p.Name))) throw UserFileError.Invalid();
    }
    private static string Label(string text)
    {
        if (text.Length > 160 || text.Any(c => char.IsControl(c) || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.Format))
            throw UserFileError.Invalid();
        return text;
    }
    private static void ValidateAccess(string access, string mode)
    {
        if (access is not ("r" or "w" or "rw") || mode is not ("open" or "create" or "open_or_create") || mode == "create" && access == "r")
            throw UserFileError.Invalid();
    }

    private static string FilterExtension(JsonElement value)
    {
        var extension = value.GetString() ?? throw UserFileError.Invalid();
        if (extension.StartsWith('.')) extension = extension[1..];
        if (extension.Length is 0 or > 24 || !extension.All(char.IsAsciiLetterOrDigit))
            throw UserFileError.Invalid();
        return extension;
    }

    private JsonArray Select(JsonElement options, bool directory)
    {
        if (++_pickers > 8) throw UserFileError.Limit();
        if (!OperatingSystem.IsWindows()) throw UserFileError.Unsupported();
        if (directory) Keys(options, "title", "fileAccess", "multiSelect", "allowDelete");
        else Keys(options, "title", "filters", "mode", "multiSelect", "fileMode", "fileAccess", "suggestedName");
        var mode = directory ? "open" : String(options, "fileMode") ?? throw UserFileError.Invalid();
        var access = String(options, "fileAccess") ?? (directory ? "r" : throw UserFileError.Invalid());
        ValidateAccess(access, mode);
        var save = !directory && String(options, "mode") switch { "save" => true, "open" => false, _ => throw UserFileError.Invalid() };
        var multi = Flag(options, "multiSelect");
        var suggested = String(options, "suggestedName");
        if (save && multi || suggested is not null && !save) throw UserFileError.Invalid();
        if (suggested is not null) Name(suggested);
        var filters = new List<ScriptFileFilter>();
        if (options.TryGetProperty("filters", out var raw))
        {
            if (raw.GetArrayLength() > 16) throw UserFileError.Limit();
            foreach (var filter in raw.EnumerateArray())
            {
                Keys(filter, "name", "extensions");
                var extensions = filter.GetProperty("extensions");
                if (extensions.GetArrayLength() is 0 or > 16) throw UserFileError.Invalid();
                var values = extensions.EnumerateArray().Select(FilterExtension).ToArray();
                filters.Add(new(Label(String(filter, "name") ?? throw UserFileError.Invalid()), values));
            }
        }
        var delete = directory && Flag(options, "allowDelete");
        var result = request.UserFilePicker!(new(request.Name!, directory, save, multi, access, mode, delete,
            String(options, "title") is { } title ? Label(title) : null, suggested, filters), beginInteraction, cancellation)
            .WaitAsync(cancellation).GetAwaiter().GetResult();
        cancellation.ThrowIfCancellationRequested();
        if (result.LocalPaths.Count > (multi ? 64 : 1)) throw UserFileError.Limit();
        if (result.LocalPaths.Count == 0) return [];
        _workbook = result.WorkbookPath;
        if (_workbook is not null && (access != "r" || delete))
            _workbookId = WorkbookIdentity(_workbook);
        var items = new JsonArray();
        foreach (var path in result.LocalPaths)
        {
            using var chain = new Chain(path);
            var stamp = chain.Observe();
            CheckType(stamp, directory, save ? mode : "open", access);
            JsonNode item = Grant(new(chain.Target, stamp, directory, access, mode, delete));
            items.Add(item);
        }
        return items;
    }

    private static void CheckType(Stamp? observed, bool directory, string mode, string access)
    {
        if (observed is null)
        {
            if (mode == "open") throw new UserFileError("ItemNotFound", "Selected item does not exist.");
            if (access == "r") throw UserFileError.Denied();
        }
        else if (observed.Directory != directory) throw UserFileError.Conflict();
    }
    private static void Verify(Entry entry, Stamp? now)
    {
        if (entry.Stamp?.Id != now?.Id || !entry.Directory && entry.Stamp != now) throw UserFileError.Conflict();
        if (entry.Directory) entry.Stamp = now;
    }
    private void Writable(string path, Stamp? stamp)
    {
        if ((_workbookId is not null && _workbookId == stamp?.Id) || string.Equals(path, _workbook, StringComparison.OrdinalIgnoreCase) ||
            Leaf(path).StartsWith(".quickexcel-", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".lock", StringComparison.OrdinalIgnoreCase) ||
            stamp is { Directory: false, Links: > 1 }) throw UserFileError.Denied();
    }
    private JsonObject Grant(Entry entry)
    {
        if (_entries.Count >= 256) throw UserFileError.Limit();
        var id = Id();
        _entries.Add(id, entry);
        return new() { ["id"] = id, ["name"] = Leaf(entry.Target.Path), ["type"] = entry.Directory ? "directory" : "file", ["access"] = entry.Access };
    }
    private static Target ChildTarget(Entry parent, string name) =>
        new(Path.Combine(parent.Target.Path, Name(name)), [..parent.Target.Parents, parent.Stamp!.Id]);

    private JsonObject Child(Entry parent, string name, JsonElement options, bool directory)
    {
        var target = ChildTarget(parent, name);
        if (directory) Keys(options, "create"); else Keys(options, "fileMode", "fileAccess");
        var access = String(options, "fileAccess") ?? parent.Access;
        var mode = String(options, "fileMode") ?? "open";
        ValidateAccess(access, mode);
        if (access.Any(c => !parent.Access.Contains(c))) throw UserFileError.Denied();
        using var chain = new Chain(target.Path, target.Parents);
        var stamp = chain.Observe();
        if (stamp is null && directory && Flag(options, "create"))
        {
            if (!parent.CanWrite) throw UserFileError.Denied();
            cancellation.ThrowIfCancellationRequested();
            MakeDirectory(chain.PhysicalPath);
            stamp = chain.Observe();
        }
        CheckType(stamp, directory, mode, access);
        return Grant(new(target, stamp, directory, access, mode, parent.Delete));
    }

    private static JsonObject Info(string path, Stamp? stamp, bool directory) => new()
    {
        ["name"] = Leaf(path), ["type"] = directory ? "directory" : "file", ["exists"] = stamp is not null,
        ["size"] = directory ? null : (JsonNode?)stamp?.Length,
        ["createdAt"] = stamp?.Created is > 0 ? DateTime.FromFileTimeUtc(stamp.Created).ToString("O") : null,
        ["modifiedAt"] = stamp?.Written is > 0 ? DateTime.FromFileTimeUtc(stamp.Written).ToString("O") : null
    };

    private void Visit() { cancellation.ThrowIfCancellationRequested(); if (++_visited > 8192) throw UserFileError.Limit(); }
    private IEnumerable<JsonObject> Enumerate(Entry entry)
    {
        using var chain = new Chain(entry.Target.Path, entry.Target.Parents);
        using var pin = chain.Pin(1);
        Verify(entry, Stat(pin));
        foreach (var path in Directory.EnumerateFileSystemEntries(chain.PhysicalPath))
        {
            Visit();
            Stamp? stamp;
            try
            {
                var target = ChildTarget(entry, Path.GetFileName(path));
                using var leaf = new Chain(target.Path, target.Parents);
                stamp = leaf.Observe();
            }
            catch (UserFileError error) when (error.Code is "UnsupportedApi" or "InvalidArgument") { continue; }
            if (stamp is not null) yield return Info(path, stamp, stamp.Directory);
        }
    }

    private JsonObject Next(string id)
    {
        if (!_cursors.TryGetValue(id, out var cursor)) throw UserFileError.Invalid();
        var items = new JsonArray();
        while (items.Count < 64 && cursor.MoveNext())
        {
            JsonNode item = cursor.Current;
            items.Add(item);
        }
        var done = items.Count < 64;
        if (done) { cursor.Dispose(); _cursors.Remove(id); }
        return new() { ["items"] = items, ["done"] = done };
    }

    private void RemoveTree(Target target, bool recursive, ref int removed)
    {
        Visit();
        using var chain = new Chain(target.Path, target.Parents);
        var stamp = chain.Observe() ?? throw new UserFileError("ItemNotFound", "Selected item does not exist.");
        Writable(target.Path, stamp);
        if (stamp.Directory && recursive)
        {
            using var pin = chain.Pin(1);
            if (Stat(pin).Id != stamp.Id) throw UserFileError.Conflict();
            foreach (var child in Directory.EnumerateFileSystemEntries(chain.PhysicalPath))
            {
                var path = Path.Combine(target.Path, Name(Path.GetFileName(child)));
                RemoveTree(new(path, [..target.Parents, stamp.Id]), true, ref removed);
            }
        }
        using var handle = chain.Pin(Delete);
        var now = Stat(handle);
        if (now.Id != stamp.Id) throw UserFileError.Conflict();
        Writable(target.Path, now);
        cancellation.ThrowIfCancellationRequested();
        Remove(handle);
        removed++;
    }

    private long Reserved() => _streams.Values.Where(s => s.Data is MemoryStream).Sum(s => (long)((MemoryStream)s.Data).Capacity);

    private static bool SameTarget(Entry a, Entry b) =>
        string.Equals(a.Target.Path, b.Target.Path, StringComparison.OrdinalIgnoreCase) || a.Stamp is not null && a.Stamp.Id == b.Stamp?.Id;

    private void RequireWriterAvailable(Entry entry)
    {
        Writable(entry.Target.Path, entry.Stamp);
        if (_streams.Values.Any(s => s.Entry.CanWrite && SameTarget(s.Entry, entry))) throw new UserFileError("Locked", "File already has a writer.");
    }

    private void PublishDraft(Entry entry, Chain chain, MemoryStream data)
    {
        Verify(entry, chain.Observe());
        Writable(entry.Target.Path, entry.Stamp);
        entry.Stamp = Publish(chain, entry.Stamp, data, cancellation);
        entry.Mode = "open";
    }

    private JsonObject OpenStream(Entry entry, Chain chain, bool buffered, bool workbook = false)
    {
        if (_streams.Count >= 16) throw UserFileError.Limit();
        var keep = entry.Stamp is not null && entry.Mode != "create";
        if (workbook)
        {
            _ = QuickExcel.Storage.TypedWorkbookStorage.IsXlsq(entry.Target.Path);
            if (keep && !entry.CanRead) throw UserFileError.Denied();
        }
        var length = keep ? entry.Stamp!.Length : 0;
        if (length > 9007199254740991L || buffered && length > BufferedLimit) throw UserFileError.Limit();
        Stream data;
        if (entry.CanWrite)
        {
            RequireWriterAvailable(entry);
            if (length + Reserved() > DraftLimit) throw UserFileError.Limit();
            var draft = new MemoryStream((int)length);
            if (keep)
            {
                using var original = chain.Pin(Read);
                Verify(entry, Stat(original));
                using var source = new FileStream(original, FileAccess.Read);
                source.CopyTo(draft);
            }
            data = draft;
        }
        else
        {
            var original = chain.Pin(Read);
            try { Verify(entry, Stat(original)); data = new FileStream(original, FileAccess.Read); }
            catch { original.Dispose(); throw; }
        }
        var id = Id();
        var opened = new OpenFile(entry, data, entry.CanWrite && !keep, buffered);
        try
        {
            if (workbook)
            {
                opened.Workbook = LoadWorkbook(opened, keep);
                var model = ScriptWorkbookModel.ForFileDraft(opened.Workbook, () =>
                {
                    cancellation.ThrowIfCancellationRequested();
                    _ = request.Model!.View; // The owning Action must still be authorized.
                });
                opened.WorkbookModel = new ScriptModelAdapter(request, cancellation, model, _workbookCellBudget);
            }
            _streams.Add(id, opened);
        }
        catch { opened.Dispose(); throw; }
        return new() { ["id"] = id, ["length"] = length };
    }

    private JsonNode? StreamCall(string method, string id, JsonElement a)
    {
        if (!_streams.TryGetValue(id, out var stream)) throw new UserFileError("InvalidReference", "Stream expired.");
        var entry = stream.Entry;
        if (method == "close")
        {
            _streams.Remove(id);
            using (stream)
            {
                if (stream.Failed) throw new UserFileError("TransactionAborted", "File draft was discarded.");
                if (stream.Dirty)
                {
                    using var encoded = stream.Workbook is null ? null : EncodeWorkbook(stream.Workbook.Snapshot, entry);
                    using var chain = new Chain(entry.Target.Path, entry.Target.Parents);
                    PublishDraft(entry, chain, encoded ?? (MemoryStream)stream.Data);
                }
            }
            return null;
        }
        if (stream.Failed) throw new UserFileError("TransactionAborted", "File draft was discarded.");
        if (stream.Workbook is not null) throw UserFileError.Invalid();
        long Seek()
        {
            var position = a[1].GetInt64();
            if (position < 0 || position > stream.Data.Length) throw UserFileError.Invalid();
            stream.Data.Position = position;
            return position;
        }
        if (method == "read")
        {
            if (!entry.CanRead) throw UserFileError.Denied();
            var position = Seek();
            var count = a[2].GetInt32();
            if (count is < 0 or > Block) throw UserFileError.Invalid();
            var bytes = new byte[Math.Min(count, stream.Data.Length - position)];
            stream.Data.ReadExactly(bytes);
            return Convert.ToBase64String(bytes).Replace('+', '-').Replace('/', '_');
        }
        try
        {
            if (!entry.CanWrite) throw UserFileError.Denied();
            var position = Seek();
            var encoded = a[2].GetString()!;
            if (encoded.Length > 4 * ((Block + 2) / 3)) throw UserFileError.Limit();
            var bytes = Convert.FromBase64String(encoded.Replace('-', '+').Replace('_', '/'));
            var draft = (MemoryStream)stream.Data;
            var end = position + bytes.Length;
            if (bytes.Length > Block || end > (stream.Buffered ? BufferedLimit : DraftLimit) ||
                Reserved() + Math.Max(0, end - draft.Capacity) > DraftLimit) throw UserFileError.Limit();
            if (end > draft.Capacity) draft.Capacity = (int)end;
            draft.Write(bytes);
            stream.Dirty |= bytes.Length != 0;
            return bytes.Length;
        }
        catch { stream.Failed = true; throw; }
    }

    public void Dispose()
    {
        foreach (var cursor in _cursors.Values) cursor.Dispose();
        foreach (var stream in _streams.Values) stream.Dispose();
        _cursors.Clear(); _streams.Clear(); _entries.Clear(); _pendingCopies.Clear();
    }
}

