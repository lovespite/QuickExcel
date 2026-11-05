using System.Text.Json;
using QuickExcel.Core;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private static bool StorageCall(string method, int count) => method is "getStorage" or "setStorage" && count == 3;

    private object? CallStorage(string method, JsonElement[] args)
    {
        var session = args[0].GetString() switch
        {
            "local" => false, "session" => true, _ => throw new ArgumentException("未知存储类型。")
        };
        var operation = args[1].GetString();
        var model = request.Model!;
        if (method == "setStorage")
        {
            if (operation == "set" && args[2].ValueKind == JsonValueKind.Array && args[2].GetArrayLength() == 2)
                model.WriteStorage(session, operation, args[2][0].GetString(), args[2][1].GetString());
            else if (operation == "remove" && args[2].ValueKind == JsonValueKind.String)
                model.WriteStorage(session, operation, args[2].GetString());
            else if (operation == "clear" && args[2].ValueKind == JsonValueKind.Null)
                model.WriteStorage(session, operation);
            else throw new ArgumentException("存储写入参数无效。");
            return null;
        }
        var storage = model.ReadStorage(session);
        if (operation == "get") return storage.Get(args[2].GetString()!);
        if (operation is not ("keys" or "values" or "entries")) throw new ArgumentException("未知存储读取操作。");
        var after = args[2].ValueKind == JsonValueKind.Null ? null : args[2].GetString();
        // Page snapshots across the bridge: a 100 MiB store never becomes a single IPC frame.
        // Even worst-case JSON escaping of a page fits the existing 16 MiB frame budget.
        var items = new List<object>();
        long bytes = 0;
        string? next = null;
        foreach (var pair in storage.Entries)
        {
            cancellation.ThrowIfCancellationRequested();
            if (after is not null && StringComparer.Ordinal.Compare(pair.Key, after) <= 0) continue;
            var size = (long)pair.Key.Length + pair.Value.Length + 32;
            if (items.Count > 0 && (bytes + size > ScriptStorage.MaxEntryBytes || items.Count == 1024))
                return new { items, next, done = false };
            items.Add(operation switch { "keys" => pair.Key, "values" => pair.Value, _ => (object)new[] { pair.Key, pair.Value } });
            bytes += size; next = pair.Key;
        }
        return new { items, next, done = true };
    }
}
