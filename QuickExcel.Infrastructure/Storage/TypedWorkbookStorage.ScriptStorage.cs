using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    private const string StoragePrefix = "state/local-storage/";
    private static bool IsStoragePath(string name) => name.TrimEnd('/').Equals(StoragePrefix.TrimEnd('/'), StringComparison.OrdinalIgnoreCase) || name.StartsWith(StoragePrefix, StringComparison.OrdinalIgnoreCase);

    private static ScriptStorage ReadStoragePart(byte[] bytes, ScriptStorage storage, FileReadBudget budget)
    {
        if (bytes.Length < 8) throw new InvalidDataException("localStorage 数据头不完整。");
        using var input = new MemoryStream(bytes, false);
        using var reader = new BinaryReader(input, Utf8);
        if (reader.ReadInt32() != 1) throw new InvalidDataException("不支持的 localStorage 数据版本。");
        var count = reader.ReadInt32();
        if (count <= 0 || count > (bytes.Length - 8) / 8) throw new InvalidDataException("localStorage 条目数量无效。");
        string ReadString(int maximum)
        {
            var length = reader.ReadInt32();
            if (length < 0 || length > maximum || length > input.Length - input.Position) throw new InvalidDataException("localStorage 条目长度无效。");
            budget.Parse(2L * length + 64);
            return Utf8.GetString(reader.ReadBytes(length));
        }
        for (var i = 0; i < count; i++)
        {
            budget.Check();
            var key = ReadString(ScriptStorage.MaxEntryBytes);
            var value = ReadString(ScriptStorage.MaxEntryBytes - Utf8.GetByteCount(key));
            if (storage.Entries.ContainsKey(key)) throw new InvalidDataException("localStorage 包含重复键。");
            storage = storage.Set(key, value);
        }
        if (input.Position != input.Length) throw new InvalidDataException("localStorage 含多余数据。");
        return storage;
    }

    private static void WriteStorageParts(ScriptStorage storage, Dictionary<string, byte[]> parts, JsonObject manifest, CancellationToken cancellationToken)
    {
        if (storage.Entries.Count == 0) { manifest.Remove("localStorage"); return; }
        manifest["requiredFeatures"]!.AsArray().Add("script-storage");
        var references = new JsonArray();
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Utf8, true);
        int count = 0, index = 0;
        void Begin() { output.SetLength(0); output.Position = 0; writer.Write(1); writer.Write(0); count = 0; }
        void Flush()
        {
            writer.Flush(); output.Position = 4; writer.Write(count); writer.Flush();
            var bytes = output.ToArray();
            var path = StoragePrefix + (index++).ToString("D6", System.Globalization.CultureInfo.InvariantCulture) + ".bin";
            parts.Add(path, bytes);
            references.Add(new JsonObject { ["path"] = path, ["sha256"] = Hash(bytes) });
        }
        Begin();
        foreach (var (key, value) in storage.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var keyBytes = Utf8.GetBytes(key); var valueBytes = Utf8.GetBytes(value);
            if (count > 0 && output.Length + keyBytes.Length + valueBytes.Length + 8 > 4 * 1024 * 1024) { Flush(); Begin(); }
            writer.Write(keyBytes.Length); writer.Write(keyBytes); writer.Write(valueBytes.Length); writer.Write(valueBytes); count++;
        }
        Flush(); manifest["localStorage"] = references;
    }
}
