using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace QuickExcel.Storage;

/// <summary>One ledger is shared by the outer package, nested XLSX and all parsers.</summary>
internal sealed class FileReadBudget(CancellationToken cancellationToken)
{
    public const int MaxFileBytes = 64 * 1024 * 1024;
    public const long MaxExpandedBytes = 256L * 1024 * 1024;
    private long _expanded, _parsed;
    public CancellationToken CancellationToken { get; } = cancellationToken;
    public void Check() => CancellationToken.ThrowIfCancellationRequested();
    public void Expand(long count)
    {
        Check();
        if (count < 0 || count > MaxExpandedBytes - _expanded) throw new InvalidDataException("压缩包累计展开量超过 256 MiB。");
        _expanded += count;
    }
    public void Parse(long count)
    {
        Check();
        if (count < 0 || count > MaxExpandedBytes - _parsed) throw new InvalidDataException("XML/JSON 与领域对象累计解析预算超过 256 MiB。");
        _parsed += count;
    }
}

internal static class PackageIO
{
    internal static readonly UTF8Encoding Utf8 = new(false, true);
    internal static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static readonly XNamespace PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static string PathName(string name)
    {
        if (string.IsNullOrEmpty(name) || name.Length > 240 || name.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 ||
            name.Any(char.IsControl) || name.Split('/').Any(s => s.Length == 0 || s is "." or ".." || s.EndsWith('.') || s.EndsWith(' ')))
            throw new InvalidDataException($"包路径无效或存在歧义：{name}");
        return name;
    }

    public static byte[] ReadFile(string path, FileReadBudget budget)
    {
        using var input = File.OpenRead(path);
        if (input.Length > FileReadBudget.MaxFileBytes) throw new InvalidDataException("文件超过 64 MiB。");
        return Read(input, FileReadBudget.MaxFileBytes, budget, countExpanded: false);
    }

    public static Dictionary<string, byte[]> Unzip(byte[] bytes, FileReadBudget budget, int maxEntries, bool xlsq = false)
    {
        ValidateCentralDirectory(bytes, maxEntries);
        using var input = new MemoryStream(bytes, false);
        using var zip = new ZipArchive(input, ZipArchiveMode.Read);
        if (zip.Entries.Count > maxEntries) throw new InvalidDataException("压缩包条目过多。");
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long declared = 0, scriptBytes = 0;
        foreach (var entry in zip.Entries)
        {
            budget.Check();
            PathName(entry.FullName);
            if (!names.Add(entry.FullName)) throw new InvalidDataException($"重复或大小写歧义条目：{entry.FullName}");
            var limit = !xlsq ? FileReadBudget.MaxFileBytes : entry.FullName switch
            { "manifest.json" or "security/script-signature.json" => 64 * 1024, "state/document.json" => 1024 * 1024, _ when entry.FullName.StartsWith("scripts/", StringComparison.Ordinal) => 1024 * 1024, _ => FileReadBudget.MaxFileBytes };
            if (xlsq && entry.FullName.StartsWith("scripts/", StringComparison.Ordinal) && (scriptBytes += entry.Length) > 1024 * 1024)
                throw new InvalidDataException("脚本声明总量超过 1 MiB。");
            if (entry.Length > limit || entry.Length < 0 || entry.Length > FileReadBudget.MaxExpandedBytes - declared)
                throw new InvalidDataException("压缩包声明容量超过限制。");
            declared += entry.Length;
        }
        var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            var content = Read(stream, (int)entry.Length, budget, true);
            if (content.Length != entry.Length) throw new InvalidDataException($"条目长度不一致：{entry.FullName}");
            result.Add(entry.FullName, content);
        }
        return result;
    }

    // Check the bounded central directory before ZipArchive materializes its entry objects.
    private static void ValidateCentralDirectory(byte[] bytes, int maximumEntries)
    {
        static ushort U16(byte[] b, int p) => System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(b.AsSpan(p, 2));
        static uint U32(byte[] b, int p) => System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(b.AsSpan(p, 4));
        var end = -1;
        for (var position = bytes.Length - 22; position >= Math.Max(0, bytes.Length - 65557); position--)
            if (U32(bytes, position) == 0x06054b50 && position + 22 + U16(bytes, position + 20) == bytes.Length) { end = position; break; }
        if (end < 0) throw new InvalidDataException("ZIP 中央目录结束记录缺失。");
        var count = U16(bytes, end + 10);
        var size = U32(bytes, end + 12); var offset = U32(bytes, end + 16);
        if (U16(bytes, end + 4) != 0 || U16(bytes, end + 6) != 0 || U16(bytes, end + 8) != count || count > maximumEntries ||
            size > 4 * 1024 * 1024 || (long)offset + size != end)
            throw new InvalidDataException("ZIP64、分卷或超限中央目录不受支持。");
        var cursor = (int)offset;
        for (var index = 0; index < count; index++)
        {
            if (cursor + 46 > end || U32(bytes, cursor) != 0x02014b50) throw new InvalidDataException("ZIP 中央目录无效。");
            var length = 46L + U16(bytes, cursor + 28) + U16(bytes, cursor + 30) + U16(bytes, cursor + 32);
            if (cursor + length > end || U16(bytes, cursor + 28) > 960 || U32(bytes, cursor + 24) > FileReadBudget.MaxFileBytes ||
                (U16(bytes, cursor + 8) & 1) != 0 || U16(bytes, cursor + 10) is not (0 or 8))
                throw new InvalidDataException("ZIP 条目路径/容量超限或压缩方式不受支持。");
            cursor += (int)length;
        }
        if (cursor != end) throw new InvalidDataException("ZIP 条目数量与中央目录不一致。");
    }
    private static byte[] Read(Stream input, int limit, FileReadBudget budget, bool countExpanded)
    {
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        int count;
        while ((count = input.Read(buffer)) != 0)
        {
            budget.Check();
            if (count > limit - output.Length) throw new InvalidDataException("实际读取量超过条目容量限制。");
            if (countExpanded) budget.Expand(count);
            output.Write(buffer, 0, count);
        }
        return output.ToArray();
    }

    public static byte[] Required(IReadOnlyDictionary<string, byte[]> parts, string path, int maximum = FileReadBudget.MaxFileBytes)
    {
        if (!parts.TryGetValue(path, out var bytes)) throw new InvalidDataException($"文件缺少 {path}。");
        if (bytes.Length > maximum) throw new InvalidDataException($"{path} 超过容量限制。");
        return bytes;
    }

    public static XDocument Xml(byte[] bytes, FileReadBudget budget)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = FileReadBudget.MaxFileBytes, IgnoreWhitespace = false };
        using (var input = new MemoryStream(bytes, false))
        using (var reader = XmlReader.Create(input, settings))
        {
            while (reader.Read())
            {
                if (reader.Depth > 64) throw new InvalidDataException("XML 层级超过 64。");
                budget.Parse(128L + reader.Value.Length * 2L + reader.AttributeCount * 128L);
            }
        }
        using var stream = new MemoryStream(bytes, false);
        using var checkedReader = XmlReader.Create(stream, settings);
        return XDocument.Load(checkedReader);
    }

    public static JsonDocument Json(byte[] bytes, FileReadBudget budget, int maxBytes)
    {
        if (bytes.Length > maxBytes) throw new InvalidDataException("JSON 大小超过限制。");
        budget.Parse(bytes.Length * 4L);
        var document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 32 });
        try { ValidateJson(document.RootElement, budget); return document; }
        catch { document.Dispose(); throw; }
    }

    private static void ValidateJson(JsonElement value, FileReadBudget budget)
    {
        budget.Parse(64);
        if (value.ValueKind == JsonValueKind.Object)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
            {
                if (!keys.Add(property.Name)) throw new InvalidDataException($"JSON 重复属性：{property.Name}");
                ValidateJson(property.Value, budget);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
            foreach (var item in value.EnumerateArray()) ValidateJson(item, budget);
    }

    public static byte[] XmlBytes(XDocument document)
    {
        using var output = new LimitedMemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false })) document.Save(writer);
        return output.ToArray();
    }
    public static byte[] Zip(IReadOnlyDictionary<string, byte[]> parts)
    {
        using var output = new LimitedMemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var (name, content) in parts)
            {
                using var stream = zip.CreateEntry(PathName(name), CompressionLevel.Optimal).Open();
                stream.Write(content);
            }
        return output.ToArray();
    }

    public static string RelationshipPart(string source)
    {
        if (source.Length == 0) return "_rels/.rels";
        var slash = source.LastIndexOf('/');
        return source[..(slash + 1)] + "_rels/" + source[(slash + 1)..] + ".rels";
    }
    public static string Resolve(string source, string target)
    {
        if (string.IsNullOrEmpty(target) || target.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 || target.StartsWith("//", StringComparison.Ordinal))
            throw new InvalidDataException($"无效的内部关系目标：{target}");
        var pieces = new List<string>();
        if (!target.StartsWith('/')) pieces.AddRange(source.Split('/').SkipLast(1));
        foreach (var piece in target.TrimStart('/').Split('/'))
        {
            if (piece == "..")
            {
                if (pieces.Count == 0) throw new InvalidDataException("关系路径越过包根目录。");
                pieces.RemoveAt(pieces.Count - 1);
            }
            else if (piece != ".") pieces.Add(piece);
        }
        return PathName(string.Join('/', pieces));
    }

    public static void AtomicWrite(string path, byte[] bytes, Action<string> validate, CancellationToken cancellationToken, QuickExcel.Core.FileVersion? expected = null)
    {
        var fullPath = Path.GetFullPath(path);
        var temp = Path.Combine(Path.GetDirectoryName(fullPath)!, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(bytes); stream.Flush(true); }
            validate(temp);
            cancellationToken.ThrowIfCancellationRequested();
            if (expected is not null) FileVersions.Require(fullPath, expected);
            if (File.Exists(fullPath)) File.Replace(temp, fullPath, null);
            else File.Move(temp, fullPath);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }

    private sealed class LimitedMemoryStream : MemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count)
        { Limit(count); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer)
        { Limit(buffer.Length); base.Write(buffer); }
        public override void WriteByte(byte value) { Limit(1); base.WriteByte(value); }
        private void Limit(int count)
        {
            if (Position + count > FileReadBudget.MaxFileBytes) throw new InvalidDataException("写入文件或 XML 超过 64 MiB。");
        }
    }
}
