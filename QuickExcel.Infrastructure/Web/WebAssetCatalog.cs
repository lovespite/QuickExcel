using QuickExcel.Core;
using System.Security.Cryptography;
using System.Text.Json;

namespace QuickExcel.Infrastructure.Web;

/// <summary>Verified immutable package resources; no workbook directory is exposed to a page.</summary>
public sealed class WebAssetCatalog
{
    internal IReadOnlyDictionary<string, byte[]> Files { get; }
    public string ManifestSha256 { get; }
    public IReadOnlyDictionary<string, string> Packages { get; }
    public WebAssetCatalog(string directory, bool allowTestResources = false)
    {
        directory = Path.GetFullPath(directory);
        var manifestPath = Path.Combine(directory, "manifest.json");
        if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0 || (File.GetAttributes(manifestPath) & FileAttributes.ReparsePoint) != 0)
            throw new IOException(MessageText.Get("Web 资源目录和清单不能是重解析点。"));
        var manifestBytes = ReadBounded(manifestPath, 2 * 1024 * 1024); ManifestSha256 = Convert.ToHexString(SHA256.HashData(manifestBytes));
        using var manifest = JsonDocument.Parse(manifestBytes, new() { MaxDepth = 12 });
        if (manifest.RootElement.GetProperty("schemaVersion").GetInt32() != 1) throw new InvalidDataException(MessageText.Get("未知 Web 资源清单版本。"));
        if (!allowTestResources && (manifest.RootElement.TryGetProperty("testBuild", out var test) && test.GetBoolean() || manifest.RootElement.TryGetProperty("probeBuild", out var probe) && probe.GetBoolean()))
            throw new InvalidDataException(MessageText.Get("正式工作区不能加载探针或故障注入资源，请重新构建正式 Web 资源。"));
        Packages = manifest.RootElement.GetProperty("packages").Deserialize<Dictionary<string, string>>()!;
        var files = new Dictionary<string, byte[]>(StringComparer.Ordinal); long total = 0;
        foreach (var item in manifest.RootElement.GetProperty("files").EnumerateArray())
        {
            var name = item.GetProperty("name").GetString()!;
            if (files.Count >= 4096 || !ValidName(name)) throw new InvalidDataException(MessageText.Get("Web 资源名称或数量无效。"));
            var path = directory;
            if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new IOException(MessageText.Get("Web 资源目录不能是重解析点。"));
            foreach (var part in name.Split('/'))
            { path = Path.Combine(path, part); if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new IOException(MessageText.Get("Web 资源不能经过重解析点。")); }
            var expected = item.GetProperty("size").GetInt32();
            if (expected is < 0 or > 32 * 1024 * 1024 || (total += expected) > 256L * 1024 * 1024)
                throw new InvalidDataException(MessageText.Get("Web 资源容量超限：") + name);
            var actual = new FileInfo(path).Length;
            if (actual != expected)
                throw new InvalidDataException(MessageText.Format($"Web 资源大小与清单不匹配：{name}（预期 {expected} 字节，实际 {actual} 字节）。请重新构建或安装完整应用包。"));
            var bytes = ReadBounded(path, expected);
            if (bytes.Length != expected || !Convert.ToHexString(SHA256.HashData(bytes)).Equals(item.GetProperty("sha256").GetString(), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(MessageText.Get("Web 资源 SHA-256 不匹配：") + name);
            files.Add(name, bytes);
        }
        Files = files;
    }
    private static byte[] ReadBounded(string path, int maximum)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > maximum) throw new InvalidDataException(MessageText.Get("Web 资源超过读取预算。"));
        var bytes = new byte[checked((int)stream.Length)]; stream.ReadExactly(bytes); return bytes;
    }
    private static bool ValidName(string? name) => name is { Length: > 0 and <= 512 } &&
        name.All(c => char.IsAsciiLetterOrDigit(c) || c is '/' or '.' or '_' or '-') && !name.Split('/').Any(p => p is "" or "." or "..");
}
