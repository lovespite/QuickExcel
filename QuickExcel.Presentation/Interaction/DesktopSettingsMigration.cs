using System.Security.Cryptography;
using System.Text.Json;

namespace QuickExcel.Presentation;

/// <summary>One-way copy of compatible preferences, never shared writes or automatic trust migration.</summary>
public static class DesktopSettingsMigration
{
    public static string? ImportOnce(string legacyDirectory, string destinationDirectory)
    {
        var target = Path.GetFullPath(destinationDirectory); var legacy = Path.GetFullPath(legacyDirectory);
        if (target.Equals(legacy, StringComparison.OrdinalIgnoreCase)) return "新旧设置目录必须分离。";
        var marker = Path.Combine(target, "migration-v1.json");
        try
        {
            Directory.CreateDirectory(target);
            if ((File.GetAttributes(target) & FileAttributes.ReparsePoint) != 0) throw new IOException("设置目录不能是重解析点。");
            if (File.Exists(marker)) return null;
            using var lease = new FileStream(Path.Combine(target, "migration.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            if (File.Exists(marker)) return null;
            var imported = new List<object>();
            var backup = Path.Combine(target, "migration-backup-" + Guid.NewGuid().ToString("N"));
            foreach (var (name, maximum) in new[] { ("settings.json", 4096), ("recent-files.json", 65536) })
            {
                var source = Path.Combine(legacy, name); var destination = Path.Combine(target, name);
                if (!File.Exists(source) || File.Exists(destination)) continue;
                if ((File.GetAttributes(source) & FileAttributes.ReparsePoint) != 0 || (File.GetAttributes(legacy) & FileAttributes.ReparsePoint) != 0) throw new IOException("旧设置不能经过重解析点。");
                byte[] bytes;
                using (var stream = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                { if (stream.Length > maximum) throw new InvalidDataException("旧设置超过容量限制。"); bytes = new byte[checked((int)stream.Length)]; stream.ReadExactly(bytes); }
                using var json = JsonDocument.Parse(bytes);
                if (name == "settings.json")
                {
                    var value = json.Deserialize<DesktopPreferences>();
                    if (value is null || !value.IsValid) throw new InvalidDataException("旧界面设置无效。");
                }
                else if (json.Deserialize<string[]>() is not { } paths || paths.Any(p => p is null)) throw new InvalidDataException("旧最近文件记录无效。");
                Directory.CreateDirectory(backup); WriteNew(Path.Combine(backup, name), bytes);
                WriteNew(destination, bytes);
                imported.Add(new { name, sha256 = Convert.ToHexString(SHA256.HashData(bytes)) });
            }
            WriteNew(marker, JsonSerializer.SerializeToUtf8Bytes(new { schemaVersion = 1, imported, createdUtc = DateTime.UtcNow }));
            return null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or ArgumentException)
        { return "旧设置未完全导入；原设置保留，可继续使用独立设置。" + ex.Message; }
    }
    private static void WriteNew(string path, byte[] bytes)
    {
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(bytes); stream.Flush(true); }
            File.Move(temporary, path, false);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
