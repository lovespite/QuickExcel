using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text.Json;
using QuickExcel.Core;

namespace QuickExcel.Presentation;

/// <summary>Bounded, atomic local library. An optimistic fingerprint prevents stale windows overwriting newer edits.</summary>
public sealed class QuickCellStyleStore(string path)
{
    private sealed record Library(int Version, QuickCellStyle[] Styles);
    private string? _fingerprint;
    public ImmutableArray<QuickCellStyle> Styles { get; private set; } = [];
    public string? Error { get; private set; }
    public const int MaximumStyles = 128;
    private const int MaximumBytes = 1024 * 1024;
    private byte[] Read()
    {
        if (!File.Exists(path)) return [];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > MaximumBytes) throw new InvalidDataException(UiText.Get("自定义样式文件过大。"));
        var bytes = new byte[checked((int)stream.Length)]; stream.ReadExactly(bytes); return bytes;
    }
    private static string Fingerprint(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    public void Load()
    {
        try
        {
            var bytes = Read();
            var library = bytes.Length == 0 ? new Library(1, []) : JsonSerializer.Deserialize<Library>(bytes);
            if (library is not { Version: 1, Styles: not null }) throw new InvalidDataException(UiText.Get("无效的自定义样式文件。"));
            Validate(library.Styles); Styles = library.Styles.ToImmutableArray(); _fingerprint = Fingerprint(bytes); Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or ArgumentException or InvalidOperationException)
        { Styles = []; _fingerprint = null; Error = ex.Message; }
    }
    public void Save(IEnumerable<QuickCellStyle> styles)
    {
        var entries = styles.Take(MaximumStyles + 1).ToArray(); Validate(entries);
        if (_fingerprint is null) throw new InvalidOperationException(UiText.Get("请先重新加载自定义样式，原文件保持不变。"));
        var bytes = JsonSerializer.SerializeToUtf8Bytes(new Library(1, entries));
        if (bytes.Length > MaximumBytes) throw new InvalidDataException(UiText.Get("自定义样式文件过大。"));
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using var lease = new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.Write, FileShare.None);
        if (Fingerprint(Read()) != _fingerprint) throw new InvalidOperationException(UiText.Get("自定义样式已在其他窗口更改，请重新打开样式库。"));
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(bytes); stream.Flush(true); }
            File.Move(temporary, path, true); Styles = entries.ToImmutableArray(); _fingerprint = Fingerprint(bytes); Error = null;
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
    public static void Validate(IReadOnlyList<QuickCellStyle> styles)
    {
        if (styles.Count > MaximumStyles) throw new ArgumentException(UiText.Format($"最多保存 {MaximumStyles} 个自定义样式。"));
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var style in styles)
        {
            if (style is null || !Guid.TryParseExact(style.Id, "N", out _) || !ids.Add(style.Id) || string.IsNullOrWhiteSpace(style.Name) || style.Name.Length > 64 || style.Name.Any(char.IsControl) || style.Name != style.Name.Trim() || !names.Add(style.Name))
                throw new ArgumentException(UiText.Get("样式名称不能为空、重复或超过 64 个字符。"));
            if (QuickCellStyles.BuiltIn.Any(s => string.Equals(s.Name, style.Name, StringComparison.OrdinalIgnoreCase) || string.Equals(QuickCellStyles.DisplayName(s), style.Name, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException(UiText.Get("此名称已用于内置样式，请使用其他名称。"));
            if (style.Format is null || style.Parts == CellStyleParts.None || (style.Parts & ~CellStyleParts.All) != 0 || style.Format.NumberFormatCode?.Length > 1024 || style.Group != "自定义")
                throw new ArgumentException(UiText.Get("请选择至少一个有效的样式类别。"));
            // Validate every category, including those currently unchecked, through the actual formatting command.
            _ = QuickCellStyle.EditFormat(new(), (style with { Parts = CellStyleParts.All }).Patch);
        }
    }
}
