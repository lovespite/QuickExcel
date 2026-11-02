using System.Globalization;
using System.Buffers.Binary;
using System.Text;
using SkiaSharp;
using Topten.RichTextKit;

namespace QuickExcel.Rendering.Skia;

/// <summary>Renderer-owned faces; explicit fonts precede system discovery, which can be disabled for fixtures.</summary>
public sealed class SkiaFonts : FontMapper, IDisposable
{
    private readonly List<SKTypeface> _faces = [];
    private readonly Dictionary<string, SKTypeface> _requests = [];
    private readonly Dictionary<SKTypeface, int> _weights = [];
    private readonly Dictionary<(SKTypeface Face, int Scalar), bool> _coverage = [];
    private readonly bool _systemFonts;
    private readonly int _limit;
    public SkiaFonts(IEnumerable<string>? fontFiles = null, bool systemFonts = true, int limit = 128)
    {
        _systemFonts = systemFonts; _limit = limit;
        try
        {
        foreach (var path in fontFiles ?? [])
            _faces.Add(SKTypeface.FromFile(path) ?? throw new InvalidDataException($"Cannot load font: {path}"));
        if (_faces.Count == 0 && systemFonts) _faces.Add(OwnSystemFace(SKTypeface.CreateDefault()) ?? throw new InvalidOperationException("No default system font is available."));
        if (_faces.Count > limit) throw new ArgumentException("Font face budget exceeded.");
        if (_faces.Count == 0 && !systemFonts) throw new ArgumentException("Explicit fonts required when system fonts are disabled.");
        }
        catch { Dispose(); throw; }
    }

    internal SKTypeface Resolve(string text, FontRequest request, Action<string, string> diagnostic)
    {
        var eastAsian = text.EnumerateRunes().Any(r => r.Value is >= 0x2e80 and <= 0x9fff or >= 0xf900 and <= 0xfaff or >= 0x20000 and <= 0x3134f);
        var family = (eastAsian ? request.EastAsianFamily : null) ?? request.Family;
        var key = $"{family}|{request.Bold}|{request.Italic}";
        using var style = new SKFontStyle(request.Bold ? 700 : 400, 5, request.Italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);
        bool HasGlyph(SKTypeface face, int scalar)
        {
            if (_coverage.TryGetValue((face, scalar), out var cached)) return cached;
            using var font = new SKFont(face); var has = font.ContainsGlyph(scalar);
            if (_coverage.Count < 16384) _coverage[(face, scalar)] = has;
            return has;
        }
        bool Covers(SKTypeface face) => text.EnumerateRunes().All(r => Rune.IsControl(r) || Rune.GetUnicodeCategory(r) == UnicodeCategory.Format || HasGlyph(face, r.Value));
        if (!_requests.TryGetValue(key, out var requested))
        {
            requested = _faces.FirstOrDefault(f => string.Equals(f.FamilyName, family, StringComparison.OrdinalIgnoreCase) && (Weight(f) >= 600) == request.Bold && f.IsItalic == request.Italic);
            if (requested is null && _systemFonts && _faces.Count < _limit)
            {
                requested = OwnSystemFace(SKFontManager.Default.MatchFamily(family, style));
                if (requested is not null) _faces.Add(requested);
            }
            requested ??= _faces.FirstOrDefault(f => string.Equals(f.FamilyName, family, StringComparison.OrdinalIgnoreCase));
            requested ??= _faces.FirstOrDefault(f => f.FamilyName == "Noto Sans" && Weight(f) == 400 && !f.IsItalic) ?? _faces.FirstOrDefault();
            if (requested is null) throw new InvalidOperationException("No usable font is available.");
            if (_requests.Count < 4096) _requests[key] = requested;
        }
        var resolved = Covers(requested) ? requested : _faces.FirstOrDefault(Covers);
        if (resolved is null && _systemFonts && _faces.Count < _limit)
        {
            var missing = text.EnumerateRunes().FirstOrDefault(r => !HasGlyph(requested, r.Value));
            var fallback = OwnSystemFace(SKFontManager.Default.MatchCharacter(family, style, ["zh-Hans", "en"], missing.Value));
            if (fallback is not null) { _faces.Add(fallback); if (Covers(fallback)) resolved = fallback; }
        }
        if (resolved is null)
        {
            resolved = requested;
            diagnostic("missing-glyph", $"{resolved.FamilyName}: {string.Join(" ", text.EnumerateRunes().Where(r => !Rune.IsControl(r) && !HasGlyph(resolved, r.Value)).Select(r => $"U+{r.Value:X}"))}");
        }
        if (family is not null && !string.Equals(family, resolved.FamilyName, StringComparison.OrdinalIgnoreCase))
            diagnostic("font-substitution", $"{family} → {resolved.FamilyName}");
        return resolved;
    }

    internal string Key(SKTypeface face) => _faces.IndexOf(face).ToString(CultureInfo.InvariantCulture);
    private static SKTypeface? OwnSystemFace(SKTypeface? matched)
    {
        if (matched is null) return null;
        using (matched)
        using (var stream = matched.OpenStream(out var index))
            return stream is null ? null : SKTypeface.FromStream(stream, index);
    }
    internal int Weight(SKTypeface face)
    {
        if (_weights.TryGetValue(face, out var weight)) return weight;
        // DirectWrite reports some static Noto fonts' STAT default (100) instead of their actual OS/2 weight.
        var table = face.GetTableData(0x4f532f32);
        weight = table is { Length: >= 6 } ? BinaryPrimitives.ReadUInt16BigEndian(table.AsSpan(4, 2)) : face.FontWeight;
        _weights[face] = weight; return weight;
    }
    public override SKTypeface TypefaceFromStyle(IStyle style, bool ignoreFontVariants) => _faces[int.Parse(style.FontFamily, CultureInfo.InvariantCulture)];
    public void Dispose() { SkiaTextLayout.ReleaseFonts(_faces.Distinct()); _faces.Clear(); _requests.Clear(); _weights.Clear(); _coverage.Clear(); }
}
