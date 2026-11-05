using System.Collections.Frozen;
using System.Globalization;

namespace QuickExcel.Core;

/// <summary>
/// Host-supplied translations for application-owned messages.
/// The default preserves source text for independent Core and worker callers.
/// Catalogs contain templates only. User values, formula codes and identifiers are never translated.
/// </summary>
public static class MessageText
{
    private static FrozenDictionary<string, string> _translations = FrozenDictionary<string, string>.Empty;

    public static void Configure(IReadOnlyDictionary<string, string>? translations) =>
        Volatile.Write(ref _translations, translations?.ToFrozenDictionary(StringComparer.Ordinal)
            ?? FrozenDictionary<string, string>.Empty);

    public static string Get(string source) => Volatile.Read(ref _translations).GetValueOrDefault(source) ?? source;

    public static string Format(FormattableString message) =>
        string.Format(CultureInfo.CurrentCulture, Get(message.Format), message.GetArguments());
}
