using System.Globalization;
using System.Text.Json;
using QuickExcel.Core;

namespace QuickExcel.Presentation;

public sealed record DisplayLanguage(string Tag, string NativeName)
{
    public override string ToString() => NativeName;
}

/// <summary>Looks up application-owned UI strings. Workbook values and format codes do not pass through this catalog.</summary>
public static class UiText
{
    public const string Chinese = "zh-CN";
    public const string English = "en-US";
    public static IReadOnlyList<DisplayLanguage> Languages { get; } = Array.AsReadOnly(new[]
    {
        new DisplayLanguage(Chinese, "简体中文"),
        new DisplayLanguage(English, "English")
    });
    public static string Language { get; private set; } = Chinese;

    public static DisplayLanguage Resolve(string? tag) =>
        Languages.FirstOrDefault(language => language.Tag == tag) ?? Languages[0];

    // Select once at startup, before views and their cached option lists are constructed.
    public static void Initialize(string language)
    {
        Language = Resolve(language).Tag;
        IReadOnlyDictionary<string, string>? catalog = null;
        if (Language != Chinese)
        {
            using var stream = typeof(UiText).Assembly.GetManifestResourceStream($"QuickExcel.Localization.{Language}.json")
                ?? throw new InvalidDataException($"Missing UI catalog for {Language}.");
            catalog = JsonSerializer.Deserialize<Dictionary<string, string>>(stream)
                ?? throw new InvalidDataException($"Empty UI catalog for {Language}.");
        }
        MessageText.Configure(catalog);
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(Language);
    }

    public static string Get(string source) => MessageText.Get(source);

    // Translate the template before inserting arguments. A sheet named "设置" stays "设置" in every language.
    public static string Format(FormattableString message) =>
        MessageText.Format(message);
}
