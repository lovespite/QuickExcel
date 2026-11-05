using System.Text.Json;
using System.Text.RegularExpressions;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class LocalizationCatalogTests
{
    [Fact]
    public void EnglishCatalogContainsTranslatedTextAndPreservesFormatArguments()
    {
        using var stream = typeof(UiText).Assembly.GetManifestResourceStream("QuickExcel.Localization.en-US.json");
        Assert.NotNull(stream);
        using var document = JsonDocument.Parse(stream);
        var entries = document.RootElement.EnumerateObject().ToArray();
        Assert.Equal(entries.Length, entries.Select(entry => entry.Name).Distinct().Count());
        foreach (var entry in entries)
        {
            var translation = entry.Value.GetString();
            Assert.False(string.IsNullOrWhiteSpace(translation), entry.Name);
            Assert.DoesNotMatch("[\u3400-\u9fff]", translation);
            Assert.Equal(Arguments(entry.Name), Arguments(translation));
        }
        Assert.Equal("Settings", document.RootElement.GetProperty("设置").GetString());
        Assert.Equal("Off", document.RootElement.GetProperty("关").GetString());
    }

    [Theory]
    [InlineData(null, UiText.Chinese)]
    [InlineData("unknown", UiText.Chinese)]
    [InlineData(UiText.Chinese, UiText.Chinese)]
    [InlineData(UiText.English, UiText.English)]
    public void LanguageTagsHaveStableFallback(string? tag, string expected) =>
        Assert.Equal(expected, UiText.Resolve(tag).Tag);

    private static string[] Arguments(string template) =>
        Regex.Matches(template, @"\{\d+(?:,[^}:]+)?(?::[^}]+)?\}")
            .Select(match => match.Value).Order(StringComparer.Ordinal).ToArray();
}
