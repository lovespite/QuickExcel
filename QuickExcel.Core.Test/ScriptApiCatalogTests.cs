using QuickExcel.Presentation;
using QuickExcel.Scripting;

namespace QuickExcel.Core.Test;

public sealed class ScriptApiCatalogTests
{
    private static readonly IReadOnlyList<ScriptApiEntry> Entries = ScriptApiCatalog.Parse(ScriptCapabilities.Current.Json);

    [Fact]
    public void ManifestEntriesCarryDeclaredSignaturesAndCallTypes()
    {
        var values = Assert.Single(Entries, entry => entry.Path == "Range.getValues");
        Assert.Equal(ScriptApiKind.Host, values.Kind);
        Assert.Equal("getValues(): (string | number | boolean)[][]", values.Signature);
        Assert.Equal("getValues", values.HostMethod);
        Assert.Equal(2, values.Arity);
        Assert.Equal(100_000, values.Cost);

        var range = Assert.Single(Entries, entry => entry.Path == "Worksheet.getRange");
        Assert.Equal(ScriptApiKind.Local, range.Kind);
        Assert.Equal("getRange(address?: string): Range", range.Signature);
        Assert.Null(range.HostMethod);
        Assert.DoesNotContain(Entries, entry => entry.Signature?.Contains('\n') == true);
        Assert.DoesNotContain(Entries, entry => entry.Signature?.Contains("/*") == true);
    }
    [Fact]
    public void EnumValuesCarryRuntimeNumbers()
    {
        var edgeTop = Assert.Single(Entries, entry => entry.Path == "BorderIndex.edgeTop");
        Assert.Equal(ScriptApiKind.EnumValue, edgeTop.Kind);
        Assert.Equal(0, edgeTop.Value);
        Assert.Null(edgeTop.Signature);
        Assert.True(Entries.Count > 700);
    }

    [Fact]
    public void SearchMatchesMemberPathsAndNarrowsResults()
    {
        Assert.Equal(Entries.Count, ScriptApiCatalog.Filter(Entries, "  ").Count);
        Assert.Equal(Entries.Count, ScriptApiCatalog.Filter(Entries, null).Count);
        Assert.Equal("Range.getValues", Assert.Single(ScriptApiCatalog.Filter(Entries, "range.getvalues")).Path);
        var narrowed = ScriptApiCatalog.Filter(Entries, "Range");
        Assert.All(narrowed, entry => Assert.Contains("Range", entry.Path, StringComparison.OrdinalIgnoreCase));
        Assert.Contains(narrowed, entry => entry.Path == "Range.getValues");
        Assert.Contains(narrowed, entry => entry.Path == "Worksheet.getRange");
        Assert.Empty(ScriptApiCatalog.Filter(Entries, "OfficeScript"));
    }
}
