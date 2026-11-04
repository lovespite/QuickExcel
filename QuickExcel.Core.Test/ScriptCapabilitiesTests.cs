using QuickExcel.Scripting;

namespace QuickExcel.Core.Test;

public sealed class ScriptCapabilitiesTests
{
    [Fact]
    public void EmbeddedManifestExposesOnlyImplementedHostMembers()
    {
        var capabilities = ScriptCapabilities.Current;
        Assert.Equal(ScriptCapabilities.ApiVersion, 3);
        Assert.Equal(64, capabilities.SourceSha256.Length);
        Assert.True(capabilities.AllowsHost("active", 0));
        Assert.True(capabilities.AllowsHost("sheet", 1));
        Assert.True(capabilities.AllowsHost("getValues", 2));
        Assert.True(capabilities.AllowsHost("setValues", 3));
        Assert.True(capabilities.AllowsHost("clear", 2));
        Assert.False(capabilities.AllowsHost("getValues", 1));
        Assert.False(capabilities.AllowsHost("getApplication", 0));
        Assert.Equal(100_000, capabilities.MaxCost("getValues"));
        Assert.Equal(20_000, capabilities.MaxCost("setValues"));
    }

    [Fact]
    public void ManifestIsInjectedIntoTheProxyWithoutEscaping()
    {
        Assert.Contains("\"apiVersion\": 3", ScriptCapabilities.Current.Json);
        Assert.Contains("\"RangeAreas\"", ScriptCapabilities.Current.Json);
        Assert.DoesNotContain("OfficeScript", ScriptCapabilities.Current.Json);
    }
}
