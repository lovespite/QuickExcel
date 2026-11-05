using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Security;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SourceMapOptionsTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-source-maps-" + Guid.NewGuid().ToString("N"));
    private readonly string _priorMode = ScriptCompilation.SourceMapMode;
    public SourceMapOptionsTests() => Directory.CreateDirectory(_folder);
    public void Dispose() { ScriptCompilation.SourceMapMode = _priorMode; Directory.Delete(_folder, true); }
    private static CalculationStamp Stamp => new(Guid.NewGuid(), 0, 0, 0);
    private static ScriptModuleSet Modules() => new(new Dictionary<string, ScriptModule>
    {
        ["main"] = new("scripts/main.ts", "enum E { A, B }\n\nhost.functions.register('FAIL',()=>{\n  const n:number=E.B;\n  throw new Error('mapped '+n);\n});", ScriptLanguage.TypeScript)
    }, ["main"], scriptApiVersion: 3);

    [Theory]
    [InlineData("none")]
    [InlineData("compact")]
    [InlineData("full")]
    public async Task CompilerEmitsSelectedMapsAndRuntimeUsesCorrectErrorPositions(string mode)
    {
        ScriptCompilation.SourceMapMode = mode;
        var modules = Modules();
        var result = (await ScriptCompilation.PrepareAsync(modules, Stamp, default))!;
        var output = Assert.Single(result.Result.Outputs);
        if (mode == "none") { Assert.Empty(output.SourceMap); Assert.DoesNotContain("sourceMappingURL", output.Javascript); }
        else
        {
            using var map = JsonDocument.Parse(output.SourceMap);
            Assert.Equal(mode == "full", map.RootElement.TryGetProperty("sourcesContent", out var source));
            if (mode == "full") Assert.Equal(modules.PrimarySource, source[0].GetString());
            Assert.NotEmpty(map.RootElement.GetProperty("mappings").GetString()!);
        }
        using var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, Stamp, TimeSpan.FromSeconds(5), default));
        var error = Assert.Throws<InvalidOperationException>(() => runtime.Invoke("FAIL", []));
        if (mode == "none") Assert.Contains(UiText.Get(" (生成 JS 位置)"), error.Message);
        else Assert.Contains("main:5:", error.Message);
        Assert.Equal(1, ScriptCompilation.ProcessLaunchCount(modules));
    }

    [Fact]
    public async Task EachModeHasAnIsolatedCacheIdentityAndReopensWithoutResigningOrCompiling()
    {
        ScriptCompilation.SourceMapMode = "full";
        var modules = Modules();
        using var certificate = ScriptSignatures.CreateSelfSigned("Source Map cache options");
        var signature = await ScriptSignatures.SignAsync(modules, certificate, Stamp);
        modules = modules.WithSignature(signature);
        var hashes = new HashSet<string>(StringComparer.Ordinal);
        var storage = new TypedWorkbookStorage();
        foreach (var mode in new[] { "full", "compact", "none" })
        {
            ScriptCompilation.SourceMapMode = mode;
            if (mode != "full") Assert.True(ScriptCompilation.IsCacheDirty(modules));
            var compiled = (await ScriptCompilation.PrepareAsync(modules, Stamp, default))!;
            Assert.True(hashes.Add(compiled.SourceSha256));
            var parts = ScriptCompilationCache.Export(modules)!;
            using var metadata = JsonDocument.Parse(parts.Metadata);
            Assert.Equal(mode, metadata.RootElement.GetProperty("compilation").GetProperty("sourceMapMode").GetString());
            if (mode == "none") Assert.All(metadata.RootElement.GetProperty("compilation").GetProperty("outputs").EnumerateArray(), o => Assert.Equal(0, o.GetProperty("sourceMapBytes").GetInt32()));
            var path = Path.Combine(_folder, mode + ".xlsq");
            storage.Save(path, Book(Sheet("S")).WithScriptModules(modules));
            var reopened = storage.Open(path).ScriptModules;
            Assert.Equal(signature, reopened.Signature);
            Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(reopened).Status);
            Assert.False(ScriptCompilation.IsCacheDirty(reopened));
            Assert.Equal(compiled.OutputSha256, (await ScriptCompilation.PrepareAsync(reopened, Stamp, default))!.OutputSha256);
            Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(reopened));
        }
        Assert.Equal(3, ScriptCompilation.ProcessLaunchCount(modules));
    }

    [Fact]
    public void LocalSettingsPersistAllModesAndOldSettingsDefaultToCompact()
    {
        var path = Path.Combine(_folder, "settings.json");
        File.WriteAllText(path, "{\"Zoom\":125,\"Language\":\"en-US\"}");
        var settings = new DesktopSettings(path); settings.Load();
        Assert.Equal(DesktopSourceMapMode.Compact, settings.Value.ScriptSourceMap);
        foreach (var mode in Enum.GetValues<DesktopSourceMapMode>())
        {
            settings.Save(settings.Value with { ScriptSourceMap = mode });
            Assert.Null(settings.Error);
            var next = new DesktopSettings(path); next.Load();
            Assert.Equal(mode, next.Value.ScriptSourceMap);
            Assert.Equal(125, next.Value.Zoom); Assert.Equal(UiText.English, next.Value.Language);
        }
        File.WriteAllText(path, "{\"ScriptSourceMap\":99}"); settings.Load();
        Assert.NotNull(settings.Error); Assert.Equal(DesktopSourceMapMode.Compact, settings.Value.ScriptSourceMap);
        Assert.Throws<ArgumentOutOfRangeException>(() => ScriptCompilation.SourceMapMode = "invalid");
        var request = ScriptCompilation.Request(Modules()) with { SourceMapMode = "invalid" };
        request = request with { Sha256 = request.ComputeDigest() };
        Assert.Throws<ProtocolException>(request.Validate);
    }

    [Fact]
    public void SourceMapLabelsAndExplanationsHaveEnglishTranslations()
    {
        using var stream = typeof(UiText).Assembly.GetManifestResourceStream("QuickExcel.Localization.en-US.json")!;
        using var catalog = JsonDocument.Parse(stream, new() { AllowDuplicateProperties = false });
        Assert.Equal("None", catalog.RootElement.GetProperty("无").GetString());
        Assert.Equal("Compact", catalog.RootElement.GetProperty("精简").GetString());
        Assert.Equal("Full", catalog.RootElement.GetProperty("完整").GetString());
        foreach (var key in new[] { "Source Map（源码映射）", "Source Map 设置已更新，编译缓存已失效；请重新应用脚本。",
            "Source Map 设置已保存，关闭设置后生效；请重新应用已启用的脚本。",
            "仅保存在本机。更改后当前脚本会停用；下次应用或签名时使用新设置，已有签名保持有效。", " (生成 JS 位置)" })
            Assert.DoesNotMatch("[\\u4e00-\\u9fff]", catalog.RootElement.GetProperty(key).GetString()!);
    }
}
