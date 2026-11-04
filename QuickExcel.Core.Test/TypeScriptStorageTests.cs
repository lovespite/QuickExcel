using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Domain;
using QuickExcel.Security;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TypeScriptStorageTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-typescript-" + Guid.NewGuid().ToString("N"));
    public TypeScriptStorageTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static ScriptModuleSet Modules() => new(new Dictionary<string, ScriptModule>
    {
        ["main"] = new("scripts/main.ts", "import {n} from 'helper'; const value:number=n; host.functions.register('F',()=>value);", ScriptLanguage.TypeScript),
        ["helper"] = new("scripts/helper.js", "export const n=7;")
    }, ["main"], scriptApiVersion: 2);

    [Fact]
    public void OriginalSourcesLanguagesAndV2SignatureSurviveSaveReopen()
    {
        var source = Modules();
        using var cert = ScriptSignatures.CreateSelfSigned("TypeScript 测试");
        var signed = source.WithSignature(ScriptSignatures.Sign(source, cert));
        Assert.Equal(2, signed.Signature!.Version);
        var file = Path.Combine(_folder, "original.xlsq");
        var storage = new TypedWorkbookStorage();
        storage.Save(file, Book(Sheet("S")).WithScriptModules(signed));
        var reopened = storage.Open(file).ScriptModules;
        Assert.Equal(signed.Modules.OrderBy(m => m.Key), reopened.Modules.OrderBy(m => m.Key));
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(reopened).Status);
        using var zip = ZipFile.OpenRead(file);
        using var reader = new StreamReader(zip.GetEntry("manifest.json")!.Open());
        var manifest = JsonNode.Parse(reader.ReadToEnd())!;
        Assert.Contains("typescript-sources", manifest["requiredFeatures"]!.AsArray().Select(n => n!.GetValue<string>()));
        Assert.Equal(ScriptModuleSet.TypeScriptOptions, manifest["scripts"]!["compilerOptions"]!.GetValue<string>());
        Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith(".map") || e.FullName.Contains("compiled"));
    }

    [Theory]
    [InlineData("feature")]
    [InlineData("language")]
    [InlineData("options")]
    [InlineData("v1")]
    public void InconsistentLanguageFeatureOrOptionsAreRejected(string mutation)
    {
        var file = Path.Combine(_folder, "bad.xlsq");
        var storage = new TypedWorkbookStorage();
        storage.Save(file, Book(Sheet("S")).WithScriptModules(Modules()));
        using (var zip = ZipFile.Open(file, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("manifest.json")!;
            JsonNode manifest;
            using (var reader = new StreamReader(entry.Open())) manifest = JsonNode.Parse(reader.ReadToEnd())!;
            switch (mutation)
            {
                case "feature":
                    var features = manifest["requiredFeatures"]!.AsArray();
                    features.Remove(features.Single(n => n!.GetValue<string>() == "typescript-sources")); break;
                case "language": manifest["scripts"]!["modules"]!["main"]!["language"] = "tsx"; break;
                case "options": manifest["scripts"]!["compilerOptions"] = "untrusted-options"; break;
                case "v1": manifest["scriptApiVersion"] = 1; break;
            }
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("manifest.json").Open(), new UTF8Encoding(false));
            writer.Write(manifest.ToJsonString());
        }
        Assert.Throws<InvalidDataException>(() => storage.Open(file));
    }

    [Fact]
    public void LanguageChangesInvalidateSignatureAuthorizationAndKeepExistingHistoryBoundary()
    {
        var modules = Modules();
        using var cert = ScriptSignatures.CreateSelfSigned("语言绑定");
        modules = modules.WithSignature(ScriptSignatures.Sign(modules, cert));
        var changed = new ScriptModuleSet(modules.Modules.SetItem("helper", modules.Modules["helper"] with { Language = ScriptLanguage.TypeScript }),
            modules.EntryModules, scriptApiVersion: 2, signature: modules.Signature);
        Assert.Equal(ScriptSignatureStatus.Invalid, ScriptSignatures.Verify(changed).Status);
        using var session = new DocumentSession(Book(Sheet("S")).WithScriptModules(modules));
        session.SetScriptsEnabled(true);
        session.Execute(new UpdateScriptModulesCommand(changed));
        Assert.False(session.ScriptsEnabled);
        Assert.Null(session.Snapshot.ScriptModules.Signature);
        Assert.False(session.Undo()); // Script edits belong to Monaco history, like existing JS edits.
        Assert.Equal(ScriptLanguage.TypeScript, session.Snapshot.ScriptModules.Modules["helper"].Language);
        Assert.False(session.ScriptsEnabled);
    }
}
