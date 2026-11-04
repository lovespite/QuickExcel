using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ScriptApiMigrationTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "QuickExcel-api-migration-" + Guid.NewGuid().ToString("N"));
    public ScriptApiMigrationTests() => Directory.CreateDirectory(_root);
    [Theory]
    [InlineData(1)] [InlineData(2)]
    public void OldSourcesRemainInspectableButCannotExecuteOrSaveUntilExplicitlyReplaced(int api)
    {
        var path = Path.Combine(_root, "legacy.xlsq"); var storage = new TypedWorkbookStorage();
        const string source = "throw Error('must never run');";
        if (api == 1) WorkbookPackageFixture.WriteV1(path, WorkbookPackageFixture.Xlsx("<row r=\"1\"><c r=\"A1\"><v>42</v></c></row>"), source);
        else
        {
            // Independent v2 envelope with hashes from its exact bytes.
            var workbook = WorkbookPackageFixture.Xlsx("<row r=\"1\"><c r=\"A1\"><v>42</v></c></row>");
            var script = Encoding.UTF8.GetBytes(source);
            var scriptSet = System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(new { scriptApiVersion = 2, entries = new[] { "main" },
                modules = new[] { new { id = "main", path = "scripts/main.js", sha256 = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(script)) } } });
            var state = Encoding.UTF8.GetBytes(new JsonObject {
                ["documentId"] = Guid.NewGuid().ToString(), ["contentRevision"] = 0, ["scriptRevision"] = 0, ["bindingRevision"] = 0,
                ["sheets"] = new JsonArray(new JsonObject { ["id"] = Guid.NewGuid().ToString(), ["workbookSheetId"] = "1", ["part"] = "xl/worksheets/sheet1.xml" }),
                ["scriptBindings"] = new JsonArray(new JsonObject { ["scope"] = "workbook", ["module"] = "main", ["kind"] = "formula-functions" }),
                ["calculation"] = new JsonObject { ["engineVersion"] = "legacy", ["state"] = "unknown", ["contentRevision"] = 0,
                    ["scriptRevision"] = 0, ["bindingRevision"] = 0, ["scriptSetSha256"] = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(scriptSet)) }
            }.ToJsonString());
            JsonObject Ref(string name, byte[] bytes) => new() { ["path"] = name, ["sha256"] = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant() };
            var manifest = new JsonObject {
                ["format"] = "QuickExcel", ["formatVersion"] = 2, ["minReaderVersion"] = 2, ["scriptApiVersion"] = 2,
                ["requiredFeatures"] = new JsonArray("typed-workbook", "stable-sheet-ids", "script-modules"),
                ["workbook"] = Ref("workbook.xlsx", workbook), ["state"] = Ref("state/document.json", state),
                ["scripts"] = new JsonObject { ["entryModules"] = new JsonArray("main"), ["modules"] = new JsonObject { ["main"] = Ref("scripts/main.js", script) } }
            };
            using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
            foreach (var (name, bytes) in new[] { ("workbook.xlsx", workbook), ("scripts/main.js", script), ("state/document.json", state), ("manifest.json", Encoding.UTF8.GetBytes(manifest.ToJsonString())) })
            { using var output = zip.CreateEntry(name).Open(); output.Write(bytes); }
        }
        var original = File.ReadAllBytes(path); var loaded = storage.Open(path);
        Assert.Equal(api, loaded.ScriptModules.ScriptApiVersion); Assert.Equal(source, loaded.ScriptSource);
        using var session = new DocumentSession(loaded);
        Assert.Throws<NotSupportedException>(() => session.SetScriptsEnabled(true));
        var target = Path.Combine(_root, "upgraded.xlsq");
        Assert.Throws<NotSupportedException>(() => storage.Save(target, loaded)); Assert.False(File.Exists(target));
        var modules = loaded.ScriptModules;
        var rewritten = new ScriptModuleSet(modules.Modules.SetItem("main", new("scripts/main.js", "host.functions.register('ANSWER',()=>42);")), modules.EntryModules, scriptApiVersion: 3);
        storage.Save(target, loaded.WithScriptModules(rewritten));
        Assert.Equal(3, storage.Open(target).ScriptModules.ScriptApiVersion); Assert.Equal(original, File.ReadAllBytes(path));
    }
    public void Dispose() => Directory.Delete(_root, true);
}
