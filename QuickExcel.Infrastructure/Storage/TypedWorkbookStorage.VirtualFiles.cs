using QuickExcel.Serialization;
using System.Collections.Immutable;
using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    private static bool IsCompilationCachePath(string path, bool directory = false) =>
        !directory && path.Equals("cache", StringComparison.OrdinalIgnoreCase) ||
        path.Equals(QuickExcel.Runtime.ScriptCompilationCache.Part, StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith(QuickExcel.Runtime.ScriptCompilationCache.Part + "/", StringComparison.OrdinalIgnoreCase) ||
        path.Equals(QuickExcel.Runtime.ScriptCompilationCache.DataPart, StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith(QuickExcel.Runtime.ScriptCompilationCache.DataPart + "/", StringComparison.OrdinalIgnoreCase);

    public VirtualFileSet Capture(WorkbookSnapshot snapshot) => new(BuildPackage(snapshot.WithLocalStorage(ScriptStorage.Empty), false, CancellationToken.None, includeCompilationCache: false)
        .Select(p => new VirtualFileEntry("//xlsq/" + p.Key.TrimEnd('/'), p.Value.ToImmutableArray(), p.Key.EndsWith('/'))));

    public WorkbookSnapshot Apply(WorkbookSnapshot snapshot, IReadOnlyList<VirtualFileEdit> edits)
    {
        if (edits.Count == 0) return snapshot;
        if (edits.Any(e => !e.Path.StartsWith("//xlsq/", StringComparison.Ordinal))) throw new ArgumentException(MessageText.Get("包编辑只能访问 xlsq 卷。"));
        if (edits.Any(e => IsStoragePath(e.Path[7..]) || IsCompilationCachePath(e.Path[7..], directory: e.Operation != "write"))) throw new ArgumentException("脚本存储或编译缓存系统路径不能用作虚拟文件。");
        var original = Capture(snapshot);
        var updated = original.Apply(edits);
        if (original.ContentEquals(updated)) return snapshot;
        var systemPaths = original.Entries.Keys.Except(snapshot.PackageFiles.Entries.Keys).ToHashSet(StringComparer.Ordinal);
        if (edits.All(e => !systemPaths.Contains(e.Path)))
            return snapshot.WithPackageFiles(new(updated.Entries.Values.Where(e => !systemPaths.Contains(e.Path))));
        var parts = updated.Entries.Values.ToDictionary(e => e.Path[7..] + (e.IsDirectory ? "/" : ""), e => e.Data.ToArray(), StringComparer.Ordinal);
        var manifest = JsonNode.Parse(Required(parts, "manifest.json", 64 * 1024))!.AsObject();
        var state = JsonNode.Parse(Required(parts, "state/document.json", 1024 * 1024))!.AsObject();
        if (state["documentId"]?.GetValue<string>() != snapshot.DocumentId.ToString("D") || state["contentRevision"]?.GetValue<long>() != snapshot.Revision)
            throw new InvalidDataException(MessageText.Get("虚拟文件不能修改会话身份或内容修订。"));
        var system = new HashSet<string>(StringComparer.Ordinal) { "manifest.json" };
        void Refresh(JsonNode reference)
        {
            var name = PathName(reference["path"]!.GetValue<string>());
            if (!system.Add(name)) throw new InvalidDataException(MessageText.Get("重复系统条目。"));
            reference["sha256"] = Hash(Required(parts, name));
        }
        Refresh(manifest["workbook"]!);
        // Metadata and script sources are interpreted by the same strict importer used for opening files.
        var scriptMap = manifest["scripts"]!["modules"]!.AsObject();
        var modules = new ScriptModuleSet(scriptMap.Select(p => new KeyValuePair<string, ScriptModule>(p.Key,
            new(p.Value!["path"]!.GetValue<string>(), Utf8.GetString(Required(parts, p.Value["path"]!.GetValue<string>(), 1024 * 1024)),
                p.Value["language"] is { } language ? ReadScriptLanguage(language.GetValue<string>()) : ScriptLanguage.JavaScript,
                p.Value["description"] is { } description ? description.GetValue<string>() : ""))),
            manifest["scripts"]!["entryModules"]!.AsArray().Select(n => n!.GetValue<string>()),
            snapshot.ScriptModules.Revision, snapshot.ScriptModules.BindingRevision, manifest["scriptApiVersion"]!.GetValue<int>(),
            state["scriptBindings"]!.AsArray().Select(n => new ScriptBinding(n!["module"]!.GetValue<string>(), n["kind"]!.GetValue<string>())));
        var changedScripts = ScriptHash(modules) != ScriptHash(snapshot.ScriptModules) ||
            !modules.Modules.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Value.Description)
                .SequenceEqual(snapshot.ScriptModules.Modules.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Value.Description));
        if (changedScripts)
        {
            modules = new(modules.Modules, modules.EntryModules, checked(snapshot.ScriptModules.Revision + 1),
                checked(snapshot.ScriptModules.BindingRevision + 1), modules.ScriptApiVersion, modules.Bindings);
            manifest.Remove("scriptSignature"); parts.Remove("security/script-signature.json");
            var features = manifest["requiredFeatures"]!.AsArray();
            foreach (var node in features.Where(n => n?.GetValue<string>() == "script-signatures").ToArray()) features.Remove(node);
        }
        foreach (var item in scriptMap) Refresh(item.Value!);
        state["scriptRevision"] = modules.Revision; state["bindingRevision"] = modules.BindingRevision;
        state["calculation"]!["scriptSetSha256"] = ScriptHash(modules);
        state["calculation"]!["scriptRevision"] = modules.Revision; state["calculation"]!["bindingRevision"] = modules.BindingRevision;
        state["calculation"]!["state"] = "unknown";
        parts["state/document.json"] = JsonSerializer.SerializeToUtf8Bytes(state, InfrastructureJsonContext.Default.JsonObject);
        Refresh(manifest["state"]!);
        if (manifest["scriptSignature"] is { } signature) Refresh(signature);
        var files = new JsonArray();
        foreach (var pair in parts.Where(p => !system.Contains(p.Key)).OrderBy(p => p.Key, StringComparer.Ordinal))
            files.Add((JsonNode)new JsonObject { ["path"] = pair.Key.TrimEnd('/'), ["directory"] = pair.Key.EndsWith('/'), ["sha256"] = Hash(pair.Value) });
        var required = manifest["requiredFeatures"]!.AsArray();
        foreach (var node in required.Where(n => n?.GetValue<string>() == "virtual-files").ToArray()) required.Remove(node);
        if (files.Count > 0) { required.Add((JsonNode?)JsonValue.Create("virtual-files")); manifest["files"] = files; } else manifest.Remove("files");
        parts["manifest.json"] = JsonSerializer.SerializeToUtf8Bytes(manifest, InfrastructureJsonContext.Default.JsonObject);
        var imported = OpenBytes(Zip(parts), snapshot.FileState?.SourcePath ?? "virtual.xlsq", true, new(CancellationToken.None));
        if (snapshot.FileState?.RequiresMigration == true)
        {
            var stateInfo = imported.FileState!;
            imported = imported.WithFileState(new(snapshot.FileState.SourcePath, 1, stateInfo.OriginalWorkbook, stateInfo.OriginalModel,
                stateInfo.SheetParts, stateInfo.CompatibilityIssues, stateInfo.ManifestJson, stateInfo.StateJson,
                stateInfo.CompatibilityDiagnostics, stateInfo.PreservedFormulas));
        }
        // A raw replacement must not silently flatten rich text or discard preserved attachments/content.
        if (!original.Entries["//xlsq/workbook.xlsx"].Data.SequenceEqual(parts["workbook.xlsx"]))
        {
            if (snapshot.FileState?.CompatibilityIssues.Length > 0 || imported.FileState?.CompatibilityIssues.Length > 0 || imported.FileState?.FormulaReadOnly == true)
                throw new InvalidDataException(MessageText.Get("含未建模内容的 workbook.xlsx 不能通过虚拟文件替换。"));
            foreach (var sheet in snapshot.Sheets)
            {
                var next = imported.Sheets.FirstOrDefault(s => s.Id == sheet.Id);
                if (!sheet.Annotations.IsEmpty && (next is null || !sheet.Annotations.ContentEquals(next.Annotations)) ||
                    sheet.Cells.Any(c => c.Value.Content.RichText is not null && next?.GetCell(c.Key.Row, c.Key.Column)?.Content.RichText != c.Value.Content.RichText))
                    throw new InvalidDataException(MessageText.Get("原始包替换不能隐式移除富文本、链接或批注。"));
            }
        }
        return imported.WithLocalStorage(snapshot.LocalStorage).WithScriptModules(changedScripts ? modules : imported.ScriptModules.Signature == snapshot.ScriptModules.Signature
            ? snapshot.ScriptModules : imported.ScriptModules);
    }
}
