using System.Security.Cryptography;
using System.Collections.Immutable;
using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

public sealed class WorkbookCompatibilityException(IEnumerable<string> issues) : InvalidOperationException(
    "此操作会丢失尚未支持的内容，请取消或明确另存受支持副本：\n" + string.Join("\n", issues.Take(20)))
{
    public IReadOnlyList<string> Issues { get; } = issues.ToArray();
}

/// <summary>M2 typed XLSX and XLSQ v2 boundary. Legacy storage remains available for explicit v1 regression.</summary>
public sealed partial class TypedWorkbookStorage : IWorkbookStorageOptions
{
    private static readonly string[] Features = ["typed-workbook", "stable-sheet-ids", "script-modules"];
    public WorkbookSnapshot Open(string path) => Open(path, CancellationToken.None);
    public WorkbookSnapshot Open(string path, CancellationToken cancellationToken)
    {
        RequireExtension(path, ".xlsx", ".xlsq");
        return OpenBytes(ReadFile(path, new(cancellationToken)), path, Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase), new(cancellationToken));
    }

    private static WorkbookSnapshot OpenBytes(byte[] bytes, string sourcePath, bool isPackage, FileReadBudget budget)
    {
        byte[] workbookBytes = bytes;
        var version = 0;
        var modules = ScriptModuleSet.Single(WorkbookDocument.DefaultScript);
        JsonDocument? manifest = null, state = null;
        try
        {
            if (isPackage)
            {
                var parts = Unzip(bytes, budget, 72, xlsq: true);
                manifest = Json(Required(parts, "manifest.json", 64 * 1024), budget, 64 * 1024);
                var root = manifest.RootElement;
                version = checked((int)Integer(root, "formatVersion", 2));
                if (version is not (1 or 2)) throw new InvalidDataException("不支持此 XLSQ 主版本。");
                if (version == 1)
                {
                    if (root.TryGetProperty("format", out _) && String(root, "format") != "QuickExcel") throw new InvalidDataException("XLSQ 格式标识无效。");
                    if (root.TryGetProperty("scriptApiVersion", out _) && Integer(root, "scriptApiVersion", 1) != 1) throw new InvalidDataException("不支持的脚本 API。");
                    if (root.TryGetProperty("workbook", out _) && String(root, "workbook") != "workbook.xlsx" ||
                        root.TryGetProperty("script", out _) && String(root, "script") != "scripts/functions.js") throw new InvalidDataException("v1 入口路径无效。");
                    if (parts.Count != 3) throw new InvalidDataException("v1 包包含未声明条目。");
                    workbookBytes = Required(parts, "workbook.xlsx");
                    modules = ScriptModuleSet.Single(Utf8.GetString(Required(parts, "scripts/functions.js", 1024 * 1024)));
                }
                else
                {
                    var scriptApiVersion = checked((int)Integer(root, "scriptApiVersion", 2));
                    if (String(root, "format") != "QuickExcel" || Integer(root, "minReaderVersion", 2) != 2 || scriptApiVersion is not (1 or 2))
                        throw new InvalidDataException("XLSQ 格式、读取器版本或脚本 API 不受支持。");
                    var features = Array(root, "requiredFeatures").Select(e => e.GetString() ?? throw new InvalidDataException("能力名称必须为字符串。")).ToArray();
                    if (features.Length != Features.Length || features.Distinct().Count() != features.Length || features.Except(Features).Any())
                        throw new InvalidDataException("XLSQ 缺少必要能力或包含未知必需能力。");
                    var referenced = new HashSet<string>(StringComparer.Ordinal) { "manifest.json" };
                    byte[] Reference(JsonElement reference, string? expectedPath = null)
                    {
                        var path = PathName(String(reference, "path"));
                        if (expectedPath is not null && path != expectedPath || !referenced.Add(path)) throw new InvalidDataException("XLSQ 入口重复或路径无效。");
                        var content = Required(parts, path);
                        var hash = String(reference, "sha256");
                        if (hash.Length != 64 || hash.Any(c => c is not (>= '0' and <= '9' or >= 'a' and <= 'f')) || Hash(content) != hash)
                            throw new InvalidDataException($"{path}: SHA-256 不一致。");
                        return content;
                    }
                    workbookBytes = Reference(Property(root, "workbook"), "workbook.xlsx");
                    state = Json(Reference(Property(root, "state"), "state/document.json"), budget, 1024 * 1024);
                    var scripts = Property(root, "scripts");
                    var moduleList = new List<KeyValuePair<string, ScriptModule>>();
                    var map = Property(scripts, "modules");
                    if (map.ValueKind != JsonValueKind.Object) throw new InvalidDataException("模块目录必须为对象。");
                    long sourceBytes = 0;
                    foreach (var module in map.EnumerateObject())
                    {
                        if (moduleList.Count >= 64) throw new InvalidDataException("模块数量超过限制。");
                        var content = Reference(module.Value);
                        if ((sourceBytes += content.Length) > 1024 * 1024) throw new InvalidDataException("脚本总量超过 1 MiB。");
                        moduleList.Add(new(module.Name, new ScriptModule(String(module.Value, "path"), Utf8.GetString(content))));
                    }
                    modules = new(moduleList, Array(scripts, "entryModules").Select(e => e.GetString() ?? throw new InvalidDataException("入口名无效。")),
                        Integer(state.RootElement, "scriptRevision"), Integer(state.RootElement, "bindingRevision"), scriptApiVersion);
                    if (referenced.Count != parts.Count) throw new InvalidDataException("XLSQ 包含未声明条目。");
                    var bindings = Array(state.RootElement, "scriptBindings").ToArray();
                    var bound = new HashSet<string>();
                    foreach (var binding in bindings)
                        if (String(binding, "scope") != "workbook" || String(binding, "kind") != "formula-functions" ||
                            !modules.EntryModules.Contains(String(binding, "module")) || !bound.Add(String(binding, "module")))
                            throw new InvalidDataException("脚本绑定无效或需要尚未支持的能力。");
                    if (bound.Count != modules.EntryModules.Length) throw new InvalidDataException("脚本入口绑定不完整。");
                    var calculation = Property(state.RootElement, "calculation");
                    _ = String(calculation, "engineVersion");
                    if (String(calculation, "state") is not ("unknown" or "fresh" or "stale" or "partial")) throw new InvalidDataException("计算状态无效。");
                    if (String(calculation, "scriptSetSha256") != ScriptHash(modules)) throw new InvalidDataException("脚本集合摘要不一致。");
                    if (Integer(calculation, "contentRevision") > Integer(state.RootElement, "contentRevision") ||
                        Integer(calculation, "scriptRevision") > modules.Revision || Integer(calculation, "bindingRevision") > modules.BindingRevision)
                        throw new InvalidDataException("缓存修订来自未来状态。");
                }
            }
            var read = TypedWorkbookXml.Read(workbookBytes, budget);
            var model = read.Snapshot;
            var sheets = model.Sheets.ToArray();
            var documentId = Guid.NewGuid();
            long revision = 0;
            if (state is not null)
            {
                var root = state.RootElement;
                documentId = Identity(String(root, "documentId"));
                revision = Integer(root, "contentRevision");
                var bindings = Array(root, "sheets").ToArray();
                if (bindings.Length != sheets.Length) throw new InvalidDataException("状态中的工作表数量不匹配。");
                var seenIds = new HashSet<Guid>();
                var seenParts = new HashSet<string>();
                foreach (var binding in bindings)
                {
                    var id = Identity(String(binding, "id"));
                    var part = String(binding, "part");
                    var number = String(binding, "workbookSheetId");
                    var index = read.Bindings.ToList().FindIndex(b => b.Part == part && b.WorkbookSheetId == number);
                    if (index < 0 || !seenIds.Add(id) || !seenParts.Add(part)) throw new InvalidDataException("稳定表身份与 XLSX 关系绑定不一致。");
                    var sheet = sheets[index];
                    sheets[index] = new(new SheetId(id), sheet.Name, sheet.Cells, sheet.Rows, sheet.Columns, sheet.MergedCells, sheet.Images, sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor);
                }
            }
            if (state is not null && state.RootElement.TryGetProperty("view", out var view))
            {
                if (view.ValueKind != JsonValueKind.Object) throw new InvalidDataException("视图状态必须为对象。");
                if (view.TryGetProperty("activeSheetId", out _) && !sheets.Any(s => s.Id.Value == Identity(String(view, "activeSheetId"))))
                    throw new InvalidDataException("活动工作表身份不存在。");
                if (view.TryGetProperty("selection", out _))
                {
                    var selection = String(view, "selection").Split(':');
                    if (selection.Length is < 1 or > 2) throw new InvalidDataException("视图选区无效。");
                    foreach (var address in selection) _ = TypedWorkbookXml.Position(address);
                }
                if (view.TryGetProperty("zoom", out var zoom) && (zoom.ValueKind != JsonValueKind.Number || !zoom.TryGetDouble(out var number) || !double.IsFinite(number) || number <= 0))
                    throw new InvalidDataException("缩放值无效。");
            }
            var names = model.DefinedNames.Select(n => n.Scope is { } scope
                ? n with { Scope = sheets[model.Sheets.IndexOf(model.GetSheet(scope))].Id } : n).ToArray();
            var snapshot = new WorkbookSnapshot(documentId, revision, model.DateSystem, sheets, modules.PrimarySource, model.Styles, modules, definedNames: names);
            var issues = read.Issues.ToList();
            // An original package can be retained byte-for-byte even for unsupported workbook features.
            var fileState = new WorkbookFileState(sourcePath, version, workbookBytes, snapshot, read.Bindings, issues,
                version == 2 ? manifest!.RootElement.GetRawText() : null, state?.RootElement.GetRawText());
            return new(documentId, revision, model.DateSystem, sheets, modules.PrimarySource, model.Styles, modules, fileState, names);
        }
        catch (Exception exception) when (exception is JsonException or System.Xml.XmlException or FormatException or OverflowException or ArgumentException or InvalidOperationException && exception is not WorkbookCompatibilityException)
        { throw new InvalidDataException("工作簿结构或字段无效：" + exception.Message, exception); }
        finally { state?.Dispose(); manifest?.Dispose(); }
    }

    public void Save(string path, WorkbookSnapshot snapshot) => Save(path, snapshot, false);
    public void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported) => Save(path, snapshot, discardUnsupported, CancellationToken.None);
    public void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken)
    {
        RequireExtension(path, ".xlsq");
        ArgumentNullException.ThrowIfNull(snapshot);
        cancellationToken.ThrowIfCancellationRequested();
        GuardOriginal(path, snapshot, discardUnsupported);
        if (snapshot.ScriptModules.Modules.Sum(m => (long)Utf8.GetByteCount(m.Value.Source)) > 1024 * 1024) throw new InvalidOperationException("Combined script source exceeds 1 MiB.");
        var unchanged = snapshot.FileState is { } state && ContentEquals(snapshot, state.OriginalModel);
        var issues = SaveIssues(snapshot);
        if (!discardUnsupported && issues.Count > 0) throw new WorkbookCompatibilityException(issues);
        var preserve = unchanged && !discardUnsupported;
        var workbook = preserve ? snapshot.FileState!.OriginalWorkbook.ToArray() : TypedWorkbookXml.Write(snapshot, false, cancellationToken);
        var parts = new Dictionary<string, byte[]> { ["workbook.xlsx"] = workbook };
        var sheetParts = preserve ? snapshot.FileState!.SheetParts.ToArray() : snapshot.Sheets.Select((_, i) => new WorkbookPartBinding((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), $"xl/worksheets/sheet{i + 1}.xml")).ToArray();
        JsonObject Object(string? json) => json is null ? new() : JsonNode.Parse(json)!.AsObject();
        var document = Object(snapshot.FileState?.StateJson);
        document["documentId"] = snapshot.DocumentId.ToString("D");
        document["contentRevision"] = snapshot.Revision;
        document["scriptRevision"] = snapshot.ScriptModules.Revision;
        document["bindingRevision"] = snapshot.ScriptModules.BindingRevision;
        document["sheets"] = MergeByKey(document["sheets"], JsonSerializer.SerializeToNode(snapshot.Sheets.Select((s, i) => new { id = s.Id.ToString(), workbookSheetId = sheetParts[i].WorkbookSheetId, part = sheetParts[i].Part }))!.AsArray(), "id");
        document["scriptBindings"] = MergeByKey(document["scriptBindings"], JsonSerializer.SerializeToNode(snapshot.ScriptModules.EntryModules.Select(id => new { module = id, scope = "workbook", kind = "formula-functions" }))!.AsArray(), "module");
        document["calculation"] = Merge(document["calculation"], JsonSerializer.SerializeToNode(new { state = CacheState(snapshot), contentRevision = snapshot.Revision,
            scriptRevision = snapshot.ScriptModules.Revision, bindingRevision = snapshot.ScriptModules.BindingRevision,
            engineVersion = "QuickExcel-M3", scriptSetSha256 = ScriptHash(snapshot.ScriptModules) })!);
        if (document["view"] is JsonObject view && view["activeSheetId"] is { } active && !snapshot.Sheets.Any(s => s.Id.ToString() == active.GetValue<string>()))
        { view["activeSheetId"] = snapshot.Sheets[0].Id.ToString(); view["selection"] = "A1"; }
        parts["state/document.json"] = JsonSerializer.SerializeToUtf8Bytes(document);
        var manifest = Object(snapshot.FileState?.ManifestJson);
        manifest["format"] = "QuickExcel"; manifest["formatVersion"] = 2; manifest["minReaderVersion"] = 2; manifest["scriptApiVersion"] = snapshot.ScriptModules.ScriptApiVersion;
        manifest["requiredFeatures"] = JsonSerializer.SerializeToNode(Features);
        JsonNode Reference(string name) => JsonSerializer.SerializeToNode(new { path = name, sha256 = Hash(parts[name]) })!;
        manifest["workbook"] = Merge(manifest["workbook"], Reference("workbook.xlsx")); manifest["state"] = Merge(manifest["state"], Reference("state/document.json"));
        var modules = new JsonObject();
        foreach (var (id, module) in snapshot.ScriptModules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal))
        { parts.Add(module.Path, Utf8.GetBytes(module.Source)); modules.Add(id, Merge(manifest["scripts"]?["modules"]?[id], Reference(module.Path))); }
        manifest["scripts"] = Merge(manifest["scripts"], new JsonObject { ["entryModules"] = JsonSerializer.SerializeToNode(snapshot.ScriptModules.EntryModules), ["modules"] = modules });
        manifest["scripts"]!["modules"] = modules.DeepClone();
        parts["manifest.json"] = JsonSerializer.SerializeToUtf8Bytes(manifest);
        var bytes = Zip(parts);
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var checkedSnapshot = OpenBytes(ReadFile(temp, new(cancellationToken)), path, true, new(cancellationToken));
            if (!ContentEquals(snapshot, checkedSnapshot) || snapshot.DocumentId != checkedSnapshot.DocumentId || snapshot.Revision != checkedSnapshot.Revision ||
                !snapshot.Sheets.Select(s => s.Id).SequenceEqual(checkedSnapshot.Sheets.Select(s => s.Id)) || snapshot.ScriptModules.ScriptApiVersion != checkedSnapshot.ScriptModules.ScriptApiVersion || ScriptHash(snapshot.ScriptModules) != ScriptHash(checkedSnapshot.ScriptModules))
                throw new InvalidDataException("保存后重开校验未能还原完整快照。");
        }, cancellationToken);
    }

    public void Export(string path, WorkbookSnapshot snapshot) => Export(path, snapshot, false, false);
    public void Export(string path, WorkbookSnapshot snapshot, bool allowStaleCaches, bool discardUnsupported)
    {
        RequireExtension(path, ".xlsx");
        GuardOriginal(path, snapshot, discardUnsupported);
        var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).ToArray();
        if (formulas.Any(c => c.Calculation?.Value is null)) throw new InvalidOperationException("部分公式缺少缓存结果，不能导出值快照。");
        if (!allowStaleCaches && formulas.Any(c => c.Calculation!.State is not (CalculationState.Fresh or CalculationState.Error) || c.Calculation.InputRevision != snapshot.Revision))
            throw new InvalidOperationException("部分公式结果可能过期；请重算或明确选择导出已有结果。");
        if (!discardUnsupported && snapshot.FileState?.CompatibilityIssues.Length > 0) throw new WorkbookCompatibilityException(snapshot.FileState.CompatibilityIssues);
        var bytes = TypedWorkbookXml.Write(snapshot.WithDefinedNames([]), true, CancellationToken.None);
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var read = TypedWorkbookXml.Read(ReadFile(temp, new(CancellationToken.None)), new(CancellationToken.None));
            var expectedSheets = snapshot.Sheets.Select(sheet => sheet.WithCells(sheet.Cells.ToImmutableDictionary(pair => pair.Key, pair =>
            {
                var value = pair.Value.Value;
                if (value.Kind == CellValueKind.Error && value.ErrorCode is CellErrorCode.Script or CellErrorCode.Cycle) value = CellValue.FromError(CellErrorCode.Value);
                return new CellSnapshot(CellContent.FromLiteral(value, pair.Value.Content.StyleId));
            })));
            if (!ContentEquals(snapshot.WithDefinedNames([]).WithSheets(expectedSheets), read.Snapshot)) throw new InvalidDataException("导出后类型/结构校验未通过。");
        }, CancellationToken.None);
    }

    public static IReadOnlyList<string> SaveIssues(WorkbookSnapshot snapshot) => snapshot.FileState is { } state && !ContentEquals(snapshot, state.OriginalModel)
        ? state.CompatibilityIssues : [];
    private static void GuardOriginal(string path, WorkbookSnapshot snapshot, bool discardUnsupported)
    {
        if (snapshot.FileState is not { } state) return;
        if ((state.RequiresMigration || discardUnsupported) && string.Equals(Path.GetFullPath(path), state.SourcePath, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("迁移或移除不支持内容时必须另存为新文件，原文件不会覆盖。");
    }

    internal static bool ContentEquals(WorkbookSnapshot left, WorkbookSnapshot right)
    {
        string NameKey(WorkbookSnapshot book, QuickExcel.Formulas.DefinedName name) =>
            (name.Scope is { } scope ? book.GetSheet(scope).Name : "") + "\0" + name.Name + "\0" + name.Formula;
        if (!left.DefinedNames.Select(n => NameKey(left, n)).Order().SequenceEqual(right.DefinedNames.Select(n => NameKey(right, n)).Order())) return false;
        if (left.DateSystem != right.DateSystem || !left.Styles.SequenceEqual(right.Styles) || left.Sheets.Length != right.Sheets.Length) return false;
        for (var i = 0; i < left.Sheets.Length; i++)
        {
            var a = left.Sheets[i]; var b = right.Sheets[i];
            if (a.Name != b.Name || a.ShowGridLines != b.ShowGridLines || a.DefaultRowHeight != b.DefaultRowHeight || a.TabColor != b.TabColor ||
                !a.MergedCells.SequenceEqual(b.MergedCells) || a.Cells.Count != b.Cells.Count || a.Rows.Count != b.Rows.Count || a.Columns.Count != b.Columns.Count || a.Images.Length != b.Images.Length) return false;
            if (a.Rows.Any(p => !b.Rows.TryGetValue(p.Key, out var value) || value != p.Value) || a.Columns.Any(p => !b.Columns.TryGetValue(p.Key, out var value) || value != p.Value)) return false;
            foreach (var (position, cell) in a.Cells)
                if (!b.Cells.TryGetValue(position, out var value) || cell.Content != value.Content ||
                    (cell.Content.Kind == CellContentKind.Formula && cell.Calculation?.Value != value.Calculation?.Value)) return false;
            for (var image = 0; image < a.Images.Length; image++)
            {
                var x = a.Images[image]; var y = b.Images[image];
                if (x.Extension != y.Extension || x.RowIndex != y.RowIndex || x.ColumnIndex != y.ColumnIndex || x.RowSpan != y.RowSpan || x.ColumnSpan != y.ColumnSpan || x.PlaceInCell != y.PlaceInCell || !x.Data.SequenceEqual(y.Data)) return false;
            }
        }
        return true;
    }
    private static string CacheState(WorkbookSnapshot snapshot)
    {
        var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).ToArray();
        if (formulas.Length == 0) return "fresh";
        var fresh = formulas.Count(c => c.Calculation is { } result && result.InputRevision == snapshot.Revision
            && result.State is CalculationState.Fresh or CalculationState.Error);
        return fresh == formulas.Length ? "fresh" : fresh > 0 ? "partial"
            : formulas.Any(c => c.Calculation?.Value is not null) ? "stale" : "unknown";
    }

    private static JsonNode Merge(JsonNode? previous, JsonNode current)
    {
        if (previous is not JsonObject oldObject || current is not JsonObject newObject) return current.DeepClone();
        var result = oldObject.DeepClone().AsObject();
        foreach (var (key, value) in newObject) result[key] = value is null ? null : Merge(result[key], value);
        return result;
    }
    private static JsonArray MergeByKey(JsonNode? previous, JsonArray current, string key)
    {
        var result = new JsonArray();
        foreach (var item in current)
        {
            var old = (previous as JsonArray)?.FirstOrDefault(p => p?[key]?.ToJsonString() == item?[key]?.ToJsonString());
            result.Add(Merge(old, item!));
        }
        return result;
    }
    private static string Hash(byte[] content) => Convert.ToHexStringLower(SHA256.HashData(content));
    private static string ScriptHash(ScriptModuleSet modules)
    {
        var sources = modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal).Select(m => new { id = m.Key, path = m.Value.Path, sha256 = Hash(Utf8.GetBytes(m.Value.Source)) });
        // Preserve existing API v1 digests; module-mode digests also bind the execution semantics.
        return modules.ScriptApiVersion == 1
            ? Hash(JsonSerializer.SerializeToUtf8Bytes(new { entries = modules.EntryModules, modules = sources }))
            : Hash(JsonSerializer.SerializeToUtf8Bytes(new { scriptApiVersion = 2, entries = modules.EntryModules, modules = sources }));
    }
    private static JsonElement Property(JsonElement parent, string name) => parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(name, out var value)
        ? value : throw new InvalidDataException($"缺少字段 {name}。");
    private static string String(JsonElement parent, string name) => Property(parent, name).ValueKind == JsonValueKind.String
        ? Property(parent, name).GetString()! : throw new InvalidDataException($"{name} 必须是字符串。");
    private static IEnumerable<JsonElement> Array(JsonElement parent, string name) => Property(parent, name).ValueKind == JsonValueKind.Array
        ? Property(parent, name).EnumerateArray() : throw new InvalidDataException($"{name} 必须是数组。");
    private static long Integer(JsonElement parent, string name, long max = long.MaxValue - 1)
    {
        var value = Property(parent, name);
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt64(out var number) || number < 0 || number > max) throw new InvalidDataException($"{name} 必须是支持范围内的非负整数。");
        return number;
    }
    private static Guid Identity(string value) => Guid.TryParseExact(value, "D", out var id) && id != Guid.Empty ? id : throw new InvalidDataException("稳定身份无效。");
    private static void RequireExtension(string path, params string[] extensions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase)) throw new ArgumentException("文件扩展名不受支持。", nameof(path));
    }
}
