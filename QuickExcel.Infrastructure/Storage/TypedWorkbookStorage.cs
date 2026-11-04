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

/// <summary>Typed XLSX and XLSQ v2 storage, with read-only import and migration of XLSQ v1 packages.</summary>
public sealed partial class TypedWorkbookStorage : IVersionedWorkbookStorage, IWorkbookVirtualFiles
{
    private static readonly string[] Features = ["typed-workbook", "stable-sheet-ids", "script-modules"];
    private static bool HasExtendedErrors(WorkbookSnapshot snapshot) => snapshot.Sheets.Any(s => s.Cells.Values.Any(c => c.Value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.Calculation or CellErrorCode.Spill }));
    public WorkbookSnapshot Open(string path) => Open(path, CancellationToken.None);
    public WorkbookSnapshot Open(string path, CancellationToken cancellationToken)
    {
        try
        {
            RequireExtension(path, ".xlsx", ".xlsq");
            return OpenBytes(ReadFile(path, new(cancellationToken)), path, Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase), new(cancellationToken));
        }
        catch (Exception exception) when (exception is InvalidDataException or IOException or ArgumentException or NotSupportedException)
        {
            exception.Data[RejectionKey] = DescribeRejection(exception) ?? new CompatibilityDiagnostic("workbook.import-rejected", CompatibilityKind.Rejected,
                "file/open", exception.Message, Path.GetFileName(path));
            throw;
        }
    }

    private static WorkbookSnapshot OpenBytes(byte[] bytes, string sourcePath, bool isPackage, FileReadBudget budget)
    {
        byte[] workbookBytes = bytes;
        var version = 0;
        var modules = ScriptModuleSet.Single("");
        var packageFiles = VirtualFileSet.Empty;
        var declaredViews = false;
        var declaredTextVersion = false;
        var declaredExtendedErrors = false;
        var declaredDynamicArrays = false;
        var declaredPivots = false;
        var declaredStrict = false;
        var declaredThreads = false;
        JsonDocument? manifest = null, state = null;
        try
        {
            if (isPackage)
            {
                var parts = Unzip(bytes, budget, VirtualFileSet.MaxEntries + 72, xlsq: true);
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
                    declaredViews = features.Contains("workbook-views");
                    declaredTextVersion = features.Contains("text-compatibility-v2");
                    declaredExtendedErrors = features.Contains("extended-cell-errors");
                    declaredDynamicArrays = features.Contains("dynamic-arrays");
                    declaredPivots = features.Contains("worksheet-pivots");
                    declaredStrict = features.Contains("strict-ooxml");
                    declaredThreads = features.Contains("preserved-threaded-comments");
                    if (features.Distinct().Count() != features.Length || Features.Except(features).Any() || features.Except(Features.Concat(["strict-ooxml", "preserved-threaded-comments", "worksheet-pivots", "dynamic-arrays", "typescript-sources", "virtual-files", "script-actions", "script-signatures", "workbook-views", "worksheet-printing", "preserved-formulas", "cell-text-annotations", "worksheet-data", "worksheet-tables", "worksheet-validation", "worksheet-conditional-formatting", "worksheet-outline", "text-compatibility-v2", "extended-cell-errors", "drawing-anchors", "drawing-transforms", "worksheet-protection", "worksheet-charts", "worksheet-charts-v2", "chart-filters", "worksheet-default-width", "extended-capacity"])).Any() ||
                        features.Contains("script-actions") && scriptApiVersion != 2)
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
                        var language = module.Value.TryGetProperty("language", out var lang) ? ReadScriptLanguage(lang.GetString()) : ScriptLanguage.JavaScript;
                        if (features.Contains("typescript-sources") && !module.Value.TryGetProperty("language", out _))
                            throw new InvalidDataException("TS 包必须明确全部模块的语言。");
                        var description = module.Value.TryGetProperty("description", out _) ? String(module.Value, "description") : "";
                        moduleList.Add(new(module.Name, new ScriptModule(String(module.Value, "path"), Utf8.GetString(content), language, description)));
                    }
                    var bindings = Array(state.RootElement, "scriptBindings").ToArray();
                    if (bindings.Any(b => String(b, "scope") != "workbook")) throw new InvalidDataException("脚本绑定作用域无效。");
                    modules = new(moduleList, Array(scripts, "entryModules").Select(e => e.GetString() ?? throw new InvalidDataException("入口名无效。")),
                        Integer(state.RootElement, "scriptRevision"), Integer(state.RootElement, "bindingRevision"), scriptApiVersion,
                        bindings.Select(b => new ScriptBinding(String(b, "module"), String(b, "kind"))));
                    if (modules.HasActions != features.Contains("script-actions")) throw new InvalidDataException("动作绑定与必需能力不一致。");
                    if (modules.HasTypeScript != features.Contains("typescript-sources") ||
                        modules.HasTypeScript && String(scripts, "compilerOptions") != ScriptModuleSet.TypeScriptOptions ||
                        !modules.HasTypeScript && scripts.TryGetProperty("compilerOptions", out _))
                        throw new InvalidDataException("模块语言、编译选项与必需能力不一致。");
                    if (root.TryGetProperty("scriptSignature", out var signatureReference))
                    {
                        if (!features.Contains("script-signatures")) throw new InvalidDataException("脚本签名缺少必需能力声明。");
                        using var signatureJson = Json(Reference(signatureReference, "security/script-signature.json"), budget, 64 * 1024);
                        var signature = signatureJson.RootElement;
                        modules = modules.WithSignature(new(checked((int)Integer(signature, "version")), String(signature, "algorithm"),
                            String(signature, "certificate"), String(signature, "value")));
                    }
                    else if (features.Contains("script-signatures")) throw new InvalidDataException("缺少脚本签名。");
                    if (features.Contains("virtual-files"))
                    {
                        var files = new List<VirtualFileEntry>();
                        foreach (var item in Array(root, "files"))
                        {
                            var path = PathName(String(item, "path"));
                            var directory = Property(item, "directory").GetBoolean();
                            if (directory)
                            {
                                if (!referenced.Add(path + "/") || Required(parts, path + "/").Length != 0) throw new InvalidDataException("无效目录。");
                                files.Add(new("//xlsq/" + path, [], true));
                            }
                            else files.Add(new("//xlsq/" + path, Reference(item).ToImmutableArray()));
                        }
                        packageFiles = new(files);
                        _ = new VirtualFileSet(parts.Select(p => new VirtualFileEntry("//xlsq/" + p.Key.TrimEnd('/'), p.Value.ToImmutableArray(), p.Key.EndsWith('/'))));
                    }
                    if (referenced.Count != parts.Count) throw new InvalidDataException("XLSQ 包含未声明条目。");
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
            if (version == 2 && (model.TextCompatibilityVersion >= 2) != declaredTextVersion)
                throw new InvalidDataException("文本计算兼容版本与必需能力不一致。");
            if (version == 2 && HasExtendedErrors(model) != declaredExtendedErrors)
                throw new InvalidDataException("扩展错误值与必需能力声明不一致。");
            if (version == 2 && XlsxDynamicArrays.Present(model) != declaredDynamicArrays && read.PreservedFormulas.Count == 0)
                throw new InvalidDataException("动态数组与必需能力声明不一致。");
            if (version == 2 && XlsxPivots.Present(model) && !declaredPivots)
                throw new InvalidDataException("透视模型缺少必需能力声明。");
            if (version == 2 && read.IsStrict != declaredStrict) throw new InvalidDataException("Strict OOXML 与必需能力声明不一致。");
            if (version == 2 && read.PreservedAttachments.Any(a => a.Feature == ExtendedObjectSafety.ThreadedComments) && !declaredThreads)
                throw new InvalidDataException("线程评论保留缺少必需能力声明。");
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
                    var stableId = new SheetId(id);
                    var cells = sheet.Cells.Select(p => p.Value.SpillOwner is { } owner ? KeyValuePair.Create(p.Key,
                        new CellSnapshot(p.Value.Content, p.Value.Calculation, owner with { SheetId = stableId })) : p);
                    sheets[index] = new(stableId, sheet.Name, cells, sheet.Rows, sheet.Columns, sheet.MergedCells, sheet.Images, sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, sheet.Data);
                }
            }
            var activeSheet = sheets[model.Sheets.IndexOf(model.GetSheet(model.ActiveSheetId))].Id;
            var sourceActiveSheet = activeSheet;
            var sourceSheets = sheets.ToArray();
            var legacyViewApplied = false;
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
                if (!read.HasWorkbookView && !declaredViews && read.PreservedFormulas.Count == 0)
                {
                    legacyViewApplied = true;
                    // Earlier v2 documents used this optional state. Its zoom is a ratio, not a percentage.
                    if (view.TryGetProperty("activeSheetId", out _)) activeSheet = new(Identity(String(view, "activeSheetId")));
                    var index = System.Array.FindIndex(sheets, s => s.Id == activeSheet);
                    var legacyView = sheets[index].View;
                    if (view.TryGetProperty("selection", out _))
                    {
                        var range = CellRectangle.Parse(String(view, "selection"));
                        legacyView = legacyView with { TopLeft = new(new(range.Top, range.Left), range) };
                    }
                    if (view.TryGetProperty("zoom", out var ratio))
                    {
                        var percent = ratio.GetDouble() * 100;
                        if (percent != Math.Truncate(percent) || percent is < 10 or > 400) throw new InvalidDataException("旧视图缩放超出支持范围。");
                        legacyView = legacyView with { Zoom = (int)percent };
                    }
                    sheets[index] = sheets[index].WithView(legacyView);
                }
            }
            var names = model.DefinedNames.Select(n => n.Scope is { } scope
                ? n with { Scope = sheets[model.Sheets.IndexOf(model.GetSheet(scope))].Id } : n).ToArray();
            var snapshot = new WorkbookSnapshot(documentId, revision, model.DateSystem, sheets, modules.PrimarySource, model.Styles, modules, definedNames: names, activeSheetId: activeSheet, window: model.Window, textCompatibilityVersion: model.TextCompatibilityVersion);
            var originalModel = legacyViewApplied ? new WorkbookSnapshot(documentId, revision, model.DateSystem, sourceSheets, modules.PrimarySource,
                model.Styles, modules, definedNames: names, activeSheetId: sourceActiveSheet, window: model.Window, textCompatibilityVersion: model.TextCompatibilityVersion) : snapshot;
            var issues = read.Issues.ToList();
            // An original package can be retained byte-for-byte even for unsupported workbook features.
            var fileState = new WorkbookFileState(sourcePath, version, workbookBytes, originalModel, read.Bindings, issues,
                version == 2 ? manifest!.RootElement.GetRawText() : null, state?.RootElement.GetRawText(), read.Diagnostics, read.PreservedFormulas, read.PreservedAttachments);
            return new(documentId, revision, model.DateSystem, sheets, modules.PrimarySource, model.Styles, modules, fileState, names, activeSheet, model.Window, model.TextCompatibilityVersion, packageFiles);
        }
        catch (Exception exception) when (exception is JsonException or System.Xml.XmlException or FormatException or OverflowException or ArgumentException or InvalidOperationException && exception is not WorkbookCompatibilityException)
        { throw new InvalidDataException("工作簿结构或字段无效：" + exception.Message, exception); }
        finally { state?.Dispose(); manifest?.Dispose(); }
    }

    public void Save(string path, WorkbookSnapshot snapshot) => Save(path, snapshot, false);
    public void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported) => Save(path, snapshot, discardUnsupported, CancellationToken.None);
    public void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken) =>
        SaveCore(path, snapshot, discardUnsupported, cancellationToken, null);
    public FileVersion SaveChecked(string path, WorkbookSnapshot snapshot, bool discardUnsupported, FileVersion expected) =>
        SaveCore(path, snapshot, discardUnsupported, CancellationToken.None, expected);
    public FileVersion OverwriteChecked(string path, WorkbookSnapshot snapshot, FileVersion expected)
    {
        if (expected.Sha256 is null) throw new InvalidOperationException("覆盖原文件需要已打开文件的版本记录。");
        return SaveCore(path, snapshot, true, CancellationToken.None, expected, overwriteOriginal: true);
    }
    private FileVersion SaveCore(string path, WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken, FileVersion? expected, bool overwriteOriginal = false)
    {
        RequireExtension(path, ".xlsq", ".xlsx");
        ArgumentNullException.ThrowIfNull(snapshot);
        cancellationToken.ThrowIfCancellationRequested();
        snapshot = PrepareReadOnlyFormulaSave(snapshot, discardUnsupported);
        if (expected is not null) FileVersions.Require(path, expected);
        GuardOriginal(path, snapshot, discardUnsupported, overwriteOriginal);
        if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            return SaveXlsx(path, snapshot, discardUnsupported, cancellationToken, expected);
        return WritePackage(path, snapshot, discardUnsupported, cancellationToken, expected);
    }

    private static Dictionary<string, byte[]> BuildPackage(WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken)
    {
        if (snapshot.ScriptModules.Modules.Sum(m => (long)Utf8.GetByteCount(m.Value.Source)) > 1024 * 1024) throw new InvalidOperationException("Combined script source exceeds 1 MiB.");
        var unchanged = snapshot.FileState is { } state && ContentEquals(snapshot, state.OriginalModel);
        var issues = SaveIssues(snapshot);
        if (!discardUnsupported && issues.Count > 0) throw new WorkbookCompatibilityException(issues);
        var preserve = unchanged && !discardUnsupported;
        var workbook = preserve ? snapshot.FileState!.OriginalWorkbook.ToArray() : TypedWorkbookXml.Write(snapshot, false, cancellationToken);
        if (!preserve) XlsxMetadata.VerifyWritten(snapshot, workbook, cancellationToken);
        var parts = new Dictionary<string, byte[]> { ["workbook.xlsx"] = workbook };
        var sheetParts = preserve ? snapshot.FileState!.SheetParts.ToArray() : snapshot.Sheets.Select((_, i) => new WorkbookPartBinding((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), $"xl/worksheets/sheet{i + 1}.xml")).ToArray();
        JsonObject Object(string? json) => json is null ? new() : JsonNode.Parse(json)!.AsObject();
        var document = Object(snapshot.FileState?.StateJson);
        document["documentId"] = snapshot.DocumentId.ToString("D");
        document["contentRevision"] = snapshot.Revision;
        document["scriptRevision"] = snapshot.ScriptModules.Revision;
        document["bindingRevision"] = snapshot.ScriptModules.BindingRevision;
        document["sheets"] = MergeByKey(document["sheets"], JsonSerializer.SerializeToNode(snapshot.Sheets.Select((s, i) => new { id = s.Id.ToString(), workbookSheetId = sheetParts[i].WorkbookSheetId, part = sheetParts[i].Part }))!.AsArray(), "id");
        var savedBindings = new JsonArray();
        foreach (var binding in snapshot.ScriptModules.Bindings)
        {
            var prior = (document["scriptBindings"] as JsonArray)?.FirstOrDefault(b => b?["module"]?.GetValue<string>() == binding.Module && b?["kind"]?.GetValue<string>() == binding.Kind);
            savedBindings.Add(Merge(prior, JsonSerializer.SerializeToNode(new { module = binding.Module, scope = "workbook", kind = binding.Kind })!));
        }
        document["scriptBindings"] = savedBindings;
        document["calculation"] = Merge(document["calculation"], JsonSerializer.SerializeToNode(new { state = CacheState(snapshot), contentRevision = snapshot.Revision,
            scriptRevision = snapshot.ScriptModules.Revision, bindingRevision = snapshot.ScriptModules.BindingRevision,
            engineVersion = "QuickExcel-M3", scriptSetSha256 = ScriptHash(snapshot.ScriptModules) })!);
        var activeView = snapshot.GetSheet(snapshot.ActiveSheetId).View;
        if (HasWorkbookViews(snapshot))
            document["view"] = Merge(document["view"], JsonSerializer.SerializeToNode(new { activeSheetId = snapshot.ActiveSheetId.ToString(),
                selection = TypedWorkbookXml.Reference(activeView.ActiveCell), zoom = activeView.Zoom / 100d })!);
        else if (document["view"] is JsonObject optionalView)
        {
            optionalView.Remove("activeSheetId"); optionalView.Remove("selection"); optionalView.Remove("zoom");
            if (optionalView.Count == 0) document.Remove("view");
        }
        parts["state/document.json"] = JsonSerializer.SerializeToUtf8Bytes(document);
        var manifest = Object(snapshot.FileState?.ManifestJson);
        manifest["format"] = "QuickExcel"; manifest["formatVersion"] = 2; manifest["minReaderVersion"] = 2; manifest["scriptApiVersion"] = snapshot.ScriptModules.ScriptApiVersion;
        manifest["requiredFeatures"] = JsonSerializer.SerializeToNode(Features
            .Concat(XlsxDynamicArrays.Present(snapshot) ? ["dynamic-arrays"] : [])
            .Concat(XlsxPivots.Present(snapshot) ? ["worksheet-pivots"] : [])
            .Concat(XlsxStrict.SourceIsStrict(snapshot, cancellationToken) ? ["strict-ooxml"] : [])
            .Concat(snapshot.FileState?.PreservedAttachments.Any(a => a.Feature == ExtendedObjectSafety.ThreadedComments) == true ? ["preserved-threaded-comments"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Data.DefaultColumnWidth is not null) ? ["worksheet-default-width"] : [])
            .Concat(snapshot.Styles.Length > 4096 || snapshot.Sheets.Any(s => s.UsedRowCount > 100_000 || s.UsedColumnCount > 1024 || s.Rows.Keys.Any(r => r >= 100_000) || s.Columns.Keys.Any(c => c >= 1024)) ? ["extended-capacity"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Images.Any(i => i.Anchor.Rotation != 0) || s.Data.Charts.Any(c => c.Anchor.Rotation != 0)) ? ["drawing-transforms"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Data.Charts.Any(c => c.Series.Any(series => series.IsFiltered || !series.HiddenPoints.IsEmpty))) ? ["chart-filters"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Data.Charts.IsEmpty) ? ["worksheet-charts", "worksheet-charts-v2"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Images.Any(i => i.Anchor.Kind != DrawingAnchorKind.TwoCell || i.Anchor.RowOffset != 0 || i.Anchor.ColumnOffset != 0 || i.Anchor.EndRowOffset != 0 || i.Anchor.EndColumnOffset != 0)) ? ["drawing-anchors"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Data.Protection is not null) || snapshot.Styles.Any(s => !s.Locked || s.FormulaHidden) ? ["worksheet-protection"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Data.IsEmpty) ? ["worksheet-data"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Data.Tables.IsEmpty) ? ["worksheet-tables"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Data.Validations.IsEmpty) ? ["worksheet-validation"] : [])
            .Concat(snapshot.Sheets.Any(s => s.Data.Outline is not null || QuickExcel.Core.WorksheetOutline.HasGroups(s)) ? ["worksheet-outline"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Data.ConditionalFormats.IsEmpty) ? ["worksheet-conditional-formatting"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Annotations.IsEmpty || s.Cells.Values.Any(c => c.Content.RichText is not null)) ? ["cell-text-annotations"] : [])
            .Concat(snapshot.FileState?.FormulaReadOnly == true ? ["preserved-formulas"] : [])
            .Concat(snapshot.Sheets.Any(s => !s.Printing.IsEmpty) ? ["worksheet-printing"] : [])
            .Concat(HasWorkbookViews(snapshot) ? ["workbook-views"] : [])
            .Concat(snapshot.TextCompatibilityVersion >= 2 ? ["text-compatibility-v2"] : [])
            .Concat(HasExtendedErrors(snapshot) ? ["extended-cell-errors"] : [])
            .Concat(snapshot.ScriptModules.HasActions ? ["script-actions"] : [])
            .Concat(snapshot.ScriptModules.HasTypeScript ? ["typescript-sources"] : [])
            .Concat(snapshot.ScriptModules.Signature is not null ? ["script-signatures"] : []));
        JsonNode Reference(string name) => JsonSerializer.SerializeToNode(new { path = name, sha256 = Hash(parts[name]) })!;
        if (snapshot.ScriptModules.Signature is { } signature)
        {
            parts["security/script-signature.json"] = JsonSerializer.SerializeToUtf8Bytes(new {
                version = signature.Version, algorithm = signature.Algorithm, certificate = signature.Certificate, value = signature.Value });
            manifest["scriptSignature"] = Reference("security/script-signature.json");
        }
        else manifest.Remove("scriptSignature");
        manifest["workbook"] = Merge(manifest["workbook"], Reference("workbook.xlsx")); manifest["state"] = Merge(manifest["state"], Reference("state/document.json"));
        var modules = new JsonObject();
        foreach (var (id, module) in snapshot.ScriptModules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal))
        {
            parts.Add(module.Path, Utf8.GetBytes(module.Source));
            var item = Merge(manifest["scripts"]?["modules"]?[id], Reference(module.Path));
            if (module.Description.Length > 0) item["description"] = module.Description;
            else item.AsObject().Remove("description");
            if (snapshot.ScriptModules.HasTypeScript) item["language"] = module.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript";
            else item.AsObject().Remove("language");
            modules.Add(id, item);
        }
        manifest["scripts"] = Merge(manifest["scripts"], new JsonObject { ["entryModules"] = JsonSerializer.SerializeToNode(snapshot.ScriptModules.EntryModules), ["modules"] = modules });
        manifest["scripts"]!["modules"] = modules.DeepClone();
        if (snapshot.ScriptModules.HasTypeScript) manifest["scripts"]!["compilerOptions"] = ScriptModuleSet.TypeScriptOptions;
        else manifest["scripts"]!.AsObject().Remove("compilerOptions");
        if (snapshot.PackageFiles.Entries.Count > 0)
        {
            manifest["requiredFeatures"]!.AsArray().Add("virtual-files");
            var files = new JsonArray();
            foreach (var entry in snapshot.PackageFiles.Entries.Values.OrderBy(e => e.Path, StringComparer.Ordinal))
            {
                var name = PathName(entry.Path[7..]);
                var zipName = name + (entry.IsDirectory ? "/" : "");
                if (name == "manifest.json" || parts.ContainsKey(name) || !parts.TryAdd(zipName, entry.Data.ToArray())) throw new InvalidDataException("虚拟文件覆盖了系统条目。");
                var reference = entry.IsDirectory ? new JsonObject { ["path"] = name } : Reference(name).AsObject();
                reference["directory"] = entry.IsDirectory; files.Add(reference);
            }
            manifest["files"] = files;
        }
        else manifest.Remove("files");
        parts["manifest.json"] = JsonSerializer.SerializeToUtf8Bytes(manifest);
        return parts;
    }

    private static byte[] SerializePackage(WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken) => Zip(BuildPackage(snapshot, discardUnsupported, cancellationToken));

    private static FileVersion WritePackage(string path, WorkbookSnapshot snapshot, bool discardUnsupported, CancellationToken cancellationToken, FileVersion? expected)
    {
        var bytes = SerializePackage(snapshot, discardUnsupported, cancellationToken);
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var checkedSnapshot = OpenBytes(ReadFile(temp, new(cancellationToken)), path, true, new(cancellationToken));
            if (!ContentEquals(snapshot, checkedSnapshot, allowRowHeightRounding: true) || snapshot.DocumentId != checkedSnapshot.DocumentId || snapshot.Revision != checkedSnapshot.Revision ||
                !snapshot.Sheets.Select(s => s.Id).SequenceEqual(checkedSnapshot.Sheets.Select(s => s.Id)) || snapshot.ScriptModules.Signature != checkedSnapshot.ScriptModules.Signature || snapshot.ScriptModules.ScriptApiVersion != checkedSnapshot.ScriptModules.ScriptApiVersion || ScriptHash(snapshot.ScriptModules) != ScriptHash(checkedSnapshot.ScriptModules) || !snapshot.PackageFiles.ContentEquals(checkedSnapshot.PackageFiles))
                throw new InvalidDataException("保存后重开校验未能还原完整快照。");
        }, cancellationToken, expected);
        return new(Convert.ToHexString(SHA256.HashData(bytes)));
    }

    public void Export(string path, WorkbookSnapshot snapshot) => Export(path, snapshot, false, false);
    public void Export(string path, WorkbookSnapshot snapshot, bool allowStaleCaches, bool discardUnsupported)
    {
        if (snapshot.FileState?.PreservedFormulas.Any(f => !f.HasCompleteCache) == true)
            throw new InvalidOperationException("数组结果区域缺少可确认的完整缓存，不能导出值快照；请在 Excel 中重算后再打开。");
        RequireExtension(path, ".xlsx");
        GuardOriginal(path, snapshot, discardUnsupported);
        var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).ToArray();
        if (formulas.Any(c => c.Calculation?.Value is null)) throw new InvalidOperationException("部分公式缺少缓存结果，不能导出值快照。");
        if (!allowStaleCaches && formulas.Any(c => c.Calculation!.State is not (CalculationState.Fresh or CalculationState.Error) || c.Calculation.InputRevision != snapshot.Revision))
            throw new InvalidOperationException("部分公式结果可能过期；请重算或明确选择导出已有结果。");
        if (!allowStaleCaches && snapshot.Sheets.Any(s=>s.Data.Pivots.Any(p=>p.Stale)))
            throw new InvalidOperationException("透视结果尚未刷新；请刷新或明确选择导出已有结果。");
        if (!discardUnsupported && snapshot.FileState?.CompatibilityIssues.Length > 0) throw new WorkbookCompatibilityException(snapshot.FileState.CompatibilityIssues);
        var bytes = TypedWorkbookXml.Write(snapshot.WithDefinedNames([]), true, CancellationToken.None);
        PackageIO.AtomicWrite(path, bytes, temp =>
        {
            var read = TypedWorkbookXml.Read(ReadFile(temp, new(CancellationToken.None)), new(CancellationToken.None));
            var expectedSheets = snapshot.Sheets.Select(sheet => sheet.WithCells(sheet.Cells.ToImmutableDictionary(pair => pair.Key, pair =>
            {
                var value = pair.Value.Value;
                if (value.Kind == CellValueKind.Error && value.ErrorCode is CellErrorCode.Script or CellErrorCode.Cycle) value = CellValue.FromError(CellErrorCode.Value);
                return new CellSnapshot(pair.Value.Content.RichText is { } rich ? CellContent.FromRichText(rich, pair.Value.Content.StyleId) : CellContent.FromLiteral(value, pair.Value.Content.StyleId));
            })).WithData(sheet.Data with { Pivots=[] }));
            if (!ContentEquals(snapshot.WithDefinedNames([]).WithSheets(expectedSheets), read.Snapshot, allowRowHeightRounding: true)) throw new InvalidDataException("导出后类型/结构校验未通过。");
        }, CancellationToken.None);
    }

    public static IReadOnlyList<string> SaveIssues(WorkbookSnapshot snapshot)
    {
        snapshot = PrepareReadOnlyFormulaSave(snapshot, false);
        return snapshot.FileState is { } state && !ContentEquals(snapshot, state.OriginalModel) ? state.CompatibilityIssues : [];
    }
    private static void GuardOriginal(string path, WorkbookSnapshot snapshot, bool discardUnsupported, bool overwriteOriginal = false)
    {
        if (snapshot.FileState is not { } state) return;
        if ((state.RequiresMigration || discardUnsupported && !overwriteOriginal) && string.Equals(Path.GetFullPath(path), state.SourcePath, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("迁移或移除不支持内容时必须另存为新文件，原文件不会覆盖。");
    }

    internal static bool ContentEquals(WorkbookSnapshot left, WorkbookSnapshot right, bool allowRowHeightRounding = false)
    {
        // Only post-write verification tolerates point/DIP floating point roundoff.
        // Change detection and unsupported-content preservation still use exact equality.
        bool SameHeight(double? a, double? b) => a == b || allowRowHeightRounding && a.HasValue && b.HasValue &&
            Math.Abs(a.Value - b.Value) <= Math.Max(1, Math.Max(Math.Abs(a.Value), Math.Abs(b.Value))) * 1e-14;
        string NameKey(WorkbookSnapshot book, QuickExcel.Formulas.DefinedName name) =>
            (name.Scope is { } scope ? book.GetSheet(scope).Name : "") + "\0" + name.Name + "\0" + (allowRowHeightRounding ? QuickExcel.Formulas.StructuredReferences.CanonicalFormula(name.Formula) : name.Formula);
        bool SameContent(CellContent a, CellContent b) => a == b || allowRowHeightRounding && a.Kind == CellContentKind.Formula && b.Kind == CellContentKind.Formula &&
            a.StyleId == b.StyleId && a.DynamicArray == b.DynamicArray && QuickExcel.Formulas.StructuredReferences.CanonicalFormula(a.FormulaText) == QuickExcel.Formulas.StructuredReferences.CanonicalFormula(b.FormulaText);
        FormulaArrayDimensions? SpillSize(CellSnapshot cell) => cell.Calculation?.BlockedSpill ??
            (cell.Value is { Kind:CellValueKind.Error,ErrorCode:CellErrorCode.Spill } ? new FormulaArrayDimensions(1,1) : null);
        if (!left.DefinedNames.Select(n => NameKey(left, n)).Order().SequenceEqual(right.DefinedNames.Select(n => NameKey(right, n)).Order())) return false;
        if (left.GetSheet(left.ActiveSheetId).Name != right.GetSheet(right.ActiveSheetId).Name || left.Window != right.Window) return false;
        if (left.TextCompatibilityVersion != right.TextCompatibilityVersion || left.DateSystem != right.DateSystem || !left.Styles.SequenceEqual(right.Styles) || left.Sheets.Length != right.Sheets.Length) return false;
        for (var i = 0; i < left.Sheets.Length; i++)
        {
            var a = left.Sheets[i]; var b = right.Sheets[i];
            if (a.Name != b.Name || a.View != b.View || a.Visibility != b.Visibility || !a.Printing.ContentEquals(b.Printing) || !a.Annotations.ContentEquals(b.Annotations) || !a.Data.ContentEquals(b.Data, allowRowHeightRounding) || a.ShowGridLines != b.ShowGridLines || !SameHeight(a.DefaultRowHeight ?? 31, b.DefaultRowHeight ?? 31) || a.TabColor != b.TabColor ||
                !a.MergedCells.SequenceEqual(b.MergedCells) || a.Cells.Count != b.Cells.Count || a.Rows.Count != b.Rows.Count || a.Columns.Count != b.Columns.Count || a.Images.Length != b.Images.Length) return false;
            if (a.Rows.Any(p => !b.Rows.TryGetValue(p.Key, out var value) || value.Hidden != p.Value.Hidden || value.OutlineLevel != p.Value.OutlineLevel || value.Collapsed != p.Value.Collapsed || value.StyleId != p.Value.StyleId || value.ApplyStyle != p.Value.ApplyStyle ||
                    !SameHeight(value.Height, p.Value.Height)) || a.Columns.Any(p => !b.Columns.TryGetValue(p.Key, out var value) || value != p.Value)) return false;
            foreach (var (position, cell) in a.Cells)
                if (!b.Cells.TryGetValue(position, out var value) || !SameContent(cell.Content, value.Content) ||
                    ((cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null) && cell.Calculation?.Value != value.Calculation?.Value) ||
                    (cell.Calculation?.Array?.Rows ?? 1, cell.Calculation?.Array?.Columns ?? 1) != (value.Calculation?.Array?.Rows ?? 1, value.Calculation?.Array?.Columns ?? 1) ||
                    cell.Calculation?.Array is { } leftArray && value.Calculation?.Array is { } rightArray && !leftArray.ContentEquals(rightArray) ||
                    SpillSize(cell) != SpillSize(value) ||
                    (cell.SpillOwner?.Row, cell.SpillOwner?.Column) != (value.SpillOwner?.Row, value.SpillOwner?.Column)) return false;
            for (var image = 0; image < a.Images.Length; image++)
            {
                var x = a.Images[image]; var y = b.Images[image];
                if (x.Extension != y.Extension || x.Anchor != y.Anchor || x.PlaceInCell != y.PlaceInCell || !x.Data.SequenceEqual(y.Data)) return false;
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
        if (modules.HasTypeScript) return Hash(JsonSerializer.SerializeToUtf8Bytes(new
        {
            scriptApiVersion = 2, compilerOptions = ScriptModuleSet.TypeScriptOptions, entries = modules.EntryModules,
            modules = modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal).Select(m => new { id = m.Key, path = m.Value.Path,
                language = m.Value.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript", sha256 = Hash(Utf8.GetBytes(m.Value.Source)) }),
            bindings = modules.Bindings.OrderBy(b => b.Module, StringComparer.Ordinal).ThenBy(b => b.Kind, StringComparer.Ordinal)
        }));
        var sources = modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal).Select(m => new { id = m.Key, path = m.Value.Path, sha256 = Hash(Utf8.GetBytes(m.Value.Source)) });
        if (modules.HasActions) return Hash(JsonSerializer.SerializeToUtf8Bytes(new { scriptApiVersion = 2, entries = modules.EntryModules, modules = sources,
            bindings = modules.Bindings.OrderBy(b => b.Module, StringComparer.Ordinal).ThenBy(b => b.Kind, StringComparer.Ordinal).Select(b => new { module = b.Module, kind = b.Kind }) }));
        // Preserve existing API v1 digests; module-mode digests also bind the execution semantics.
        return modules.ScriptApiVersion == 1
            ? Hash(JsonSerializer.SerializeToUtf8Bytes(new { entries = modules.EntryModules, modules = sources }))
            : Hash(JsonSerializer.SerializeToUtf8Bytes(new { scriptApiVersion = 2, entries = modules.EntryModules, modules = sources }));
    }
    private static ScriptLanguage ReadScriptLanguage(string? value) => value switch
    {
        "javascript" => ScriptLanguage.JavaScript, "typescript" => ScriptLanguage.TypeScript,
        _ => throw new InvalidDataException("未知脚本语言。")
    };
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
