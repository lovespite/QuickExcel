using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

internal sealed record PartRelationship(string Id, string Type, string Target, bool External);

internal sealed class XlsxPackage
{
    private readonly Dictionary<string, XDocument> _xml = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Dictionary<string, PartRelationship>> _relationships = new(StringComparer.Ordinal);
    private readonly HashSet<(string Source, string Id)> _usedRelationships = [];
    private readonly HashSet<string> _used = new(StringComparer.Ordinal) { "[Content_Types].xml" };
    private readonly Dictionary<string, string> _types = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _sheetNames = new(StringComparer.Ordinal);
    private (string Part, string Range)? _sourceRange;
    private (string Part, string? Range, string Feature)? _readContext;
    private Dictionary<(string Part, string Range), CompatibilityDiagnostic[]> _sharedDiagnostics = [];
    private readonly List<CompatibilityDiagnostic> _notices = [];
    public List<PreservedCellAttachment> PreservedAttachments { get; } = [];
    public HashSet<string> ThreadedFallbackParts { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, byte[]> Parts { get; }
    public bool IsStrict { get; private set; }
    public FileReadBudget Budget { get; }
    public CompatibilityIssueCollection Issues { get; } = new();

    public XlsxPackage(byte[] bytes, FileReadBudget budget)
    {
        Budget = budget;
        Parts = Unzip(bytes, budget, 2048);
        var contentTypes = XmlPart("[Content_Types].xml");
        if (contentTypes.Root?.Name != ContentTypes + "Types") throw new InvalidDataException("无效的 OPC 内容类型目录。");
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in contentTypes.Root.Elements())
        {
            if (item.Name == ContentTypes + "Default")
            {
                if (!defaults.TryAdd(Attr(item, "Extension"), Attr(item, "ContentType"))) throw new InvalidDataException("重复的默认内容类型。");
            }
            else if (item.Name == ContentTypes + "Override")
            {
                var name = Attr(item, "PartName");
                if (!name.StartsWith('/') || !_types.TryAdd(PathName(name[1..]), Attr(item, "ContentType")) || !Parts.ContainsKey(name[1..]))
                    throw new InvalidDataException("无效或重复的内容类型覆盖。");
            }
            else throw new InvalidDataException("未知的内容类型目录结构。");
        }
        foreach (var name in Parts.Keys.Where(n => n != "[Content_Types].xml"))
        {
            if (!_types.ContainsKey(name))
            {
                var ext = name[(name.LastIndexOf('.') + 1)..];
                if (!defaults.TryGetValue(ext, out var type)) throw new InvalidDataException($"部件缺少内容类型：{name}");
                _types.Add(name, type);
            }
            if (!name.EndsWith(".rels", StringComparison.Ordinal)) continue;
            var source = SourceForRelationships(name);
            if (source.Length > 0 && !Parts.ContainsKey(source)) throw new InvalidDataException($"关系源不存在：{source}");
            var root = XmlPart(name).Root;
            if (root?.Name != PackageRel + "Relationships") throw new InvalidDataException("无效的关系 XML。");
            var mapping = new Dictionary<string, PartRelationship>(StringComparer.Ordinal);
            foreach (var item in root.Elements())
            {
                if (item.Name != PackageRel + "Relationship") throw new InvalidDataException("关系目录包含未知元素。");
                var mode = (string?)item.Attribute("TargetMode");
                if (mode is not (null or "Internal" or "External")) throw new InvalidDataException("无效的关系目标模式。");
                var external = mode == "External";
                var target = external ? Attr(item, "Target") : Resolve(source, Attr(item, "Target"));
                if (!external && !Parts.ContainsKey(target)) throw new InvalidDataException($"关系目标不存在：{target}");
                var relationship = new PartRelationship(Attr(item, "Id"), Attr(item, "Type"), target, external);
                if (!mapping.TryAdd(relationship.Id, relationship)) throw new InvalidDataException("关系 ID 重复。");
            }
            if (!_relationships.TryAdd(source, mapping)) throw new InvalidDataException("关系来源重复。");
        }
        var active = new HashSet<string>();
        var visited = new HashSet<string>();
        void Visit(string source, int depth)
        {
            if (depth > 64 || active.Contains(source)) throw new InvalidDataException("部件关系包含循环或超过深度限制。");
            if (!visited.Add(source)) return;
            active.Add(source);
            foreach (var relation in Relationships(source).Values.Where(r => !r.External)) Visit(relation.Target, depth + 1);
            active.Remove(source);
        }
        foreach (var source in _relationships.Keys) Visit(source, 0);
    }
    public XDocument XmlPart(string path)
    {
        SetReadContext(path, null, "XML/part");
        _used.Add(path);
        if (!_xml.TryGetValue(path, out var result))
        {
            result = Xml(Required(Parts, path), Budget);
            if (path == "_rels/.rels") IsStrict = result.Root?.Elements(PackageRel + "Relationship").Any(r =>
                (string?)r.Attribute("Type") == "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument") == true;
            if (result.Root?.Name.LocalName == "workbook" && result.Root.Name.NamespaceName is "http://purl.oclc.org/ooxml/spreadsheetml/main" or "http://schemas.openxmlformats.org/spreadsheetml/2006/main" &&
                (result.Root.Name.NamespaceName == XlsxStrict.StrictMain) != IsStrict)
                throw new InvalidDataException("工作簿根命名空间与入口关系的 Strict/Transitional 类别不一致。");
            _ = XlsxStrict.Map(result, false, Budget.Check);
            _xml[path] = result;
        }
        return result;
    }
    public byte[] BinaryPart(string path) { _used.Add(path); return Required(Parts, path); }
    public IReadOnlyDictionary<string, PartRelationship> Relationships(string source) =>
        _relationships.GetValueOrDefault(source) ?? new Dictionary<string, PartRelationship>();
    public string Related(string source, string id, string type) => RelatedUri(source, id, Rel.NamespaceName + "/" + type);
    public string ExternalHyperlink(string source, string id)
    {
        if (!Relationships(source).TryGetValue(id, out var relation) || !relation.External || relation.Type != Rel.NamespaceName + "/hyperlink")
            throw new InvalidDataException("超链接关系无效。");
        _used.Add(RelationshipPart(source)); _usedRelationships.Add((source, id));
        return relation.Target; // Opaque address only; never access the target.
    }
    public string RelatedUri(string source, string id, string type)
    {
        _used.Add(RelationshipPart(source));
        if (!Relationships(source).TryGetValue(id, out var relation) || relation.External || relation.Type != type)
            throw new InvalidDataException($"{source}: 缺少有效的 {type} 关系 {id}。");
        _usedRelationships.Add((source, id));
        _used.Add(relation.Target);
        if (_sheetNames.TryGetValue(source, out var sheetName)) _sheetNames.TryAdd(relation.Target, sheetName);
        return relation.Target;
    }
    public string? RelatedByType(string source, string type, bool required = false) => RelatedByUri(source, Rel.NamespaceName + "/" + type, required);
    public string? RelatedByUri(string source, string type, bool required = false)
    {
        var matches = Relationships(source).Values.Where(r => r.Type == type).ToArray();
        if (matches.Length == 0 && !required) return null;
        if (matches.Length != 1) throw new InvalidDataException($"{source}: {type} 关系数量无效。");
        return RelatedUri(source, matches[0].Id, type);
    }
    public void RequireType(string path, string type)
    {
        if (!_types.TryGetValue(path, out var actual) || actual != type) throw new InvalidDataException($"{path}: 内容类型不匹配。");
    }
    public void FinishInventory()
    {
        foreach (var (source, relations) in _relationships)
        foreach (var relation in relations.Values)
            if (!_usedRelationships.Contains((source, relation.Id)))
                Loss($"{RelationshipPart(source)}: 未建模{(relation.External ? "外部" : "")}关系 {relation.Id}" + (relation.External ? "，未连接外部目标" : ""), "relationship/" + relation.Type.Split('/')[^1], RelationshipPart(source), range: relation.Id);
        foreach (var name in Parts.Keys.Where(n => !_used.Contains(n))) Loss($"未建模部件：{name}", "part/" + _types.GetValueOrDefault(name), name);
    }
    public void RegisterSheet(string path, string name) => _sheetNames[path] = name;
    public void SetReadContext(string part, string? range, string feature) => _readContext = (part, range, feature);
    public void AttachRejection(Exception exception)
    {
        if (_readContext is { } context)
            exception.Data[TypedWorkbookStorage.RejectionKey] = new CompatibilityDiagnostic("xlsx.import-rejected", CompatibilityKind.Rejected,
                context.Feature, exception.Message, context.Part, _sheetNames.GetValueOrDefault(context.Part), context.Range);
    }
    public IDisposable SourceRange(string part, string range)
    {
        var previous = _sourceRange; _sourceRange = (part, range); SetReadContext(part, range, "shared-content");
        return new RestoreScope(() => _sourceRange = previous);
    }
    private sealed class RestoreScope(Action restore) : IDisposable { public void Dispose() => restore(); }
    public void PrepareSharedLocations() => _sharedDiagnostics = Issues.Diagnostics.Where(d => d.Part is not null && d.Range is not null)
        .GroupBy(d => (d.Part!, d.Range!)).ToDictionary(g => g.Key, g => g.ToArray());
    public bool HasSharedLocations => _sharedDiagnostics.Count != 0;
    public void RegisterSharedUsage(string? sourcePart, string sourceRange, string sheetPart, string range)
    {
        if (sourcePart is null || !_sharedDiagnostics.TryGetValue((sourcePart, sourceRange), out var diagnostics)) return;
        foreach (var diagnostic in diagnostics)
            Issues.Add(diagnostic with { Code = "xlsx.affected-location", SheetName = _sheetNames.GetValueOrDefault(sheetPart),
                Range = range, Feature = diagnostic.Feature + " · " + sourceRange });
    }
    public void Notice(CompatibilityDiagnostic diagnostic) { if (_notices.Count < 128) _notices.Add(diagnostic); }
    public IReadOnlyList<CompatibilityDiagnostic> Diagnostics => Issues.Diagnostics.Concat(_notices).Select(d => d with {
        SheetName = d.SheetName ?? (d.Part is { } part ? _sheetNames.GetValueOrDefault(
            part.EndsWith(".rels", StringComparison.Ordinal) ? SourceForRelationships(part) : part) : null)
    }).ToArray();

    public void Loss(string message, string feature, string part, XElement? element = null, string? range = null, string? sheetName = null)
    {
        var cell = element?.AncestorsAndSelf(Main + "c").FirstOrDefault();
        range ??= (string?)cell?.Attribute("r") ?? (string?)element?.Attribute("ref");
        if (range is null && element?.Name == Main + "row") range = "row:" + (string?)element.Attribute("r");
        if (range is null && element?.Name == Main + "col") range = "columns:" + (string?)element.Attribute("min") + ":" + (string?)element.Attribute("max");
        if (range is null && _sourceRange is { } source && source.Part == part) range = source.Range;
        Issues.Add(new CompatibilityDiagnostic("xlsx.unmodeled", CompatibilityKind.PreservedOnly, feature, message, part, sheetName, range));
    }
    public void Shape(XElement element, string children, string attributes, string part)
    {
        var childSet = children.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
        var attributeSet = attributes.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
        foreach (var child in element.Elements())
            if (child.Name.Namespace != element.Name.Namespace || !childSet.Contains(child.Name.LocalName))
                Loss($"{part}: 未建模元素 {child.Name}", element.Name.LocalName + "/" + child.Name.LocalName, part, child);
        foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            if (attribute.Name.NamespaceName.Length != 0 || !attributeSet.Contains(attribute.Name.LocalName))
                Loss($"{part}: 未建模属性 {element.Name.LocalName}/@{attribute.Name}", element.Name.LocalName + "/@" + attribute.Name.LocalName, part, element);
    }
    internal static string Attr(XElement element, XName name) =>
        (string?)element.Attribute(name) is { Length: > 0 } text ? text : throw new InvalidDataException($"缺少属性 {name}。");
    private static string SourceForRelationships(string path)
    {
        if (path == "_rels/.rels") return "";
        var index = path.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (index < 0) throw new InvalidDataException("关系目录路径无效。");
        return path[..(index + 1)] + path[(index + 7)..^5];
    }
}


internal sealed class CompatibilityIssueCollection : System.Collections.ObjectModel.Collection<string>
{
    private readonly HashSet<string> _unique = new(StringComparer.Ordinal);
    private readonly HashSet<CompatibilityDiagnostic> _locations = [];
    private readonly List<CompatibilityDiagnostic> _diagnostics = [];
    public IReadOnlyList<CompatibilityDiagnostic> Diagnostics => _diagnostics;
    public void Add(CompatibilityDiagnostic diagnostic)
    {
        Add(diagnostic.Message); // Keep the existing loss guards and unique-message budget.
        if (_diagnostics.Count < 1024)
        {
            if (_locations.Add(diagnostic)) _diagnostics.Add(diagnostic);
        }
        else if (_diagnostics.Count == 1024)
            _diagnostics.Add(new("report.locations-truncated", CompatibilityKind.PreservedOnly, "报告容量",
                "仅列出前 1024 个兼容位置；其余定位已省略，保存保护仍覆盖全部未支持内容。"));
    }
    protected override void InsertItem(int index, string item)
    {
        if (!_unique.Add(item)) return;
        if (Count >= 1024) throw new InvalidDataException("未支持内容种类超过报告容量，无法可靠编辑此文件。");
        base.InsertItem(index, item);
    }
}
