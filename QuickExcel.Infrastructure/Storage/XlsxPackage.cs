using System.Xml.Linq;
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
    public Dictionary<string, byte[]> Parts { get; }
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
                if (external) Issues.Add($"{name}: 外部关系 {relationship.Id}，未连接外部目标");
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
        _used.Add(path);
        if (!_xml.TryGetValue(path, out var result)) _xml[path] = result = Xml(Required(Parts, path), Budget);
        return result;
    }
    public byte[] BinaryPart(string path) { _used.Add(path); return Required(Parts, path); }
    public IReadOnlyDictionary<string, PartRelationship> Relationships(string source) =>
        _relationships.GetValueOrDefault(source) ?? new Dictionary<string, PartRelationship>();
    public string Related(string source, string id, string type)
    {
        _used.Add(RelationshipPart(source));
        if (!Relationships(source).TryGetValue(id, out var relation) || relation.External || relation.Type != (Rel.NamespaceName + "/" + type))
            throw new InvalidDataException($"{source}: 缺少有效的 {type} 关系 {id}。");
        _usedRelationships.Add((source, id));
        _used.Add(relation.Target);
        return relation.Target;
    }
    public string? RelatedByType(string source, string type, bool required = false)
    {
        var matches = Relationships(source).Values.Where(r => r.Type == (Rel.NamespaceName + "/" + type)).ToArray();
        if (matches.Length == 0 && !required) return null;
        if (matches.Length != 1) throw new InvalidDataException($"{source}: {type} 关系数量无效。");
        return Related(source, matches[0].Id, type);
    }
    public void RequireType(string path, string type)
    {
        if (!_types.TryGetValue(path, out var actual) || actual != type) throw new InvalidDataException($"{path}: 内容类型不匹配。");
    }
    public void FinishInventory()
    {
        foreach (var (source, relations) in _relationships)
        foreach (var relation in relations.Values)
            if (!_usedRelationships.Contains((source, relation.Id))) Issues.Add($"{RelationshipPart(source)}: 未建模关系 {relation.Id}");
        foreach (var name in Parts.Keys.Where(n => !_used.Contains(n))) Issues.Add($"未建模部件：{name}");
    }
    public void Shape(XElement element, string children, string attributes, string part)
    {
        var childSet = children.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
        var attributeSet = attributes.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
        foreach (var child in element.Elements())
            if (child.Name.Namespace != element.Name.Namespace || !childSet.Contains(child.Name.LocalName))
                Issues.Add($"{part}: 未建模元素 {child.Name}");
        foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            if (attribute.Name.NamespaceName.Length != 0 || !attributeSet.Contains(attribute.Name.LocalName))
                Issues.Add($"{part}: 未建模属性 {element.Name.LocalName}/@{attribute.Name}");
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
    protected override void InsertItem(int index, string item)
    {
        if (!_unique.Add(item)) return;
        if (Count >= 1024) throw new InvalidDataException("未支持内容种类超过报告容量，无法可靠编辑此文件。");
        base.InsertItem(index, item);
    }
}
