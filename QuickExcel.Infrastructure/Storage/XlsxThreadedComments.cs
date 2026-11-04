using System.Xml;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

/// <summary>Validated opaque local thread bundles. Never connects, looks up people or sends notifications.</summary>
internal static partial class XlsxThreadedComments
{
    internal static readonly XNamespace Thread = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
    private const string ThreadRelation = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    private const string PersonRelation = "http://schemas.microsoft.com/office/2017/10/relationships/person";
    private const string ThreadType = "application/vnd.ms-excel.threadedcomments+xml", PersonType = "application/vnd.ms-excel.person+xml";
    private static readonly XNamespace Revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    internal sealed record Bundle(byte[] Threads, byte[] Persons, byte[]? Fallback, CellPosition[] Positions)
    {
        internal System.Collections.Immutable.ImmutableArray<WorksheetComment> Models { get; init; } = [];
    }
    private static bool Attributes(XElement e, params XName[] names) => e.Attributes().All(a => a.IsNamespaceDeclaration || names.Contains(a.Name));
    private static Guid Identity(XAttribute? a) => Guid.TryParse(a?.Value, out var id) && id != Guid.Empty ? id : throw new InvalidDataException("线程评论标识无效。");

    internal static Bundle? Read(XlsxPackage package, string workbook, string sheet, string sheetName)
    {
        if (package.RelatedByUri(sheet, ThreadRelation) is not { } path) return null;
        package.RequireType(path, ThreadType);
        var persons = package.RelatedByUri(workbook, PersonRelation, true)!; package.RequireType(persons, PersonType);
        var people = package.XmlPart(persons).Root!; var comments = package.XmlPart(path).Root!;
        if (people.Name != Thread + "personList" || comments.Name != Thread + "ThreadedComments") throw new InvalidDataException("线程评论部件根元素无效。");
        Bundle? Unsupported(string reason)
        {
            package.Loss("线程评论只支持原位保留：" + reason, ExtendedObjectSafety.ThreadedComments, path, sheetName: sheetName); return null;
        }
        if (package.Relationships(path).Count != 0 || package.Relationships(persons).Count != 0) return Unsupported("包含未建模关系，未访问目标。");
        if (!Attributes(people) || !Attributes(comments) || people.Elements().Any(e => e.Name != Thread + "person") || comments.Elements().Any(e => e.Name != Thread + "threadedComment")) return Unsupported("包含未知集合扩展。");
        var personIds = new HashSet<Guid>();
        foreach (var person in people.Elements())
        {
            package.Budget.Parse(256);
            if (personIds.Count >= 4096 || !personIds.Add(Identity(person.Attribute("id")))) throw new InvalidDataException("线程评论作者重复或超过 4096。");
            if (person.HasElements || !Attributes(person, "id", "displayName", "userId", "providerId")) return Unsupported("作者含未知扩展。");
            if (person.Attribute("displayName") is null || person.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Value.Length > 1024)) throw new InvalidDataException("作者字段无效或过长。");
        }
        var entries = new Dictionary<Guid, (CellPosition Cell, Guid? Parent)>(); long textBytes = 0;
        foreach (var comment in comments.Elements())
        {
            package.Budget.Parse(512);
            var id = Identity(comment.Attribute("id")); var person = Identity(comment.Attribute("personId")); var position = TypedWorkbookXml.Position(Attr(comment, "ref"));
            var parent = comment.Attribute("parentId") is { } parentAttribute ? Identity(parentAttribute) : (Guid?)null;
            if (entries.Count >= 4096 || !entries.TryAdd(id, (position, parent)) || !personIds.Contains(person)) throw new InvalidDataException("评论重复、超限或作者不存在。");
            try { _ = XmlConvert.ToDateTimeOffset(Attr(comment, "dT")); } catch (FormatException e) { throw new InvalidDataException("线程评论时间无效。", e); }
            if (!Attributes(comment, "ref", "dT", "personId", "id", "parentId", "done") || comment.Elements().Count() != 1 || comment.Element(Thread + "text") is not { } text || text.HasElements || !Attributes(text, XNamespace.Xml + "space")) return Unsupported("提及或文本含尚未建模扩展。");
            if ((string?)comment.Attribute("done") is { } done && done is not ("0" or "1" or "true" or "false")) throw new InvalidDataException("线程评论状态无效。");
            if (text.Value.Length > WorkbookLimits.MaxCellTextLength || (textBytes += text.Value.Length * 2L) > 4 * 1024 * 1024) throw new InvalidDataException("线程评论文本超过限制。");
        }
        if (entries.Count == 0) throw new InvalidDataException("线程评论集合为空。");
        var roots = new Dictionary<CellPosition, Guid>();
        foreach (var (id, entry) in entries)
        {
            if (entry.Parent is { } parent)
            {
                if (!entries.TryGetValue(parent, out var owner) || owner.Parent is not null || owner.Cell != entry.Cell) throw new InvalidDataException("回复的根评论或单元格关联无效。");
            }
            else if (!roots.TryAdd(entry.Cell, id)) throw new InvalidDataException("同一单元格包含多个根评论。");
        }
        byte[]? fallback = null;
        if (package.RelatedByType(sheet, "comments") is { } legacy)
        {
            var root = package.XmlPart(legacy).Root!; var authors = root.Element(Main + "authors")?.Elements(Main + "author").Select(a => a.Value).ToArray() ?? [];
            var matches = new HashSet<CellPosition>();
            foreach (var comment in root.Element(Main + "commentList")?.Elements(Main + "comment") ?? [])
            {
                var position = TypedWorkbookXml.Position(Attr(comment, "ref")); if (!roots.TryGetValue(position, out var id)) continue;
                var author = TypedStyles.Int(Attr(comment, "authorId"), authors.Length - 1);
                if (!authors[author].StartsWith("tc=", StringComparison.Ordinal) || !Guid.TryParse(authors[author][3..], out var authorId) || authorId != id || Identity(comment.Attribute(Revision + "uid")) != id || !matches.Add(position)) throw new InvalidDataException("线程评论与旧版显示副本关联不一致。");
                if (comment.Attribute("shapeId") is { } shape && !uint.TryParse(shape.Value, out _)) throw new InvalidDataException("线程评论旧版形状标识无效。");
            }
            if (matches.Count != roots.Count) return Unsupported("缺少完整旧版显示副本。");
            package.ThreadedFallbackParts.Add(legacy); fallback = package.BinaryPart(legacy);
        }
        if (package.PreservedAttachments.Count + roots.Count > 4096) throw new InvalidDataException("工作簿线程评论锚点超过 4096。");
        package.PreservedAttachments.AddRange(roots.Keys.Select(p => new PreservedCellAttachment(ExtendedObjectSafety.ThreadedComments, sheetName, p)));
        package.Notice(new("xlsx.threaded-comments-preserved", CompatibilityKind.PreservedOnly, ExtendedObjectSafety.ThreadedComments,
            "线程评论、回复及作者关联按原位保留；只读显示旧版副本，不发送通知，不支持移动或编辑讨论。", path, sheetName));
        var authorsById = people.Elements().ToDictionary(p => Identity(p.Attribute("id")));
        var models = comments.Elements().Select(c =>
        {
            var author = authorsById[Identity(c.Attribute("personId"))];
            return new WorksheetComment(Identity(c.Attribute("id")), TypedWorkbookXml.Position(Attr(c, "ref")), c.Element(Thread + "text")!.Value,
                Attr(author, "displayName"), XmlConvert.ToDateTimeOffset(Attr(c, "dT")), Identity(c.Attribute("personId")),
                c.Attribute("parentId") is { } parent ? Identity(parent) : null, (string?)c.Attribute("done") is "1" or "true",
                (string?)author.Attribute("userId"), (string?)author.Attribute("providerId"));
        });
        return new(package.BinaryPart(path), package.BinaryPart(persons), fallback, roots.Keys.ToArray()) { Models = [.. models] };
    }

    internal static void Write(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, XElement workbookRelationships, XElement types, CancellationToken cancellationToken)
    {
        if (CommentsChanged(snapshot)) { WriteModeled(snapshot, parts, workbookRelationships, types, cancellationToken); return; }
        if (snapshot.FileState is not { } source || !source.PreservedAttachments.Any(a => a.Feature == ExtendedObjectSafety.ThreadedComments)) return;
        ExtendedObjectSafety.ValidateRetention(snapshot);
        var package = new XlsxPackage(source.OriginalWorkbook.ToArray(), new(cancellationToken)); var workbook = package.RelatedByType("", "officeDocument", true)!; var personsWritten = false;
        for (var i = 0; i < source.SheetParts.Length; i++)
        {
            var bundle = Read(package, workbook, source.SheetParts[i].Part, source.OriginalModel.Sheets[i].Name); if (bundle is null) continue;
            void Part(string path, string type, byte[] content)
            {
                parts[path] = content;
                if (!types.Elements(ContentTypes + "Override").Any(e => (string?)e.Attribute("PartName") == "/" + path)) types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", type)));
            }
            if (!personsWritten)
            {
                Part("xl/persons/person.xml", PersonType, bundle.Persons);
                workbookRelationships.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "threadPersons"), new XAttribute("Type", PersonRelation), new XAttribute("Target", "persons/person.xml"))); personsWritten = true;
            }
            var index = i + 1; Part($"xl/threadedComments/threadedComment{index}.xml", ThreadType, bundle.Threads);
            var relationPath = $"xl/worksheets/_rels/sheet{index}.xml.rels";
            var relations = parts.TryGetValue(relationPath, out var bytes) ? Xml(bytes, new(cancellationToken)) : new XDocument(new XElement(PackageRel + "Relationships"));
            relations.Root!.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "threadedComments"), new XAttribute("Type", ThreadRelation), new XAttribute("Target", $"../threadedComments/threadedComment{index}.xml"))); parts[relationPath] = XmlBytes(relations);
            if (bundle.Fallback is { } fallback)
            {
                var path = $"xl/comments{index}.xml"; if (!parts.ContainsKey(path)) throw new InvalidDataException("线程评论旧版副本未写入。"); parts[path] = fallback;
            }
        }
        if (package.PreservedAttachments.Count != source.PreservedAttachments.Count(a => a.Feature == ExtendedObjectSafety.ThreadedComments)) throw new InvalidDataException("线程评论保留清单不一致。");
    }
}
