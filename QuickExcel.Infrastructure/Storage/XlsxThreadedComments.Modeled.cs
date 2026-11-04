using System.Xml;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static partial class XlsxThreadedComments
{
    private static bool CommentsChanged(WorkbookSnapshot snapshot) => snapshot.FileState is not { } state
        ? snapshot.Sheets.Any(s => !s.Data.ThreadComments.IsEmpty)
        : snapshot.Sheets.Any(s => !s.Data.ThreadComments.SequenceEqual(state.OriginalModel.Sheets.FirstOrDefault(o => o.Id == s.Id)?.Data.ThreadComments ?? []));

    private static void WriteModeled(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, XElement relationships, XElement types, CancellationToken token)
    {
        ExtendedObjectSafety.ValidateRetention(snapshot);
        if (snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Code == "xlsx.unmodeled" && d.Feature.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException("未建模线程评论不能重写。");
        var all = snapshot.Sheets.SelectMany(s => s.Data.ThreadComments).ToArray();
        if (all.Length == 0) return;
        if (all.Select(c => c.Id).Distinct().Count() != all.Length || all.GroupBy(c => c.PersonId).Any(g => g.Select(c => (c.Author, c.UserId, c.ProviderId)).Distinct().Count() != 1))
            throw new InvalidDataException("工作簿评论或作者身份冲突。");
        if (all.Length > 4096 || all.Sum(c => (long)c.Content.Length * 2) > 4 * 1024 * 1024) throw new InvalidDataException("工作簿线程评论容量超限。");
        void Part(string path, string type, XElement xml)
        {
            parts[path] = XmlBytes(new XDocument(xml));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", type)));
        }
        var people = new XElement(Thread + "personList", all.DistinctBy(c => c.PersonId).Select(c => new XElement(Thread + "person",
            new XAttribute("id", c.PersonId.ToString("B")), new XAttribute("displayName", c.Author),
            c.UserId is null ? null : new XAttribute("userId", c.UserId), c.ProviderId is null ? null : new XAttribute("providerId", c.ProviderId))));
        // Preserve unused author records and their account metadata from opaque fallback threads.
        var source = snapshot.FileState is { } state ? new XlsxPackage(state.OriginalWorkbook.ToArray(), new(token)) : null;
        var workbook = source?.RelatedByType("", "officeDocument", true);
        if (source is not null && workbook is not null && source.RelatedByUri(workbook, PersonRelation) is { } peoplePath)
            foreach (var person in source.XmlPart(peoplePath).Root!.Elements())
                if (!people.Elements().Any(p => Identity(p.Attribute("id")) == Identity(person.Attribute("id")))) people.Add(new XElement(person));
        Part("xl/persons/person.xml", PersonType, people);
        relationships.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "threadPersons"), new XAttribute("Type", PersonRelation), new XAttribute("Target", "persons/person.xml")));
        for (var i = 0; i < snapshot.Sheets.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var sheet = snapshot.Sheets[i]; if (sheet.Data.ThreadComments.IsEmpty) continue;
            var index = i + 1; var path = $"xl/threadedComments/threadedComment{index}.xml";
            var xml = new XElement(Thread + "ThreadedComments", sheet.Data.ThreadComments.Select(c => new XElement(Thread + "threadedComment",
                new XAttribute("ref", new CellRectangle(c.Position.Row, c.Position.Column, c.Position.Row, c.Position.Column).ToString().Split(':')[0]),
                new XAttribute("dT", XmlConvert.ToString(c.Created)), new XAttribute("personId", c.PersonId.ToString("B")), new XAttribute("id", c.Id.ToString("B")),
                c.ParentId is { } p ? new XAttribute("parentId", p.ToString("B")) : null, c.Resolved ? new XAttribute("done", "1") : null,
                new XElement(Thread + "text", new XAttribute(XNamespace.Xml + "space", "preserve"), c.Content))));
            Part(path, ThreadType, xml);
            if (sheet.Data.CommentFallback)
            {
                var sourceIndex = snapshot.FileState!.OriginalModel.Sheets.Select((s, index) => (s, index)).Single(p => p.s.Id == sheet.Id).index;
                var bundle = Read(source!, workbook!, snapshot.FileState.SheetParts[sourceIndex].Part, sheet.Name)!;
                if (!bundle.Models.SequenceEqual(sheet.Data.ThreadComments)) throw new NotSupportedException("旧版显示副本线程评论不可编辑。");
                parts[path] = bundle.Threads;
                if (bundle.Fallback is { } fallback) parts[$"xl/comments{index}.xml"] = fallback;
            }
            var relPath = $"xl/worksheets/_rels/sheet{index}.xml.rels";
            var rels = parts.TryGetValue(relPath, out var bytes) ? Xml(bytes, new(token)) : new XDocument(new XElement(PackageRel + "Relationships"));
            rels.Root!.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "threadedComments"), new XAttribute("Type", ThreadRelation), new XAttribute("Target", $"../threadedComments/threadedComment{index}.xml")));
            parts[relPath] = XmlBytes(rels);
        }
    }
}
