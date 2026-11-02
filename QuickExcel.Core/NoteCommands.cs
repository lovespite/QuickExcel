using System.Text;
using System.Xml;
using System.Xml.Linq;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class SetNoteCommand(SheetId sheetId, CellPosition position, CellNote? note, bool flattenRichText = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var old = sheet.Annotations;
        if (old.Notes.GetValueOrDefault(position) == note) return new(snapshot, [], false);
        if (old.Notes.GetValueOrDefault(position)?.Text.Runs.Any(r => r.Font is not null) == true && note is not null && !flattenRichText)
            throw new InvalidOperationException("此批注包含分段格式；请明确选择转换为普通文本后再替换。");
        var notes = note is null ? old.Notes.Remove(position) : old.Notes.SetItem(position, note);
        var drawing = old.NoteDrawing.ToArray();
        if (!old.Notes.Keys.ToHashSet().SetEquals(notes.Keys)) drawing = NoteDrawing.Reconcile(drawing, notes.Keys);
        var annotations = new WorksheetAnnotations(old.Links, notes, drawing);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithAnnotations(annotations) : s)), [], true, CalculationUnchanged: true);
    }
}

internal static class NoteDrawing
{
    internal static byte[] Reconcile(byte[] original, IEnumerable<CellPosition> positions)
    {
        var required = positions.ToHashSet(); if (required.Count == 0) return [];
        XNamespace v = "urn:schemas-microsoft-com:vml", x = "urn:schemas-microsoft-com:office:excel", o = "urn:schemas-microsoft-com:office:office";
        XElement root;
        if (original.Length == 0) root = new XElement("xml", new XAttribute(XNamespace.Xmlns + "v", v), new XAttribute(XNamespace.Xmlns + "x", x), new XAttribute(XNamespace.Xmlns + "o", o));
        else
        {
            using var reader = XmlReader.Create(new MemoryStream(original), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 });
            root = XElement.Load(reader, LoadOptions.PreserveWhitespace);
        }
        var existing = new HashSet<CellPosition>(); var ids = root.Descendants(v + "shape").Attributes("id").Select(a => a.Value).ToHashSet();
        foreach (var shape in root.Elements(v + "shape").ToArray())
        {
            var data = shape.Element(x + "ClientData");
            if ((string?)data?.Attribute("ObjectType") != "Note") throw new NotSupportedException("未建模 VML 不能编辑。");
            var position = new CellPosition((int)data!.Element(x + "Row")!, (int)data.Element(x + "Column")!);
            if (!required.Contains(position)) shape.Remove(); else existing.Add(position);
        }
        if (!root.Elements(v + "shapetype").Any(e => (string?)e.Attribute("id") == "_x0000_t202"))
            root.AddFirst(new XElement(v + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(o + "spt", 202), new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                new XElement(v + "stroke", new XAttribute("joinstyle", "miter")), new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect"))));
        var nextId = 1025;
        foreach (var p in required.Except(existing).OrderBy(p => p.Row).ThenBy(p => p.Column))
        {
            while (ids.Contains("_x0000_s" + nextId)) nextId++;
            root.Add(new XElement(v + "shape", new XAttribute("id", "_x0000_s" + nextId++), new XAttribute("type", "#_x0000_t202"),
                new XAttribute("style", "position:absolute;margin-left:0;margin-top:0;width:144pt;height:79pt;z-index:1;visibility:hidden"), new XAttribute("fillcolor", "#ffffe1"), new XAttribute(o + "insetmode", "auto"),
                new XElement(v + "fill", new XAttribute("color2", "#ffffe1")), new XElement(v + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
                new XElement(v + "path", new XAttribute(o + "connecttype", "none")), new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"))),
                new XElement(x + "ClientData", new XAttribute("ObjectType", "Note"), new XElement(x + "MoveWithCells"), new XElement(x + "SizeWithCells"),
                    new XElement(x + "Anchor", $"{p.Column}, 15, {p.Row}, 2, {Math.Min(PrintRange.MaxColumns - 1, p.Column + 2)}, 15, {Math.Min(PrintRange.MaxRows - 1, p.Row + 4)}, 2"),
                    new XElement(x + "AutoFill", "False"), new XElement(x + "Row", p.Row), new XElement(x + "Column", p.Column))));
        }
        return Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));
    }
}
