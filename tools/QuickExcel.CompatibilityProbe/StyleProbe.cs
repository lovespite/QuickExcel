using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class StyleProbe
{
    public static int Run(string sourcePath, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Evidence directory already exists.");
        Directory.CreateDirectory(directory);
        var bytes=File.ReadAllBytes(sourcePath); var storage=new TypedWorkbookStorage(); var book=storage.Open(sourcePath);
        using var session=new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id,0,1),CellContent.FromLiteral(CellValue.FromText("value edit")))]));
        session.Execute(new SetCellStyleCommand([new(book.Sheets[0].Id,2,0)],new(Italic:true)));
        var output=Path.Combine(directory,"edited.diagnostic-copy.xlsq"); storage.Save(output,session.Snapshot,true);
        var savedBytes=OpenXmlAudit.ReadParts(File.ReadAllBytes(output))["workbook.xlsx"];
        ProbeRunner.WriteNew(Path.ChangeExtension(output,"xlsx"),savedBytes);
        var source=OpenXmlAudit.Inspect(bytes);var saved=OpenXmlAudit.Inspect(savedBytes);
        var originalParts=OpenXmlAudit.ReadParts(bytes);var savedParts=OpenXmlAudit.ReadParts(savedBytes);
        var ns=OpenXmlAudit.Main;
        var original=XDocument.Load(new MemoryStream(originalParts["xl/styles.xml"]));
        var written=XDocument.Load(new MemoryStream(savedParts["xl/styles.xml"]));
        var retained=new[]{"cellStyleXfs","cellStyles"}.All(n=>XNode.DeepEquals(original.Root!.Element(ns+n),written.Root!.Element(ns+n))) &&
            original.Root!.Element(ns+"cellXfs")!.Elements().Zip(written.Root!.Element(ns+"cellXfs")!.Elements()).All(p=>XNode.DeepEquals(p.First,p.Second)) &&
            originalParts["xl/theme/theme1.xml"].SequenceEqual(savedParts["xl/theme/theme1.xml"]);
        ProbeRunner.WriteNew(Path.Combine(directory,"styles-report.json"),JsonSerializer.SerializeToUtf8Bytes(new{
            Source=source,Output=saved,OriginalStylesAndThemeRetained=retained,SourceUnchanged=bytes.SequenceEqual(File.ReadAllBytes(sourcePath)),
            InputStyles=book.Styles,OutputStyles=storage.Open(output).Styles,RemainingIssues=book.FileState!.CompatibilityIssues
        },ProbeRunner.JsonOptions));
        Console.WriteLine($"Original styles/theme retained={retained}; output schema issues={saved.Issues.Count}.");
        return retained && saved.Issues.Count==0?0:2;
    }
}
