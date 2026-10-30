using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class MetadataProbe
{
    public static int Run(string sourcePath, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory);
        var storage = new TypedWorkbookStorage(); var bytes = File.ReadAllBytes(sourcePath); var original = storage.Open(sourcePath);
        using var session = new DocumentSession(original);
        session.Execute(new SetCellsCommand([new(new(original.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(7)))]));
        session.Execute(new RenameSheetCommand(original.Sheets[0].Id, "保存后"));
        var plan = TypedWorkbookStorage.InspectSavePlan(session.Snapshot);
        // Views and print settings remain separately gated until XC-0104/0105. Keep the exact loss list.
        var output = Path.Combine(directory, "edited.diagnostic-copy.xlsx");
        storage.SaveChecked(output, session.Snapshot, original.FileState!.CompatibilityIssues.Length > 0, new(null));
        var reopened = storage.Open(output);
        using var next = new DocumentSession(reopened);
        next.Execute(new SetCellsCommand([new(new(reopened.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        storage.SaveChecked(output, next.Snapshot, false, FileVersions.Read(output));
        var written = File.ReadAllBytes(output); var originalParts = OpenXmlAudit.ReadParts(bytes); var savedParts = OpenXmlAudit.ReadParts(written);
        var unchangedProperties = new[] { "docProps/core.xml", "docProps/custom.xml" }.All(p => originalParts.ContainsKey(p) && savedParts.ContainsKey(p) && originalParts[p].SequenceEqual(savedParts[p]));
        var appRetained = false;
        if (originalParts.TryGetValue("docProps/app.xml", out var originalApp) && savedParts.TryGetValue("docProps/app.xml", out var savedApp))
        {
            XElement Scalars(byte[] data)
            {
                var root = XDocument.Load(new MemoryStream(data)).Root!;
                root.Elements().Where(e => e.Name.LocalName is "HeadingPairs" or "TitlesOfParts").Remove(); return root;
            }
            appRetained = XNode.DeepEquals(Scalars(originalApp), Scalars(savedApp));
        }
        var audit = OpenXmlAudit.Inspect(written);
        var sourceUnchanged = bytes.SequenceEqual(File.ReadAllBytes(sourcePath));
        ProbeRunner.WriteNew(Path.Combine(directory, "metadata-report.json"), JsonSerializer.SerializeToUtf8Bytes(new
        {
            SourceSha256 = OpenXmlAudit.Inspect(bytes).Sha256, SourceUnchanged = sourceUnchanged, SavePlan = plan,
            UnchangedProperties = unchangedProperties, ExtendedPropertiesRetained = appRetained, RemainingLosses = original.FileState.CompatibilityIssues,
            Output = audit, Formula = storage.Open(output).Sheets[0].GetCell(0, 1)?.Content.FormulaText
        }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Properties retained={unchangedProperties}; source unchanged={sourceUnchanged}; schema issues={audit.Issues.Count}.");
        return unchangedProperties && appRetained && sourceUnchanged && audit.Issues.Count == 0 ? 0 : 2;
    }
}
