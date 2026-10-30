using System.Text;
using System.Text.Json;
using QuickExcel.CompatibilityProbe;

try
{
    if (args is ["number-formats", var reference, var numberDirectory]) return NumberFormatProbe.Run(reference, numberDirectory);
    if (args is ["styles", var styleSource, var styleDirectory]) return StyleProbe.Run(styleSource, styleDirectory);
    if (args is ["metadata", var metadataSource, var metadataDirectory]) return MetadataProbe.Run(metadataSource, metadataDirectory);
    if (args is ["views", var viewSource, var viewDirectory]) return ViewProbe.Run(viewSource, viewDirectory);
    if (args is ["printing", var printSource, var printDirectory]) return PrintingProbe.Run(printSource, printDirectory);
    if (args is ["baseline", var directory])
    {
        directory = Path.GetFullPath(directory);
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Output directory must be new; previous evidence is never overwritten.");
        Directory.CreateDirectory(directory);
        var reports = FixtureCatalog.Create().Select(f => ProbeRunner.Run(f, Path.Combine(directory, f.Id))).ToArray();
        var summary = new StringBuilder("# XLSX compatibility baseline\n\nSynthetic fixtures; real Excel comparison pending. Differences include intended edits and representation changes; no automatic fidelity pass. XLSX outputs are extracted from saved XLSQ, not the desktop XLSX save flow.\n\n| Fixture | Input validation errors | Outputs | Output validation errors | Normal save blocked | Rejected paths |\n| --- | ---: | ---: | ---: | ---: | ---: |\n");
        foreach (var report in reports)
        {
            summary.AppendLine($"| {report.Id} | {report.Source.Issues.Count} | {report.Paths.Count(p => p.Output is not null)} | {report.Paths.Sum(p => p.Output?.Issues.Count ?? 0)} | {report.Paths.Count(p => p.Status.StartsWith("normal-save-blocked", StringComparison.Ordinal))} | {report.Paths.Count(p => p.Output is null)} |");
            Console.WriteLine($"{report.Id}: {report.Paths.Count} paths, source unchanged={report.SourceUnchanged}, Excel=pending");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "summary.md"), Encoding.UTF8.GetBytes(summary.ToString()));
        // A diagnostic run may expose existing output defects. Only broken input fixtures or source mutation fail baseline generation.
        return reports.All(r => r.SourceUnchanged && r.Source.Issues.Count == 0) ? 0 : 2;
    }
    if (args is ["compare", var before, var after, var output])
    {
        byte[] ReadBounded(string path) => new FileInfo(path).Length <= 64 * 1024 * 1024 ? File.ReadAllBytes(path) : throw new InvalidDataException("Input exceeds 64 MiB.");
        var left = OpenXmlAudit.Inspect(ReadBounded(before));
        var right = OpenXmlAudit.Inspect(ReadBounded(after));
        var differences = OpenXmlAudit.Compare(left, right);
        ProbeRunner.WriteNew(output, JsonSerializer.SerializeToUtf8Bytes(new { Before = left, After = right, Differences = differences, ExcelVerification = "not-performed" }, ProbeRunner.JsonOptions));
        Console.WriteLine($"{differences.Count} differences; {left.Issues.Count + right.Issues.Count} validation issues.");
        return differences.Count == 0 && left.Issues.Count == 0 && right.Issues.Count == 0 ? 0 : 2;
    }
    Console.Error.WriteLine("Usage: baseline <new-output-directory> | compare <before.xlsx> <after.xlsx> <new-report.json>");
    return 1;
}
catch (Exception ex) when (ex is IOException or InvalidOperationException or ArgumentException or System.Xml.XmlException)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}
