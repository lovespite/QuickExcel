using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

/// <summary>Checks visual rules after a normal Excel save/reopen, independently of the UI smoke.</summary>
public static class ConditionalVisualProbe
{
    public static int Compare(string sourcePath, string excelPath, string reportPath)
    {
        if (File.Exists(reportPath)) throw new IOException("Use a new report path.");
        var storage = new TypedWorkbookStorage(); var source = storage.Open(sourcePath); var excel = storage.Open(excelPath);
        var checks = new List<object>(); var passed = true;
        foreach (var sheet in source.Sheets)
        {
            var expected = sheet.Data.ConditionalFormats.OrderBy(r => r.Priority).ToArray();
            var actual = excel.Sheets.Single(s => s.Name == sheet.Name).Data.ConditionalFormats.OrderBy(r => r.Priority).ToArray();
            var countMatches = expected.Length == actual.Length; passed &= countMatches;
            checks.Add(new { Sheet = sheet.Name, Check = "rule count", Passed = countMatches, Expected = expected.Length, Actual = actual.Length });
            foreach (var pair in expected.Zip(actual))
            {
                var same = pair.First.ContentEquals(pair.Second); passed &= same;
                checks.Add(new { Sheet = sheet.Name, Check = pair.First.Kind.ToString(), Priority = pair.First.Priority, Passed = same, Expected = pair.First, Actual = pair.Second });
            }
        }
        var diagnostics = excel.FileState?.CompatibilityDiagnostics.Where(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)).ToArray() ?? [];
        passed &= diagnostics.Length == 0;
        var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(sourcePath)); passed &= audit.Issues.Count == 0;
        File.WriteAllText(reportPath, JsonSerializer.Serialize(new { Passed = passed, Checks = checks, ConditionalDiagnostics = diagnostics, SourceAudit = audit.Issues }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Visual roundtrip: {checks.Count} checks; passed={passed}."); return passed ? 0 : 1;
    }
}
