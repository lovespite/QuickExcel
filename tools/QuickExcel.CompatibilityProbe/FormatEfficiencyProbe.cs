using System.Globalization;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class FormatEfficiencyProbe
{
    public static int Run(string specPath, string oraclePath, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); using var spec = JsonDocument.Parse(File.ReadAllBytes(specPath)); using var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath));
        if (oracle.RootElement.GetProperty("SpecSha256").GetString() != OpenXmlAudit.Hash(File.ReadAllBytes(specPath)) || !oracle.RootElement.GetProperty("Stable").GetBoolean()) throw new InvalidDataException("Unbound or unstable Excel oracle.");
        var cases = spec.RootElement.GetProperty("Cases").EnumerateArray().ToArray(); var expected = oracle.RootElement.GetProperty("Cases").EnumerateArray().ToDictionary(c => c.GetProperty("Id").GetString()!);
        var styles = new List<StyleDefinition> { new() }; var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < cases.Length; row++)
        {
            var code = cases[row].GetProperty("Format").GetString(); var style = code == "General" ? new StyleDefinition() : new StyleDefinition { NumberFormatCode = code };
            var id = styles.IndexOf(style); if (id < 0) { id = styles.Count; styles.Add(style); }
            cells.Add(new(row, 0), new(CellContent.FromLiteral(CellValue.FromNumber(cases[row].GetProperty("Value").GetDouble()), id)));
        }
        var sheet = new WorksheetSnapshot(SheetId.New(), "Precision", cells, columns: [new(0, new(Width: 60))]);
        using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles));
        for (var row = 0; row < cases.Length; row++)
            foreach (var step in cases[row].GetProperty("Steps").EnumerateArray()) session.Execute(new AdjustDecimalPlacesCommand([new(sheet.Id, row, 0)], step.GetInt32()));
        var checks = new List<object>(); var failures = 0;
        void Check(WorkbookSnapshot snapshot, string stage)
        {
            for (var row = 0; row < cases.Length; row++)
            {
                var cell = snapshot.Sheets[0].GetCell(row, 0)!; var style = snapshot.Styles[cell.Content.StyleId];
                var text = CellDisplayFormatter.Format(cell.Value, style, snapshot.DateSystem, CultureInfo.GetCultureInfo("zh-CN"));
                var wanted = expected[cases[row].GetProperty("Id").GetString()!];
                var passed = cell.Value.NumberValue == wanted.GetProperty("Value").GetDouble() && text.Replace(" ", "") == wanted.GetProperty("Text").GetString()!.Replace(" ", "");
                if (!passed) failures++; checks.Add(new { Stage = stage, Id = cases[row].GetProperty("Id").GetString(), Passed = passed, Text = text, style.NumberFormatCode });
            }
        }
        Check(session.Snapshot, "adjusted");
        var storage = new TypedWorkbookStorage(); var paths = new List<object>();
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "precision." + extension); storage.Save(path, session.Snapshot); Check(storage.Open(path), extension);
            var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.ToArray() : [];
            if (issues.Length > 0) failures++; paths.Add(new { Extension = extension, Sha256 = OpenXmlAudit.Hash(File.ReadAllBytes(path)), Issues = issues });
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "precision-model.json"), JsonSerializer.SerializeToUtf8Bytes(new { Passed = failures == 0, Failures = failures,
            SpecSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(specPath)), OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), Paths = paths, Checks = checks }, ProbeRunner.JsonOptions));
        Console.WriteLine($"Precision checks={checks.Count}; failures={failures}"); return failures == 0 ? 0 : 2;
    }
}
