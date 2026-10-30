using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class DataValidationProbe
{
    public static int Verify(string expectedPath, string actualPath, string verdictPath)
    {
        var expected = JsonDocument.Parse(File.ReadAllBytes(expectedPath)).RootElement;
        var actual = JsonDocument.Parse(File.ReadAllBytes(actualPath)).RootElement;
        if (!expected.GetProperty("Stable").GetBoolean() || !actual.GetProperty("Stable").GetBoolean()) throw new InvalidDataException("Unstable Excel evidence.");
        foreach (var item in new[] { expected, actual })
            if (OpenXmlAudit.Hash(File.ReadAllBytes(item.GetProperty("Path").GetString()!)) != item.GetProperty("Sha256").GetString()) throw new InvalidDataException("Unbound evidence.");
        var actualRows = actual.GetProperty("After").EnumerateArray().ToDictionary(e => e.GetProperty("Case").GetString()!);
        var differences = new List<object>(); var count = 0;
        if (actualRows.Count != expected.GetProperty("After").GetArrayLength()) throw new InvalidDataException("Case count changed.");
        foreach (var row in expected.GetProperty("After").EnumerateArray())
        {
            var key = row.GetProperty("Case").GetString()!; var next = actualRows[key];
            foreach (var property in row.EnumerateObject())
            {
                string Value(JsonElement record)
                {
                    var text = record.GetProperty(property.Name).ToString();
                    if (property.Name is "Formula1" or "Formula2" && text.StartsWith('='))
                        return FormulaSyntax.Print(FormulaSyntax.Parse(text));
                    return text;
                }
                var a = Value(row); var b = Value(next); count++;
                if (a != b) differences.Add(new { Case = key, Field = property.Name, Expected = a, Actual = b });
            }
        }
        var report = new { Passed = differences.Count == 0, Count = count, Differences = differences };
        File.WriteAllText(verdictPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Excel validation output: {count} checks; {differences.Count} differences."); return differences.Count == 0 ? 0 : 1;
    }

    public static int Run(string oraclePath, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory."); Directory.CreateDirectory(directory);
        var oracle = JsonDocument.Parse(File.ReadAllBytes(oraclePath)).RootElement;
        var source = oracle.GetProperty("Path").GetString()!;
        if (!oracle.GetProperty("Stable").GetBoolean() || OpenXmlAudit.Hash(File.ReadAllBytes(source)) != oracle.GetProperty("Sha256").GetString())
            throw new InvalidDataException("Unbound Excel oracle.");
        var storage = new TypedWorkbookStorage(); var validator = new WorkbookValidation(); var snapshot = storage.Open(source);
        var checks = new List<object>(); var differences = new List<object>(); var outputs = new List<object>();
        void Compare(string phase, string field, object? expected, object? actual)
        {
            var record = new { Phase = phase, Field = field, Expected = expected, Actual = actual, Passed = Equals(expected, actual) };
            checks.Add(record); if (!record.Passed) differences.Add(record);
        }
        void Check(WorkbookSnapshot book, string phase)
        {
            foreach (var row in oracle.GetProperty("After").EnumerateArray())
            {
                var sheet = book.Sheets.Single(s => s.Name == row.GetProperty("Sheet").GetString());
                var rect = CellRectangle.Parse(row.GetProperty("Address").GetString()!); var p = new CellPosition(rect.Top, rect.Left);
                var label = row.GetProperty("Case").GetString()!; var rules = sheet.Data.Validations.Where(v => v.Contains(p)).ToArray();
                Compare(phase, label + "/rule", 1, rules.Length); if (rules.Length != 1) continue;
                var rule = rules[0]; var failures = validator.Check(book, [new(sheet.Id, p.Row, p.Column)]);
                Compare(phase, label + "/valid", row.GetProperty("Valid").GetBoolean(), failures.Count == 0);
                Compare(phase, label + "/blank", row.GetProperty("AllowBlank").GetBoolean(), rule.AllowBlank);
                Compare(phase, label + "/dropDown", row.GetProperty("ShowDropDown").GetBoolean(), rule.ShowDropDown);
                Compare(phase, label + "/showError", row.GetProperty("ShowError").GetBoolean(), rule.ShowError);
                Compare(phase, label + "/showInput", row.GetProperty("ShowInput").GetBoolean(), rule.ShowInput);
                Compare(phase, label + "/alert", row.GetProperty("Alert").GetInt32() - 1, (int)rule.Alert);
                Compare(phase, label + "/inputTitle", row.GetProperty("InputTitle").GetString(), rule.InputTitle);
                Compare(phase, label + "/inputMessage", row.GetProperty("InputMessage").GetString(), rule.InputMessage);
                Compare(phase, label + "/errorTitle", row.GetProperty("ErrorTitle").GetString(), rule.ErrorTitle);
                Compare(phase, label + "/errorMessage", row.GetProperty("ErrorMessage").GetString(), rule.ErrorMessage);
            }
        }
        Check(snapshot, "import");
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "validation." + extension);
            storage.Save(path, snapshot, true); // Explicit modeled copy; original Excel UI extension metadata is reported separately.
            var read = storage.Open(path); Check(read, extension);
            var schema = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(path)).Issues.Select(i => i.ToString()).ToArray() : [];
            if (schema.Length > 0) differences.AddRange(schema);
            outputs.Add(new { Path = path, Sha256 = OpenXmlAudit.Hash(File.ReadAllBytes(path)), Schema = schema });
        }
        var report = new { Passed = differences.Count == 0, Count = checks.Count, SourceSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(source)), Differences = differences,
            Diagnostics = snapshot.FileState?.CompatibilityDiagnostics, Outputs = outputs };
        File.WriteAllText(Path.Combine(directory, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Data validation: {checks.Count} checks; {differences.Count} differences.");
        return differences.Count == 0 ? 0 : 1;
    }
}
