using System.Text.Json;

namespace QuickExcel.CompatibilityProbe;

public static class EditMatrixVerification
{
    public sealed record Difference(string Location, string? Expected, string? Actual, string Classification);
    /// <summary>All captured fields are exact except explicitly proven formula AST equivalence. Raw expressions stay in evidence.</summary>
    public static IReadOnlyList<Difference> Compare(JsonElement expected, JsonElement actual)
    {
        var differences = new List<Difference>();
        void Visit(JsonElement left, JsonElement right, string path)
        {
            if (left.ValueKind == JsonValueKind.Object && right.ValueKind == JsonValueKind.Object)
            {
                foreach (var key in left.EnumerateObject().Select(p => p.Name).Union(right.EnumerateObject().Select(p => p.Name)).Order(StringComparer.Ordinal))
                {
                    var a = left.TryGetProperty(key, out var l); var b = right.TryGetProperty(key, out var r);
                    if (a && b) Visit(l, r, path + "/" + key);
                    else differences.Add(new(path + "/" + key, a ? l.GetRawText() : null, b ? r.GetRawText() : null, "unexpected"));
                }
                return;
            }
            if (left.ValueKind == JsonValueKind.Array && right.ValueKind == JsonValueKind.Array)
            {
                var l = left.EnumerateArray().ToArray(); var r = right.EnumerateArray().ToArray();
                if (l.Length != r.Length) differences.Add(new(path + "/length", l.Length.ToString(), r.Length.ToString(), "unexpected"));
                for (var i = 0; i < Math.Min(l.Length, r.Length); i++) Visit(l[i], r[i], path + $"[{i}]");
                return;
            }
            if (left.ValueKind == right.ValueKind && (left.GetRawText() == right.GetRawText() ||
                left.ValueKind == JsonValueKind.Number && left.GetDouble() == right.GetDouble() ||
                left.ValueKind == JsonValueKind.String && left.GetString() == right.GetString())) return;
            var equivalent = false;
            if (path.EndsWith("/Formula", StringComparison.Ordinal) && left.ValueKind == JsonValueKind.String && right.ValueKind == JsonValueKind.String)
            {
                try { equivalent = BusinessBaselineProbe.CanonicalFormula(left.GetString()) == BusinessBaselineProbe.CanonicalFormula(right.GetString()); }
                catch (Exception ex) when (ex is FormatException or OverflowException) { /* Unrecognized expressions never receive an equivalence waiver. */ }
            }
            differences.Add(new(path, left.GetRawText(), right.GetRawText(), equivalent ? "equivalent-formula-serialization" : "unexpected"));
        }
        Visit(expected, actual, "workbook"); return differences;
    }

    public static int Run(string oraclePath, string actualPath, string modelPath, string reportPath)
    {
        using var oracleJson = JsonDocument.Parse(File.ReadAllBytes(oraclePath));
        using var actualJson = JsonDocument.Parse(File.ReadAllBytes(actualPath));
        using var modelJson = JsonDocument.Parse(File.ReadAllBytes(modelPath));
        var oracle = oracleJson.RootElement; var actual = actualJson.RootElement; var model = modelJson.RootElement;
        if (oracle.GetProperty("Origin").GetString() != "real-excel-edit-oracle" || actual.GetProperty("Origin").GetString() != "real-excel-output-verification")
            throw new InvalidDataException("Expected independent Excel oracle and output verification.");
        if (oracle.GetProperty("Failures").GetArrayLength() != 0 || actual.GetProperty("Failures").GetArrayLength() != 0)
            throw new InvalidDataException("Excel capture or source preservation failed.");
        if (Compare(oracle, model.GetProperty("Oracle")).Count != 0) throw new InvalidDataException("Model report was generated against a different Excel oracle.");
        var cases = oracle.GetProperty("Cases").EnumerateArray().ToArray(); var actualCases = actual.GetProperty("Cases").EnumerateArray().ToArray();
        if (cases.Length != 6 || actualCases.Length != cases.Length || model.GetProperty("Paths").GetArrayLength() != 48) throw new InvalidDataException("Incomplete edit matrix.");
        var reports = new List<object>(); var failures = model.GetProperty("Failures").GetInt32();
        foreach (var fixture in cases)
        {
            var id = fixture.GetProperty("Id").GetString(); var observed = actualCases.Single(c => c.GetProperty("Id").GetString() == id);
            if (!fixture.GetProperty("Paths").EnumerateArray().Select(p => p.GetProperty("Mode").GetString()).SequenceEqual(ProbeRunner.Modes) ||
                !observed.GetProperty("Paths").EnumerateArray().Select(p => p.GetProperty("Mode").GetString()).SequenceEqual(ProbeRunner.Modes))
                throw new InvalidDataException("Missing, extra or reordered edit paths.");
            foreach (var path in fixture.GetProperty("Paths").EnumerateArray())
            {
                var mode = path.GetProperty("Mode").GetString(); var output = observed.GetProperty("Paths").EnumerateArray().Single(p => p.GetProperty("Mode").GetString() == mode);
                var modelPathReport = model.GetProperty("Paths").EnumerateArray().Single(p => p.GetProperty("Id").GetString() == id && p.GetProperty("Mode").GetString() == mode);
                var differences = Compare(path.GetProperty("AfterReopen"), output.GetProperty("AfterReopen"));
                var hashMatches = string.Equals(output.GetProperty("SourceSha256").GetString(), modelPathReport.GetProperty("Output").GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase);
                var passed = hashMatches && output.GetProperty("SameAfterReopen").GetBoolean() && output.GetProperty("SourceUnchanged").GetBoolean() &&
                    modelPathReport.GetProperty("Passed").GetBoolean() && differences.All(d => d.Classification != "unexpected");
                if (!passed) failures++;
                reports.Add(new { Id = id, Mode = mode, Passed = passed, HashMatches = hashMatches, Differences = differences });
            }
        }
        ProbeRunner.WriteNew(reportPath, JsonSerializer.SerializeToUtf8Bytes(new
        { Passed = failures == 0, Failures = failures, OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)),
            ExcelVerificationSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(actualPath)), ModelReportSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(modelPath)), Paths = reports,
            Scope = "48 fixed business edit paths, all captured fields; package relationships/schema and source fingerprints in bound model report. Equivalence applies only to recognized formulas, not arbitrary unknown package XML." }, ProbeRunner.JsonOptions));
        Console.WriteLine($"48 edit paths; {failures} unexpected failures."); return failures == 0 ? 0 : 2;
    }
}
