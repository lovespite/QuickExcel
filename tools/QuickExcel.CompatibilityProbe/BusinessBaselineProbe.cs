using System.Globalization;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Formulas;

namespace QuickExcel.CompatibilityProbe;

/// <summary>Read-only import/display checks against independently captured, real Excel fixture values.</summary>
public static class BusinessBaselineProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory
    { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Business baseline must not execute scripts."); }
    public sealed record Check(string Location, string Property, string? Expected, string? Actual, bool Passed);

    public static string? CanonicalFormula(string? formula) => formula is null ? null : FormulaSyntax.Print(FormulaSyntax.Parse(formula));
    public static IReadOnlyList<Check> Compare(WorkbookSnapshot book, JsonElement expected, bool equivalentFormulas = false)
    {
        var checks = new List<Check>();
        void Equal(string location, string property, object? left, object? right)
        {
            var a = Convert.ToString(left, CultureInfo.InvariantCulture); var b = Convert.ToString(right, CultureInfo.InvariantCulture);
            checks.Add(new(location, property, a, b, a == b));
        }
        Equal("workbook", "Date1904", expected.GetProperty("Date1904").GetBoolean(), book.DateSystem == ExcelDateSystem.Mac1904);
        var sheets = expected.GetProperty("Sheets").EnumerateArray().ToArray();
        Equal("workbook", "SheetCount", sheets.Length, book.Sheets.Length);
        for (var index = 0; index < Math.Min(sheets.Length, book.Sheets.Length); index++)
        {
            var source = sheets[index]; var sheet = book.Sheets[index];
            Equal($"sheet[{index}]", "Name", source.GetProperty("Name").GetString(), sheet.Name);
            foreach (var item in source.GetProperty("Cells").EnumerateArray())
            {
                var row = item.GetProperty("Row").GetInt32() - 1; var col = item.GetProperty("Column").GetInt32() - 1;
                var location = $"{sheet.Name}!R{row + 1}C{col + 1}";
                var cell = sheet.GetCell(row, col); var value = cell?.Value ?? CellValue.Blank; var style = book.Styles[sheet.GetStyleId(row, col)];
                Equal(location, "Kind", item.GetProperty("Kind").GetString(), value.Kind);
                object? actual = value.Kind switch { CellValueKind.Blank => null, CellValueKind.Text => value.TextValue,
                    CellValueKind.Number => value.NumberValue, CellValueKind.Boolean => value.BooleanValue, _ => CellInput.DisplayText(value) };
                var raw = item.GetProperty("Value"); object? wanted = raw.ValueKind switch { JsonValueKind.Null => null,
                    JsonValueKind.Number => raw.GetDouble(), JsonValueKind.True => true, JsonValueKind.False => false, _ => raw.GetString() };
                Equal(location, "Value", wanted, actual);
                var expectedFormula = item.GetProperty("Formula").GetString(); var actualFormula = cell?.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText : null;
                Equal(location, "Formula", equivalentFormulas ? CanonicalFormula(expectedFormula) : expectedFormula,
                    equivalentFormulas ? CanonicalFormula(actualFormula) : actualFormula);
                Equal(location, "Text", item.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, style, book.DateSystem, CultureInfo.GetCultureInfo("zh-CN")));
                // Excel returns an empty aggregate font name for mixed Latin/East Asian text; per-character E1 evidence owns that case.
                if (item.GetProperty("Font").GetString() is { Length: > 0 } font)
                    Equal(location, "Font", font, value.Kind == CellValueKind.Blank && font == style.EastAsianFontName ? style.EastAsianFontName : style.FontName);
                Equal(location, "Size", item.GetProperty("Size").GetDouble(), style.FontSize);
                Equal(location, "Bold", item.GetProperty("Bold").GetBoolean(), style.Bold);
                Equal(location, "Italic", item.GetProperty("Italic").GetBoolean(), style.Italic);
            }
        }
        return checks;
    }

    public static int Run(string reference, string directory)
    {
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Evidence directory must be new.");
        Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(reference));
        var manifest = json.RootElement;
        if (manifest.GetProperty("Origin").GetString() != "real-excel-generated-fixtures" || manifest.GetProperty("Failures").GetArrayLength() != 0)
            throw new InvalidDataException("Expected a successful real Excel fixture capture.");
        var results = new List<object>(); var failures = 0; var totalChecks = 0;
        foreach (var fixture in manifest.GetProperty("Cases").EnumerateArray())
        {
            var id = fixture.GetProperty("Id").GetString()!; var file = fixture.GetProperty("File").GetString()!;
            if (Path.GetFileName(file) != file || !file.EndsWith(".xlsx", StringComparison.Ordinal)) throw new InvalidDataException("Invalid fixture filename.");
            var source = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(reference))!, file); var bytes = File.ReadAllBytes(source);
            if (!string.Equals(OpenXmlAudit.Hash(bytes), fixture.GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Fixture fingerprint mismatch.");
            var storage = new TypedWorkbookStorage(); var book = storage.Open(source); var expected = fixture.GetProperty("AfterReopen");
            var import = Compare(book, expected);
            var output = Path.Combine(directory, file); storage.Save(output, book);
            var reopen = Compare(storage.Open(output), expected);
            var xlsq = Path.ChangeExtension(output, ".xlsq"); storage.Save(xlsq, book);
            var xlsqReopen = Compare(storage.Open(xlsq), expected);
            using var session = new DocumentSession(book);
            var calculation = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
            if (!session.TryCommitCalculation(calculation)) throw new InvalidOperationException("Baseline calculation commit failed.");
            var recalculated = Compare(session.Snapshot, expected);
            var allFormulasFresh = calculation.Cells.Values.All(c => c.State == CalculationState.Fresh);
            var originalPreserved = bytes.SequenceEqual(File.ReadAllBytes(output)); var unchanged = bytes.SequenceEqual(File.ReadAllBytes(source));
            var audit = OpenXmlAudit.Inspect(bytes);
            var count = import.Concat(reopen).Concat(xlsqReopen).Concat(recalculated).Count(c => !c.Passed);
            failures += count + audit.Issues.Count + (originalPreserved && unchanged && allFormulasFresh ? 0 : 1);
            totalChecks += import.Count + reopen.Count + xlsqReopen.Count + recalculated.Count;
            results.Add(new { Id = id, Source = audit, SourceUnchanged = unchanged, OriginalPreserved = originalPreserved,
                Diagnostics = TypedWorkbookStorage.InspectCompatibility(book).Diagnostics, Import = import, Reopen = reopen, XlsqReopen = xlsqReopen,
                Recalculated = recalculated, AllFormulasFresh = allFormulasFresh });
            Console.WriteLine($"{id}: {count} comparison failures; original preserved={originalPreserved}; source schema issues={audit.Issues.Count}");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "business-baseline-report.json"), JsonSerializer.SerializeToUtf8Bytes(new
        { Excel = manifest.Clone(), Results = results, TotalChecks = totalChecks, Failures = failures, Scope = "import/display/recalculation/unchanged-XLSX/XLSQ; no edited-save claim" }, ProbeRunner.JsonOptions));
        return failures == 0 ? 0 : 2;
    }
}
