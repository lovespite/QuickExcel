using System.Collections.Immutable;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DataCleaningFixtureTests
{
    [Theory]
    [InlineData("1900")]
    [InlineData("1904")]
    public void NativeCommandsAndDisplayedTypesSurviveUndoAndBothFormats(string system)
    {
        using var fixture = ScalarFormulaFixture.Load("excel-data-cleaning-" + system); var root = fixture.RootElement;
        var directory = Path.Combine(Path.GetTempPath(), "QuickExcel-cleaning-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "native.xlsx");
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in root.GetProperty("Parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
            var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(path));
            void Calculate(DocumentSession document) => Assert.True(document.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(document.BeginCalculation())));
            Calculate(session);
            foreach (var test in root.GetProperty("Cases").EnumerateArray())
            {
                var name = test.GetProperty("Sheet").GetString()!; var sheet = session.Snapshot.Sheets.Single(s => s.Name == name);
                var range = CellRectangle.Parse(test.GetProperty("Range").GetString()!); CleaningPreview preview;
                if (test.GetProperty("Operation").GetString() == "Duplicates")
                    preview = DataCleaning.RemoveDuplicates(session.Snapshot, sheet.Id, range, test.GetProperty("Keys").EnumerateArray().Select(v => range.Left + v.GetInt32() - 1), test.GetProperty("Header").GetBoolean());
                else
                {
                    var dest = CellRectangle.Parse(test.GetProperty("Destination").GetString()!);
                    var options = new TextSplitOptions(test.GetProperty("Delimiter").GetString() ?? ",", test.GetProperty("Collapse").GetBoolean(),
                        Breaks: test.GetProperty("Operation").GetString() == "Fixed" ? test.GetProperty("Fixed").EnumerateArray().Select(v => v.GetInt32()).Where(n => n > 0).ToImmutableArray() : default);
                    var types = test.GetProperty("Formats").EnumerateArray().Select(v => v.GetInt32() switch { 1 => CleaningColumnType.Number, 4 => CleaningColumnType.DateDmy, 5 => CleaningColumnType.DateYmd, _ => CleaningColumnType.Text }).ToArray();
                    if (name == "NumericChoice")
                    {
                        Assert.Throws<ArgumentException>(() => DataCleaning.SplitText(session.Snapshot, sheet.Id, range, new(dest.Top, dest.Left), options, types));
                        range = range with { Bottom = 9 }; // Explicitly reject Excel's General date/bool/16-digit inference.
                    }
                    preview = DataCleaning.SplitText(session.Snapshot, sheet.Id, range, new(dest.Top, dest.Left), options, types);
                }
                session.Execute(new ApplyCleaningCommand(preview)); session.Undo(); session.Redo(); Calculate(session);
            }
            Check(session.Snapshot);
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, "result." + extension); storage.Save(output, session.Snapshot, true);
                using var reopened = new DocumentSession(storage.Open(output)); Calculate(reopened); Check(reopened.Snapshot);
            }
            void Check(WorkbookSnapshot book)
            {
                foreach (var expectedSheet in root.GetProperty("After").EnumerateArray())
                {
                    var sheet = book.Sheets.Single(s => s.Name == expectedSheet.GetProperty("Sheet").GetString());
                    foreach (var expected in expectedSheet.GetProperty("Cells").EnumerateArray())
                    {
                        var p = CellRectangle.Parse(expected.GetProperty("Address").GetString()!);
                        if (sheet.Name == "NumericChoice" && p.Top >= 10 && p.Left == 2) { Assert.Null(sheet.GetCell(p.Top, p.Left)); continue; }
                        var cell = sheet.GetCell(p.Top, p.Left); var value = cell?.Value ?? CellValue.Blank;
                        Assert.Equal(expected.GetProperty("Kind").GetString(), value.Kind.ToString());
                        var wanted = expected.GetProperty("Value");
                        if (value.Kind == CellValueKind.Number) Assert.Equal(wanted.GetDouble(), value.NumberValue);
                        else Assert.Equal(wanted.ValueKind == JsonValueKind.Null ? "" : wanted.ValueKind == JsonValueKind.True ? "TRUE" : wanted.ValueKind == JsonValueKind.False ? "FALSE" : wanted.ToString(), CellInput.DisplayText(value));
                        string? Canonical(string? formula) => formula is null ? null : FormulaSyntax.Print(FormulaSyntax.Parse(formula));
                        Assert.Equal(Canonical(expected.GetProperty("Formula").GetString()), cell?.Content.Kind == CellContentKind.Formula ? Canonical(cell.Content.FormulaText) : null);
                        Assert.Equal(expected.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, book.Styles[sheet.GetStyleId(p.Top, p.Left)], book.DateSystem));
                    }
                }
            }
        }
        finally { Directory.Delete(directory, true); }
    }
}
