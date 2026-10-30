using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

internal static class ScalarFormulaFixture
{
    public static JsonDocument Load(string name) => JsonDocument.Parse(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".json")));
    public static CellContent Content(JsonElement item) => item.GetProperty("Kind").GetString() == "Formula"
        ? CellContent.FromFormula(item.GetProperty("Value").GetString()!) : CellContent.FromLiteral(Value(item));
    public static CellValue Value(JsonElement item) => item.GetProperty("Kind").GetString() switch
    {
        "Number" => CellValue.FromNumber(item.GetProperty("Value").GetDouble()), "Text" => CellValue.FromText(item.GetProperty("Value").GetString()!),
        "Boolean" => CellValue.FromBoolean(item.GetProperty("Value").GetBoolean()), "Error" => CellInput.FromErrorText(item.GetProperty("Value").GetString()!), _ => CellValue.Blank
    };
    public static CellPosition Position(string address) { var r = (ReferenceNode)FormulaSyntax.Parse("=" + address); return new(r.Row, r.Column); }
    public static WorkbookSnapshot Book(JsonElement spec)
    {
        var sheets = spec.GetProperty("Sheets").EnumerateArray().Select(s => s.GetString()!).ToDictionary(s => s, _ => new Dictionary<CellPosition, CellSnapshot>());
        foreach (var input in spec.GetProperty("Inputs").EnumerateArray()) sheets[input.GetProperty("Sheet").GetString()!][Position(input.GetProperty("Address").GetString()!)] = new(Content(input));
        foreach (var item in spec.GetProperty("Cases").EnumerateArray()) sheets[item.GetProperty("Sheet").GetString()!][Position(item.GetProperty("Address").GetString()!)] = F(item.GetProperty("Formula").GetString()!);
        return new(Guid.NewGuid(), 0, spec.GetProperty("Date1904").GetBoolean() ? ExcelDateSystem.Mac1904 : ExcelDateSystem.Windows1900,
            sheets.Select(p => new WorksheetSnapshot(SheetId.New(), p.Key, p.Value)),
            textCompatibilityVersion: spec.TryGetProperty("TextCompatibilityVersion", out var version) ? version.GetInt32() : 1);
    }
    public static void Edit(DocumentSession session, JsonElement spec) => session.Execute(new SetCellsCommand(spec.GetProperty("Edits").EnumerateArray().Select(e =>
    {
        var sheet = session.Snapshot.Sheets.Single(s => s.Name == e.GetProperty("Sheet").GetString()); var p = Position(e.GetProperty("Address").GetString()!);
        return new CellEdit(new(sheet.Id, p.Row, p.Column), Content(e));
    })));
    public static void Calculate(DocumentSession session, WorkbookCalculator calculator)
    { Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation()))); }
    public static void Compare(WorkbookSnapshot book, JsonElement spec, JsonElement expected)
    {
        var cases = spec.GetProperty("Cases").EnumerateArray().ToDictionary(c => c.GetProperty("Id").GetString()!); var differences = new List<string>();
        Assert.Equal(cases.Count, expected.GetArrayLength());
        foreach (var item in expected.EnumerateArray())
        {
            var id = item.GetProperty("Id").GetString()!; var c = cases[id]; var p = Position(c.GetProperty("Address").GetString()!);
            var cell = book.Sheets.Single(s => s.Name == c.GetProperty("Sheet").GetString()).GetCell(p.Row, p.Column)!; var value = Value(item);
            var same = cell.Value.Kind == value.Kind && (value.Kind == CellValueKind.Number ? Math.Abs(cell.Value.NumberValue - value.NumberValue) <= 1e-12 * Math.Max(1, Math.Abs(value.NumberValue)) : cell.Value == value);
            if (!same || cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error)) differences.Add($"{id} {c.GetProperty("Formula").GetString()}: expected {value}, actual {cell.Value} ({cell.Calculation?.State})");
        }
        Assert.True(differences.Count == 0, $"{differences.Count} mismatches\n" + string.Join("\n", differences.Take(100)));
    }
}
