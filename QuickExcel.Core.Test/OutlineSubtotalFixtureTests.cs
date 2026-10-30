using System.IO.Compression;
using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class OutlineSubtotalFixtureTests
{
    [Fact]
    public void NativeMixedOutlineFilterRetainsOriginalVisibilityAndBlocksAmbiguousSubtotalRecalculation()
    {
        using var fixture = ScalarFormulaFixture.Load("excel-outline-filter-boundary"); var root = fixture.RootElement;
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-outline-boundary-" + Guid.NewGuid().ToString("N") + ".xlsx"); var saved = path + ".saved.xlsx";
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in root.GetProperty("Parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
            var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(path));
            Assert.Contains(session.Snapshot.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/autoFilter-outline");
            var sheet = session.Snapshot.Sheets[0]; Assert.Null(sheet.Data.Filter);
            Assert.All(Enumerable.Range(1, 11), row => Assert.True(sheet.Rows[row].Hidden));
            Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
            Assert.Equal(CalculationState.Blocked, session.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State);
            Assert.Equal(0, session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
            Assert.Throws<NotSupportedException>(() => session.Execute(new SetAutoFilterCommand(sheet.Id, null)));
            Assert.Throws<NotSupportedException>(() => session.Execute(new SetOutlineCollapsedCommand(sheet.Id, SheetAxis.Row, 1, 11, 1, false)));
            storage.Save(saved, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(saved));
        }
        finally { File.Delete(path); if (File.Exists(saved)) File.Delete(saved); }
    }
    [Theory]
    [InlineData("initial")]
    [InlineData("manual")]
    [InlineData("filtered")]
    [InlineData("collapsed")]
    [InlineData("expanded")]
    [InlineData("generated")]
    public void NativeExcelMatrixRecalculatesAllSubtotalModesAndPreservesOutline(string stageName)
    {
        using var fixture = ScalarFormulaFixture.Load("excel-outline-subtotal"); var root = fixture.RootElement;
        var stage = root.GetProperty("Stages").EnumerateArray().Single(s => s.GetProperty("Name").GetString() == stageName);
        Assert.True(stage.GetProperty("Stable").GetBoolean());
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-subtotal-native-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var parts = root.GetProperty("BaseParts").EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!);
            foreach (var removed in stage.GetProperty("Removed").EnumerateArray()) parts.Remove(removed.GetString()!);
            foreach (var changed in stage.GetProperty("Changes").EnumerateObject()) parts[changed.Name] = changed.Value.GetString()!;
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in parts) { using var writer = new StreamWriter(zip.CreateEntry(part.Key).Open(), new UTF8Encoding(false)); writer.Write(part.Value); }
            using var session = new DocumentSession(new TypedWorkbookStorage().Open(path));
            Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
            foreach (var expectedSheet in stage.GetProperty("Sheets").EnumerateArray())
            {
                var sheet = session.Snapshot.Sheets.Single(s => s.Name == expectedSheet.GetProperty("Sheet").GetString());
                Assert.Equal(expectedSheet.GetProperty("SummaryBelow").GetBoolean(), sheet.Data.Outline?.SummaryBelow ?? true);
                Assert.Equal(expectedSheet.GetProperty("SummaryRight").GetBoolean(), sheet.Data.Outline?.SummaryRight ?? true);
                var hidden = WorksheetFilterEngine.Evaluate(session.Snapshot, sheet.Id);
                foreach (var row in expectedSheet.GetProperty("Rows").EnumerateArray())
                {
                    var index = row.GetProperty("Index").GetInt32() - 1; var metadata = sheet.Rows.GetValueOrDefault(index);
                    Assert.Equal(row.GetProperty("Level").GetDouble(), (metadata?.OutlineLevel ?? 0) + 1);
                    Assert.Equal(row.GetProperty("Hidden").GetBoolean(), metadata?.Hidden == true || hidden.Contains(index));
                    Assert.Equal(row.GetProperty("Collapsed").GetBoolean(), metadata?.Collapsed ?? false);
                }
                foreach (var column in expectedSheet.GetProperty("Columns").EnumerateArray())
                {
                    var index = column.GetProperty("Index").GetInt32() - 1; var metadata = sheet.Columns.GetValueOrDefault(index);
                    Assert.Equal(column.GetProperty("Level").GetDouble(), (metadata?.OutlineLevel ?? 0) + 1);
                    Assert.Equal(column.GetProperty("Hidden").GetBoolean(), metadata?.Hidden ?? false);
                    Assert.Equal(column.GetProperty("Collapsed").GetBoolean(), metadata?.Collapsed ?? false);
                }
                foreach (var expected in expectedSheet.GetProperty("Cells").EnumerateArray().Where(c => c.GetProperty("Formula").ValueKind != System.Text.Json.JsonValueKind.Null))
                {
                    var p = CellRectangle.Parse(expected.GetProperty("Address").GetString()!); var cell = sheet.GetCell(p.Top, p.Left)!;
                    Assert.True(cell.Calculation?.State is CalculationState.Fresh or CalculationState.Error, expected.GetProperty("Formula").GetString());
                    Assert.Equal(expected.GetProperty("Kind").GetString(), cell.Value.Kind.ToString());
                    if (cell.Value.Kind == CellValueKind.Number)
                    {
                        var value = expected.GetProperty("Value").GetDouble();
                        Assert.InRange(Math.Abs(value - cell.Value.NumberValue), 0, Math.Max(1, Math.Abs(value)) * 1e-12);
                    }
                    else Assert.Equal(expected.GetProperty("Value").ToString(), CellInput.DisplayText(cell.Value));
                }
            }
        }
        finally { File.Delete(path); }
    }
}
