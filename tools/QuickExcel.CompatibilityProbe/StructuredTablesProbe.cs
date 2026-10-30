using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class StructuredTablesProbe
{
    private sealed class NoScripts : IScriptRuntimeFactory { public IScriptRuntime Create(string source) => throw new InvalidOperationException("Scripts disabled."); }
    public static int Run(string oracleDirectory, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new directory."); Directory.CreateDirectory(directory);
        string[] stages = ["initial", "rename", "insert-row", "insert-column", "delete-inserted", "copy-sheet", "copy-formula", "expand", "insert-first", "delete-column"];
        var storage = new TypedWorkbookStorage(); var checks = new List<object>(); var diffs = new List<object>(); var paths = new List<object>();
        void Compare(string phase, string field, object? expected, object? actual)
        {
            var passed = expected is double e && actual is double a ? Math.Abs(e - a) <= 1e-12 * Math.Max(1, Math.Abs(e)) : Equals(expected, actual);
            var record = new { Phase = phase, Field = field, Expected = expected, Actual = actual, Passed = passed }; checks.Add(record); if (!passed) diffs.Add(record);
        }
        JsonElement Oracle(string stage)
        {
            var json = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(oracleDirectory, stage + ".json"))).RootElement;
            var path = Path.Combine(oracleDirectory, stage + ".xlsx");
            if (!json.GetProperty("Stable").GetBoolean() || OpenXmlAudit.Hash(File.ReadAllBytes(path)) != json.GetProperty("Sha256").GetString()) throw new InvalidDataException("Unbound oracle.");
            return json;
        }
        void Calculate(DocumentSession session)
        { if (!session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation()))) throw new InvalidOperationException("Calculation was not committed."); }
        void Check(WorkbookSnapshot snapshot, string phase, JsonElement oracle)
        {
            Compare(phase, "read-only", false, snapshot.FileState?.FormulaReadOnly ?? false);
            Compare(phase, "sheet count", oracle.GetProperty("After").GetArrayLength(), snapshot.Sheets.Length);
            foreach (var sheetRecord in oracle.GetProperty("After").EnumerateArray())
            {
                var name = sheetRecord.GetProperty("Sheet").GetString()!; var sheet = snapshot.Sheets.Single(s => s.Name == name);
                Compare(phase, name + "/table count", sheetRecord.GetProperty("Tables").GetArrayLength(), sheet.Data.Tables.Length);
                foreach (var tableRecord in sheetRecord.GetProperty("Tables").EnumerateArray())
                {
                    var tableName = tableRecord.GetProperty("Name").GetString(); var table = sheet.Data.Tables.SingleOrDefault(t => t.Name == tableName);
                    Compare(phase, name + "/" + tableName, true, table is not null); if (table is null) continue;
                    Compare(phase, name + "/range", tableRecord.GetProperty("Range").GetString(), table.Range.ToString());
                    Compare(phase, name + "/columns", string.Join("|", tableRecord.GetProperty("Headers").EnumerateArray().Select(h => h.GetString())), string.Join("|", table.Columns.Select(c => c.Name)));
                    Compare(phase, name + "/totals", tableRecord.GetProperty("Totals").GetBoolean(), table.TotalsRow);
                }
                var expectedAddresses = new HashSet<CellPosition>();
                foreach (var record in sheetRecord.GetProperty("Cells").EnumerateArray())
                {
                    var address = record.GetProperty("Address").GetString()!; var rectangle = CellRectangle.Parse(address); var pos = new CellPosition(rectangle.Top, rectangle.Left); expectedAddresses.Add(pos);
                    var cell = sheet.GetCell(pos.Row, pos.Column); var value = cell?.Value ?? CellValue.Blank; var key = name + "!" + address;
                    Compare(phase, key + "/kind", record.GetProperty("Kind").GetString(), value.Kind.ToString());
                    var raw = record.GetProperty("Value"); object? expected = raw.ValueKind switch { JsonValueKind.Number => raw.GetDouble(), JsonValueKind.True => true, JsonValueKind.False => false, JsonValueKind.Null => null, _ => raw.GetString() };
                    object? actual = value.Kind switch { CellValueKind.Number => value.NumberValue, CellValueKind.Text => value.TextValue, CellValueKind.Boolean => value.BooleanValue, CellValueKind.Blank => null, _ => CellInput.DisplayText(value) };
                    Compare(phase, key + "/value", expected, actual);
                    var formula = record.GetProperty("Formula").GetString(); Compare(phase, key + "/formula", formula is not null, cell?.Content.Kind == CellContentKind.Formula);
                    if (formula is not null && cell?.Content.Kind == CellContentKind.Formula)
                    {
                        string Bound(string text)
                        { try { return FormulaSyntax.Print(FormulaBinding.Bind(FormulaSyntax.Parse(text), snapshot, sheet.Id, pos)); } catch (FormatException ex) { return "UNSUPPORTED: " + text + "/" + ex.Message; } }
                        Compare(phase, key + "/binding", Bound(formula), Bound(cell.Content.FormulaText));
                        Compare(phase, key + "/fresh", true, cell.Calculation is { State: CalculationState.Fresh or CalculationState.Error });
                    }
                }
                foreach (var (pos, cell) in sheet.Cells.Where(p => !expectedAddresses.Contains(p.Key) && (p.Value.Content.Kind == CellContentKind.Formula || p.Value.Value.Kind != CellValueKind.Blank)))
                    Compare(phase, name + "/extra/" + pos, null, cell.Content.ToString());
            }
        }
        void Save(DocumentSession session, string stage, JsonElement oracle)
        {
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, stage + "." + extension); var scoped = false;
                try { scoped = stage == "initial"; storage.Save(output, session.Snapshot, scoped); } catch (WorkbookCompatibilityException) { scoped = true; storage.Save(output, session.Snapshot, true); }
                using var reopened = new DocumentSession(storage.Open(output)); Calculate(reopened); Check(reopened.Snapshot, stage + "/" + extension, oracle);
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(output)).Issues.Select(i => i.ToString()).ToArray() : [];
                paths.Add(new { Path = output, ScopedDiagnosticCopy = scoped, Issues = issues }); if (issues.Length != 0) diffs.AddRange(issues);
            }
        }
        using var session = new DocumentSession(storage.Open(Path.Combine(oracleDirectory, "initial.xlsx")));
        foreach (var stage in stages)
        {
            var oracle = Oracle(stage);
            using (var native = new DocumentSession(storage.Open(Path.Combine(oracleDirectory, stage + ".xlsx"))))
            { Calculate(native); Check(native.Snapshot, stage + "/native-import", oracle); }
            var data = session.Snapshot.Sheets.Single(s => s.Name == "Data"); var table = data.Data.Tables.Single(); var commands = 0;
            void Execute(DocumentCommand command) { session.Execute(command); commands++; }
            switch (stage)
            {
                case "rename": Execute(new RenameTableColumnCommand(data.Id, table.Id, table.Columns[1].Id, "Unit Price")); Execute(new RenameTableCommand(data.Id, table.Id, "Revenue")); break;
                case "insert-row": Execute(new ShiftAxisCommand(data.Id, SheetAxis.Row, 3, 1)); break;
                case "insert-column": Execute(new ShiftAxisCommand(data.Id, SheetAxis.Column, 2, 1)); break;
                case "delete-inserted": Execute(new ShiftAxisCommand(data.Id, SheetAxis.Column, 2, 1, true)); Execute(new ShiftAxisCommand(data.Id, SheetAxis.Row, 3, 1, true)); break;
                case "copy-sheet":
                    var copy = new CopySheetCommand(data.Id, "Copy", session.Snapshot.Sheets.Length); Execute(copy);
                    var copiedTable = session.Snapshot.GetSheet(copy.Id).Data.Tables.Single(); Execute(new RenameTableCommand(copy.Id, copiedTable.Id, "Revenue_Copy"));
                    Execute(new SetCellsCommand([new(new(copy.Id, 2, 1), CellContent.FromLiteral(CellValue.FromNumber(7)))])); break;
                case "copy-formula":
                    var clipboard = WorkbookClipboard.Capture(session.Snapshot, data.Id, new(2, 3, 2, 3));
                    Execute(new PasteCellsCommand(clipboard, session.Snapshot.Sheets.Single(s => s.Name == "Summary").Id, new(11, 3, 11, 3), session.Revision));
                    Execute(new PasteCellsCommand(clipboard, data.Id, new(2, 10, 2, 10), session.Revision)); break;
                case "expand":
                    var a1 = session.Snapshot.Sheets.Single(s => s.Name == "A1Column"); Execute(new ResizeTableCommand(a1.Id, a1.Data.Tables.Single().Id, 4)); break;
                case "insert-first": Execute(new ShiftAxisCommand(session.Snapshot.Sheets.Single(s => s.Name == "A1Column").Id, SheetAxis.Row, 1, 1)); break;
                case "delete-column": Execute(new ShiftAxisCommand(data.Id, SheetAxis.Column, 2, 1, true)); break;
            }
            Calculate(session); Check(session.Snapshot, stage + "/edited", oracle); Save(session, stage, oracle);
            if (commands > 0)
            {
                for (var i = 0; i < commands; i++) session.Undo(); Calculate(session); Check(session.Snapshot, stage + "/undo", Oracle(stages[Array.IndexOf(stages, stage) - 1]));
                for (var i = 0; i < commands; i++) session.Redo(); Calculate(session); Check(session.Snapshot, stage + "/redo", oracle);
            }
        }
        File.WriteAllText(Path.Combine(directory, "structured-tables.json"), JsonSerializer.Serialize(new { Passed = diffs.Count == 0, Count = checks.Count, Diffs = diffs, Checks = checks, Paths = paths }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} structured table checks, {diffs.Count} differences."); return diffs.Count == 0 ? 0 : 1;
    }
}
