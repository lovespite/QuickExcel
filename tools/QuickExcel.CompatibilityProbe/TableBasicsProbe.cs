using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class TableBasicsProbe
{
    public static int Run(string oraclePath, string directory, string? editedOraclePath = null)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new directory."); Directory.CreateDirectory(directory);
        using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
        var path = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(oraclePath))!, "tables.xlsx"); var source = File.ReadAllBytes(path);
        if (!oracle.GetProperty("Stable").GetBoolean() || OpenXmlAudit.Hash(source) != oracle.GetProperty("Sha256").GetString()) throw new InvalidDataException("Unbound or unstable oracle.");
        var storage = new TypedWorkbookStorage(); var opened = storage.Open(path); var checks = new List<object>(); var diffs = new List<object>(); var paths = new List<object>();
        void Check(WorkbookSnapshot snapshot, string stage, JsonElement? expectedRecords = null)
        {
            foreach (var item in (expectedRecords ?? oracle.GetProperty("After")).EnumerateArray())
            {
                var sheet = snapshot.Sheets.Single(s => s.Name == item.GetProperty("Sheet").GetString()); var name = item.GetProperty("Name").GetString();
                var table = sheet.Data.Tables.SingleOrDefault(t => t.Name == name);
                void Compare(string field, object? expected, object? actual)
                {
                    var pass = Equals(expected, actual); var record = new { Stage = stage, Table = name, Field = field, Expected = expected, Actual = actual, Passed = pass };
                    checks.Add(record); if (!pass) diffs.Add(record);
                }
                Compare("modeled", true, table is not null); if (table is null) continue;
                Compare("range", item.GetProperty("Range").GetString(), table.Range.ToString()); Compare("header", item.GetProperty("Header").GetBoolean(), table.HeaderRow);
                Compare("totals", item.GetProperty("Totals").GetBoolean(), table.TotalsRow); Compare("style", item.GetProperty("Style").GetString(), table.Style.Name);
                Compare("first", item.GetProperty("First").GetBoolean(), table.Style.FirstColumn); Compare("last", item.GetProperty("Last").GetBoolean(), table.Style.LastColumn);
                Compare("row stripes", item.GetProperty("Rows").GetBoolean(), table.Style.RowStripes); Compare("column stripes", item.GetProperty("Columns").GetBoolean(), table.Style.ColumnStripes);
                Compare("column count", item.GetProperty("Headers").GetArrayLength(), table.Columns.Length);
                var headers = item.GetProperty("Headers").EnumerateArray().ToArray();
                for (var c = 0; c < Math.Min(headers.Length, table.Columns.Length); c++) Compare("column/" + c, headers[c].GetProperty("Name").GetString(), table.Columns[c].Name);
                var hidden = WorksheetFilterEngine.Evaluate(snapshot, sheet.Id, true);
                foreach (var cell in item.GetProperty("Cells").EnumerateArray())
                {
                    var address = cell.GetProperty("Address").GetString()!; var pos = CellRectangle.Parse(address); var style = TablePresentation.GetStyle(snapshot, sheet, pos.Top, pos.Left);
                    var value = sheet.GetCell(pos.Top, pos.Left)?.Value ?? CellValue.Blank;
                    Compare(address + "/text", cell.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, style, snapshot.DateSystem));
                    Compare(address + "/bold", cell.GetProperty("Bold").GetBoolean(), style.Bold);
                    Compare(address + "/font", cell.GetProperty("Font").GetString(), style.FontColor ?? "FF000000");
                    Compare(address + "/fill", cell.GetProperty("Fill").GetString(), style.FillColor ?? "FFFFFFFF");
                    Compare(address + "/hidden", cell.GetProperty("Hidden").GetBoolean(), hidden.Contains(pos.Top) || sheet.Rows.GetValueOrDefault(pos.Top)?.Hidden == true);
                    if (cell.TryGetProperty("Borders", out var borders))
                        foreach (var (side, kind, color) in new[] { ("Left", style.LeftBorder, style.LeftBorderColor), ("Right", style.RightBorder, style.RightBorderColor),
                            ("Top", style.TopBorder, style.TopBorderColor), ("Bottom", style.BottomBorder, style.BottomBorderColor) })
                        {
                            var expected = borders.GetProperty(side); Compare(address + "/border-" + side, expected.GetProperty("Style").GetInt32(), kind);
                            if (expected.GetProperty("Style").GetInt32() != 0) Compare(address + "/border-color-" + side, expected.GetProperty("Color").GetString(), color ?? "FF000000");
                        }
                }
            }
        }
        Check(opened, "import"); using var session = new DocumentSession(opened);
        session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, 19, 6), CellContent.FromLiteral(CellValue.FromText("outside edit")))]));
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var output = Path.Combine(directory, "tables." + extension); var scoped = false;
            try { storage.Save(output, session.Snapshot); } catch (WorkbookCompatibilityException) { scoped = true; storage.Save(output, session.Snapshot, true); }
            Check(storage.Open(output), extension); var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(output)).Issues.Select(i => i.ToString()).ToArray() : [];
            paths.Add(new { Path = output, ScopedDiagnosticCopy = scoped, Issues = issues }); if (issues.Length > 0) diffs.AddRange(issues);
        }
        if (editedOraclePath is not null)
        {
            using var edited = JsonDocument.Parse(File.ReadAllBytes(editedOraclePath)); var expected = edited.RootElement;
            if (!expected.GetProperty("Stable").GetBoolean() || !expected.GetProperty("SourceUnchanged").GetBoolean() || expected.GetProperty("SourceSha256").GetString() != OpenXmlAudit.Hash(source))
                throw new InvalidDataException("Edited oracle does not bind the initial source.");
            var sheet = session.Snapshot.Sheets[0];
            session.Execute(new SetCellsCommand([new(new(sheet.Id, 2, 2), CellContent.FromLiteral(CellValue.FromNumber(25)))]));
            session.Execute(new ResizeTableCommand(sheet.Id, sheet.Data.Tables.Single().Id, 8));
            Check(session.Snapshot, "edit", expected.GetProperty("After"));
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, "edited." + extension); var scoped = false;
                try { storage.Save(output, session.Snapshot); } catch (WorkbookCompatibilityException) { scoped = true; storage.Save(output, session.Snapshot, true); }
                Check(storage.Open(output), "edited-" + extension, expected.GetProperty("After"));
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(output)).Issues.Select(i => i.ToString()).ToArray() : [];
                paths.Add(new { Path = output, ScopedDiagnosticCopy = scoped, Issues = issues }); if (issues.Length > 0) diffs.AddRange(issues);
            }
            session.Undo(); session.Undo(); Check(session.Snapshot, "undo"); session.Redo(); session.Redo(); Check(session.Snapshot, "redo", expected.GetProperty("After"));
        }
        File.WriteAllText(Path.Combine(directory, "table-basics.json"), JsonSerializer.Serialize(new { Passed = diffs.Count == 0, Count = checks.Count, Diffs = diffs, Checks = checks, Paths = paths,
            SourceSha256 = OpenXmlAudit.Hash(source), OracleSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(oraclePath)), SourceUnchanged = source.SequenceEqual(File.ReadAllBytes(path)), Diagnostics = opened.FileState?.CompatibilityDiagnostics }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} table checks, {diffs.Count} differences."); return diffs.Count == 0 ? 0 : 1;
    }
}
