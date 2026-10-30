using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class ConditionalFormattingProbe
{
    public static int VerifyModel(string directory, string workbookPath, string verdictPath)
    {
        if (File.Exists(verdictPath)) throw new IOException("Use a new verdict path.");
        using var json = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(directory, "output.json")));
        using var source = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(directory, "source.json")));
        var oracle = json.RootElement;
        if (!oracle.GetProperty("Stable").GetBoolean() || !source.RootElement.GetProperty("Unchanged").GetBoolean() ||
            OpenXmlAudit.Hash(File.ReadAllBytes(workbookPath)) != source.RootElement.GetProperty("Sha256").GetString() ||
            OpenXmlAudit.Hash(File.ReadAllBytes(oracle.GetProperty("Path").GetString()!)) != oracle.GetProperty("Sha256").GetString())
            throw new InvalidDataException("Unbound native UI output evidence.");
        var book = new TypedWorkbookStorage().Open(workbookPath); var checks = new List<object>(); var diffs = new List<object>();
        foreach (var expected in oracle.GetProperty("After").EnumerateArray())
        {
            var sheet = book.Sheets.Single(s => s.Name == expected.GetProperty("Sheet").GetString());
            void Compare(string field, object? excel, object? model)
            {
                var pass = Equals(excel, model); var item = new { Sheet = sheet.Name, Field = field, Expected = excel, Actual = model, Passed = pass };
                checks.Add(item); if (!pass) diffs.Add(item);
            }
            Compare("rules", expected.GetProperty("Rules").GetArrayLength(), sheet.Data.ConditionalFormats.Length);
            foreach (var cell in expected.GetProperty("Cells").EnumerateArray())
            {
                var address = cell.GetProperty("Address").GetString()!; var p = CellRectangle.Parse(address); var style = ConditionalPresentation.GetStyle(book, sheet, p.Top, p.Left);
                Compare(address + "/text", cell.GetProperty("Text").GetString(), CellDisplayFormatter.Format(sheet.GetCell(p.Top, p.Left)?.Value ?? CellValue.Blank, style, book.DateSystem));
                Compare(address + "/fill", cell.GetProperty("FillColor").GetString(), style.FillColor);
                Compare(address + "/font", cell.GetProperty("FontColor").GetString(), style.FontColor ?? "FF000000");
                Compare(address + "/bold", cell.GetProperty("Bold").GetBoolean(), style.Bold);
                Compare(address + "/italic", cell.GetProperty("Italic").GetBoolean(), style.Italic);
                Compare(address + "/strike", cell.GetProperty("Strike").GetBoolean(), style.Strike);
            }
        }
        foreach (var error in ConditionalPresentation.Diagnostics(book)) diffs.Add(error);
        File.WriteAllText(verdictPath, JsonSerializer.Serialize(new { Passed = diffs.Count == 0, Count = checks.Count, Diffs = diffs, Checks = checks }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} model/native output checks, {diffs.Count} differences."); return diffs.Count == 0 ? 0 : 1;
    }
    public static int Verify(string expectedPath, string actualDirectory, string sourcePath, string verdictPath)
    {
        if (File.Exists(verdictPath)) throw new IOException("Use a new verdict path.");
        using var expectedJson = JsonDocument.Parse(File.ReadAllBytes(expectedPath));
        using var actualJson = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(actualDirectory, "output.json")));
        using var sourceJson = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(actualDirectory, "source.json")));
        var expected = expectedJson.RootElement; var actual = actualJson.RootElement; var source = sourceJson.RootElement;
        if (!expected.GetProperty("Stable").GetBoolean() || !actual.GetProperty("Stable").GetBoolean() || !source.GetProperty("Unchanged").GetBoolean() ||
            OpenXmlAudit.Hash(File.ReadAllBytes(sourcePath)) != source.GetProperty("Sha256").GetString() ||
            OpenXmlAudit.Hash(File.ReadAllBytes(actual.GetProperty("Path").GetString()!)) != actual.GetProperty("Sha256").GetString() ||
            OpenXmlAudit.Hash(File.ReadAllBytes(expected.GetProperty("Path").GetString()!)) != expected.GetProperty("Sha256").GetString())
            throw new InvalidDataException("Unbound or unstable native evidence.");
        var checks = new List<object>(); var diffs = new List<object>();
        void Compare(string field, object? before, object? after)
        {
            var pass = Equals(before, after); var item = new { Field = field, Expected = before, Actual = after, Passed = pass };
            checks.Add(item); if (!pass) diffs.Add(item);
        }
        foreach (var sheet in expected.GetProperty("After").EnumerateArray())
        {
            var name = sheet.GetProperty("Sheet").GetString(); var other = actual.GetProperty("After").EnumerateArray().Single(s => s.GetProperty("Sheet").GetString() == name);
            Compare(name + "/rules", sheet.GetProperty("Rules").GetArrayLength(), other.GetProperty("Rules").GetArrayLength());
            foreach (var rule in sheet.GetProperty("Rules").EnumerateArray())
            {
                var priority = rule.GetProperty("Priority").GetInt32();
                var matching = other.GetProperty("Rules").EnumerateArray().Single(r => r.GetProperty("Priority").GetInt32() == priority);
                foreach (var field in new[] { "Range", "Priority", "Stop", "Type", "Operator", "Text", "Formula1", "Formula2" })
                {
                    string? Canonical(JsonElement item)
                    {
                        if (item.ValueKind == JsonValueKind.Null) return null; var text = item.ToString();
                        if (field.StartsWith("Formula"))
                        { try { return FormulaSyntax.Print(FormulaSyntax.Parse(text.StartsWith('=') ? text : "=" + text)); } catch (FormatException) { } }
                        return text;
                    }
                    Compare(name + "/rule" + priority + "/" + field, Canonical(rule.GetProperty(field)), Canonical(matching.GetProperty(field)));
                }
            }
            var cells = other.GetProperty("Cells").EnumerateArray().ToDictionary(c => c.GetProperty("Address").GetString()!);
            foreach (var cell in sheet.GetProperty("Cells").EnumerateArray())
            {
                var address = cell.GetProperty("Address").GetString()!; var match = cells[address];
                foreach (var field in new[] { "Kind", "Value", "Text", "Bold", "Italic", "Strike", "Underline", "FontColor", "FillColor", "NumberFormat" })
                    Compare(name + "/" + address + "/" + field, cell.GetProperty(field).ToString(), match.GetProperty(field).ToString());
                for (var i = 0; i < 4; i++)
                {
                    var edge = cell.GetProperty("Borders")[i]; var next = match.GetProperty("Borders")[i];
                    Compare(name + "/" + address + "/border" + i, edge.GetProperty("Line").GetInt32(), next.GetProperty("Line").GetInt32());
                    if (edge.GetProperty("Line").GetInt32() == -4142) continue;
                    foreach (var field in new[] { "Weight", "Color" }) Compare(name + "/" + address + "/border" + i + "/" + field, edge.GetProperty(field).ToString(), next.GetProperty(field).ToString());
                }
            }
        }
        File.WriteAllText(verdictPath, JsonSerializer.Serialize(new { Passed = diffs.Count == 0, Count = checks.Count, Diffs = diffs, Checks = checks }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} native Excel output checks, {diffs.Count} differences."); return diffs.Count == 0 ? 0 : 1;
    }
    public static int Run(string oracleDirectory, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new evidence directory.");
        Directory.CreateDirectory(directory);
        var storage = new TypedWorkbookStorage(); var checks = new List<object>(); var diffs = new List<object>(); var outputs = new List<object>();
        using var commands = new DocumentSession(storage.Open(Path.Combine(oracleDirectory, "initial.xlsx")));
        var structure = commands.Snapshot.Sheets.Single(s => s.Name == "Structure").Id;
        foreach (var name in new[] { "initial", "edited", "insert-row", "delete-first", "copy-sheet" })
        {
            var oraclePath = Path.Combine(oracleDirectory, name + ".json");
            using var json = JsonDocument.Parse(File.ReadAllBytes(oraclePath)); var oracle = json.RootElement;
            var path = Path.Combine(oracleDirectory, name + ".xlsx"); var source = File.ReadAllBytes(path);
            if (!oracle.GetProperty("Stable").GetBoolean() || OpenXmlAudit.Hash(source) != oracle.GetProperty("Sha256").GetString())
                throw new InvalidDataException("Unbound or unstable Excel oracle.");
            var opened = storage.Open(path);
            void Check(WorkbookSnapshot snapshot, string stage)
            {
                foreach (var item in oracle.GetProperty("After").EnumerateArray())
                {
                    var sheet = snapshot.Sheets.Single(s => s.Name == item.GetProperty("Sheet").GetString());
                    void Compare(string field, object? expected, object? actual)
                    {
                        var pass = Equals(expected, actual); var record = new { Phase = name, Stage = stage, Sheet = sheet.Name, Field = field, Expected = expected, Actual = actual, Passed = pass };
                        checks.Add(record); if (!pass) diffs.Add(record);
                    }
                    Compare("rules", item.GetProperty("Rules").GetArrayLength(), sheet.Data.ConditionalFormats.Length);
                    foreach (var expected in item.GetProperty("Rules").EnumerateArray())
                    {
                        var priority = expected.GetProperty("Priority").GetInt32(); var rule = sheet.Data.ConditionalFormats.SingleOrDefault(r => r.Priority == priority);
                        Compare("rule/" + priority, true, rule is not null); if (rule is null) continue;
                        Compare("rule/" + priority + "/range", expected.GetProperty("Range").GetString(), string.Join(",", rule.Ranges));
                        Compare("rule/" + priority + "/stop", expected.GetProperty("Stop").GetBoolean(), rule.StopIfTrue);
                    }
                    foreach (var cell in item.GetProperty("Cells").EnumerateArray())
                    {
                        var address = cell.GetProperty("Address").GetString()!; var p = CellRectangle.Parse(address);
                        var style = ConditionalPresentation.GetStyle(snapshot, sheet, p.Top, p.Left);
                        var value = sheet.GetCell(p.Top, p.Left)?.Value ?? CellValue.Blank;
                        Compare(address + "/text", cell.GetProperty("Text").GetString(), CellDisplayFormatter.Format(value, style, snapshot.DateSystem));
                        Compare(address + "/bold", cell.GetProperty("Bold").GetBoolean(), style.Bold);
                        Compare(address + "/italic", cell.GetProperty("Italic").GetBoolean(), style.Italic);
                        Compare(address + "/strike", cell.GetProperty("Strike").GetBoolean(), style.Strike);
                        Compare(address + "/underline", cell.GetProperty("Underline").GetInt32(), style.Underline switch { "single" => 2, "double" => -4119, "singleAccounting" => 4, "doubleAccounting" => 5, _ => -4142 });
                        Compare(address + "/font", cell.GetProperty("FontColor").GetString(), style.FontColor ?? "FF000000");
                        Compare(address + "/fill", cell.GetProperty("FillColor").GetString(), style.FillColor);
                        var edges = new[] { (style.LeftBorder, style.LeftBorderColor), (style.RightBorder, style.RightBorderColor), (style.TopBorder, style.TopBorderColor), (style.BottomBorder, style.BottomBorderColor) };
                        for (var i = 0; i < 4; i++)
                        {
                            var edge = cell.GetProperty("Borders")[i]; var line = edge.GetProperty("Line").GetInt32(); var weight = edge.GetProperty("Weight").GetInt32();
                            var kind = line switch { -4142 => 0, 1 => weight switch { 1 => 7, 2 => 1, -4138 => 2, 4 => 5, _ => -1 }, -4115 => 3, -4118 => 4, -4119 => 6, _ => -1 };
                            Compare(address + "/edge" + i, kind, edges[i].Item1);
                            if (kind != 0) Compare(address + "/edgeColor" + i, edge.GetProperty("Color").GetString(), edges[i].Item2 ?? "FF000000");
                        }
                    }
                }
                foreach (var error in ConditionalPresentation.Diagnostics(snapshot)) diffs.Add(new { Phase = name, Stage = stage, Error = error });
            }
            Check(opened, "import");
            switch (name)
            {
                case "edited": commands.Execute(new SetCellsCommand([new(new(structure, 2, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))])); break;
                case "insert-row": commands.Execute(new ShiftAxisCommand(structure, SheetAxis.Row, 3, 1)); break;
                case "delete-first": commands.Execute(new ShiftAxisCommand(structure, SheetAxis.Row, 1, 1, true)); break;
                case "copy-sheet": commands.Execute(new CopySheetCommand(structure, "Copy", commands.Snapshot.Sheets.Length)); break;
            }
            Check(commands.Snapshot, "commands");
            if (name != "initial") { commands.Undo(); commands.Redo(); Check(commands.Snapshot, "redo"); }
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var output = Path.Combine(directory, name + "." + extension);
                // Force serialization: unchanged imported packages otherwise take the original-package preservation path.
                using var session = new DocumentSession(opened);
                session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, 60, 20), CellContent.FromLiteral(CellValue.FromText("CF serialization evidence")))]));
                var scoped = false;
                try { storage.Save(output, session.Snapshot); } catch (WorkbookCompatibilityException) { scoped = true; storage.Save(output, session.Snapshot, true); }
                Check(storage.Open(output), extension);
                var issues = extension == "xlsx" ? OpenXmlAudit.Inspect(File.ReadAllBytes(output)).Issues.Select(i => i.ToString()).ToArray() : [];
                outputs.Add(new { Path = output, ScopedDiagnosticCopy = scoped, Issues = issues }); diffs.AddRange(issues);
            }
            var commandPath = Path.Combine(directory, name + "-commands.xlsx");
            storage.Save(commandPath, commands.Snapshot, true); Check(storage.Open(commandPath), "command-save");
            var commandIssues = OpenXmlAudit.Inspect(File.ReadAllBytes(commandPath)).Issues.Select(i => i.ToString()).ToArray();
            outputs.Add(new { Path = commandPath, Issues = commandIssues }); diffs.AddRange(commandIssues);
            if (!source.SequenceEqual(File.ReadAllBytes(path))) throw new InvalidDataException("Source was modified.");
        }
        File.WriteAllText(Path.Combine(directory, "conditional-formatting.json"), JsonSerializer.Serialize(new { Passed = diffs.Count == 0, Count = checks.Count, Diffs = diffs, Checks = checks, Outputs = outputs }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{checks.Count} conditional format checks, {diffs.Count} differences."); return diffs.Count == 0 ? 0 : 1;
    }
}
