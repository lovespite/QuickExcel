using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class DelayedClipboard(ClipboardExchange value, Action change) : IClipboardExchange
    {
        public Task<ClipboardExchange> ReadAsync() { change(); return Task.FromResult(value); }
        public Task WriteAsync(ClipboardExchange value) => Task.CompletedTask;
    }
    private async Task RunExternalClipboardChecksAsync(Action<bool, string> check)
    {
        var original = _clipboard;
        try
        {
            _workspace.New(); ShowBook(); var memory = new MemoryClipboard(); _clipboard = memory;
            var source = new ExternalClipboardRegion(1, 3, [new(CellValue.FromText("00123"), new() { Bold = true }), new(CellValue.FromNumber(.12), new() { NumberFormatCode = "0.00%" }), new(CellValue.Blank, new())]);
            memory.Value = new("wrong text", SpreadsheetXml: SpreadsheetClipboardCodec.Write(source), Html: HtmlClipboardCodec.Write(source));
            _grid.Select(new(2, 2)); await PasteAsync("all");
            check(Document.Snapshot.Sheets[0].GetCell(2, 2)!.Value == CellValue.FromText("00123") && Document.Snapshot.Sheets[0].GetCell(2, 3)!.Value == CellValue.FromNumber(.12), "External XML takes priority and handler preserves source types instead of TSV display text");
            await History(false); check(Document.Snapshot.Sheets[0].Cells.Count == 0, "One external paste undo restores all cells and styles"); await History(true);
            _grid.SelectRange(new(2, 2, 2, 4), new(2, 2)); await CopyAsync(false);
            check(memory.Value.SpreadsheetXml is not null && memory.Value.Html is not null && memory.Value.Token is not null, "Copy handler publishes XML, CF_HTML, text and local token together");
            _grid.Select(new(4, 2)); await PasteAsync("all");
            check(Document.Snapshot.Sheets[0].GetCell(4, 2)!.Value == CellValue.FromText("00123"), "Matching local token still selects the existing internal capture");
            memory.Value = new("replacement", SpreadsheetXml: SpreadsheetClipboardCodec.Write(new(1, 1, [new(CellValue.FromBoolean(true), new())])));
            _grid.Select(new(6, 2)); await PasteAsync("values");
            check(Document.Snapshot.Sheets[0].GetCell(6, 2)!.Value == CellValue.FromBoolean(true), "Clipboard replacement cannot reuse the old capture");
            var before = Document.Snapshot; memory.Value = new("fallback", SpreadsheetXml: [1, 2, 3]);
            try { await PasteAsync("all"); throw new Exception("Damaged XML accepted"); } catch (System.Xml.XmlException) { }
            check(ReferenceEquals(before, Document.Snapshot), "Damaged rich payload does not silently use text or change history");
            _clipboard = new DelayedClipboard(new("late"), () => _grid.Select(new(8, 8)));
            try { await PasteAsync("all"); throw new Exception("Changed selection accepted"); } catch (InvalidOperationException) { }
            check(before.Revision == Document.Snapshot.Revision && ReferenceEquals(before.Sheets[0].Cells, Document.Snapshot.Sheets[0].Cells), "Async selection change rejects the pending paste");
            _clipboard = memory; _workspace.New(); ShowBook();
            var sheet = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([new(new(sheet, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(60)))]));
            await ExecuteAsync(new SetCellStyleCommand([new(sheet, 0, 0)], new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd")));
            _grid.Select(new(0, 0)); await CopyAsync(false);
            check(memory.Value.Token is not null && memory.Value.Text is null && memory.Value.SpreadsheetXml is null && memory.Value.Html is null,
                "HTML boundary-date rejection publishes only the internal token, with no partial XML or text");
            _grid.Select(new(1, 0)); await PasteAsync("all");
            check(Document.Snapshot.GetSheet(sheet).GetCell(1, 0)!.Value == CellValue.FromNumber(60), "Internal copy remains usable when external export is rejected");
            var previous = memory.Value;
            try { await CopyAsync(false, externalFormulas: true); throw new Exception("Boundary date exported"); } catch (NotSupportedException) { }
            check(ReferenceEquals(previous, memory.Value), "Failed explicit external copy retains the previous clipboard object");
            _workspace.New(); ShowBook();
            var filtered = System.Xml.Linq.XElement.Parse(System.Text.Encoding.UTF8.GetString(SpreadsheetClipboardCodec.Write(source)).TrimEnd('\0'));
            System.Xml.Linq.XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
            filtered.Element(ss + "Worksheet")!.AddFirst(new System.Xml.Linq.XElement(ss + "Names",
                new System.Xml.Linq.XElement(ss + "NamedRange", new System.Xml.Linq.XAttribute(ss + "Name", "_FilterDatabase"),
                    new System.Xml.Linq.XAttribute(ss + "RefersTo", "=Source!R1C1:R1C3"), new System.Xml.Linq.XAttribute(ss + "Hidden", "1"))));
            memory.Value = new("wrong text", SpreadsheetXml: System.Text.Encoding.UTF8.GetBytes(filtered.ToString()));
            _grid.Select(new(2, 2)); await PasteAsync("all");
            check(Document.Snapshot.Sheets[0].GetCell(2, 2)!.Value == CellValue.FromText("00123") &&
                Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(2, 2)].Bold && _status.Text!.Contains("筛选"),
                "Ordinary paste accepts Excel filter headers, keeps typed values and styles, and reports omitted filter controls");
            await History(false);
            check(Document.Snapshot.Sheets[0].Cells.IsEmpty, "Filter-header paste is one undoable operation");
            filtered.Descendants(ss + "Cell").First().Add(new System.Xml.Linq.XElement(ss + "Comment", "note"));
            memory.Value = new("wrong text", SpreadsheetXml: System.Text.Encoding.UTF8.GetBytes(filtered.ToString()));
            _grid.Select(new(2, 2)); await PasteAsync("all");
            check(Document.Snapshot.Sheets[0].GetCell(2, 2)!.Value == CellValue.FromText("00123") &&
                Document.Snapshot.Sheets[0].GetCell(2, 3)!.Value == CellValue.FromNumber(.12) &&
                !Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(2, 2)].Bold && _status.Text!.Contains("自动改为只贴值"),
                "Filter-header source with unsupported attachments automatically uses typed values and reports the fallback");
            before = Document.Snapshot;
            try { await PasteAsync("formats"); throw new Exception("Format paste became value writes"); } catch (NotSupportedException) { }
            check(ReferenceEquals(before, Document.Snapshot), "Explicit format paste never silently becomes an automatic value paste");
            await RunDrawingClipboardChecksAsync(check);
        }
        finally { _clipboard = original; }
    }

    // Development-only end-to-end driver: all OS exchange uses the production adapter and real handlers.
    private async Task RunClipboardExchangeAsync(string directory)
    {
        // Drive the handlers as one UI operation, just like PerformAsync. Prevent the debounce timer
        // from racing the diagnostic's immediate copy/undo/read and process shutdown.
        _working = true;
        var checks = new List<string>(); Exception? error = null;
        void Check(bool value, string description) { if (!value) throw new InvalidOperationException(description); checks.Add(description); }
        try
        {
            using var request = JsonDocument.Parse(await File.ReadAllBytesAsync(Path.Combine(directory, "request.json")));
            var operation = request.RootElement.GetProperty("operation").GetString();
            var mode = request.RootElement.TryGetProperty("mode", out var m) ? m.GetString()! : "all";
            _workspace.New(); ShowBook(); _clipboard = new AvaloniaClipboard(this);
            await ExecuteAsync(new UpdateScriptModulesCommand(ScriptModuleSet.Single("")));
            _files = new SmokeFiles(); _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            if (operation == "object") await RunDrawingClipboardExchangeAsync(directory, mode, Check);
            else if (operation == "export")
            {
                var sheet = Document.Snapshot.ActiveSheetId;
                await ExecuteAsync(new SetCellsCommand([
                    new(new(sheet, 3, 2), CellContent.FromLiteral(CellValue.FromText("00123"))),
                    new(new(sheet, 3, 3), CellContent.FromLiteral(CellValue.FromNumber(.12))),
                    new(new(sheet, 3, 4), CellContent.FromLiteral(CellValue.FromBoolean(true))),
                    new(new(sheet, 3, 5), CellContent.FromFormula("=1/0")),
                    new(new(sheet, 4, 2), CellContent.FromLiteral(CellValue.FromText("中文\n行\t列"))),
                    new(new(sheet, 4, 3), CellContent.FromLiteral(CellValue.FromNumber(45000.5))),
                    new(new(sheet, 4, 4), CellContent.FromFormula("=D4*4")),
                    new(new(sheet, 4, 5), CellContent.FromLiteral(CellValue.FromNumber(45000.5))),
                    new(new(sheet, 5, 4), CellContent.FromLiteral(CellValue.FromNumber(.5))),
                    new(new(sheet, 5, 5), CellContent.FromLiteral(CellValue.FromNumber(12.123456789))),
                    new(new(sheet, 5, 2), CellContent.FromLiteral(CellValue.FromText("merged")))]));
                await ExecuteAsync(new SetCellStyleCommand(Enumerable.Range(3, 6).SelectMany(r => Enumerable.Range(2, 4).Select(c => new CellAddress(sheet, r, c))), new(FontName: "Arial", FontSize: 12)));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 3, 2)], new(Bold: true, FontColor: "FFFF0000", FillColor: "FFFFFF00", NumberFormatId: 0, NumberFormatCode: "@")));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 3, 3)], new(NumberFormatId: 0, NumberFormatCode: "0.00%", Border: true)));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 4, 2)], new(WrapText: true)));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 4, 3)], new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd hh:mm:ss")));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 4, 5)], new(NumberFormatId: 0, NumberFormatCode: "mmmm")));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 5, 4)], new(NumberFormatId: 0, NumberFormatCode: "hh:mm:ss")));
                await ExecuteAsync(new SetCellStyleCommand([new(sheet, 5, 5)], new(NumberFormatId: 0, NumberFormatCode: "[h]:mm:ss")));
                await ExecuteAsync(new MergeCellsCommand(sheet, new(5, 2, 6, 3)));
                await CalculateAsync(); _grid.SelectRange(new(3, 2, 8, 5), new(3, 2));
                await CopyAsync(false, externalFormulas: mode == "formulas");
                Check((await _clipboard.ReadAsync()).SpreadsheetXml is not null, "QuickExcel synthetic workbook publishes real OS XML");
            }
            else
            {
                var incoming = await _clipboard.ReadAsync();
                Check(incoming.SpreadsheetXml is not null || incoming.Html is not null, "Avalonia reads XML or HTML from one real OS data object");
                if (incoming.SpreadsheetXml is { } spreadsheet) await File.WriteAllBytesAsync(Path.Combine(directory, "incoming.xml"), spreadsheet);
                if (incoming.Html is { } html) await File.WriteAllBytesAsync(Path.Combine(directory, "incoming.cfhtml"), html);
                if (incoming.Biff8 is { } biff8) await File.WriteAllBytesAsync(Path.Combine(directory, "incoming.biff8"), biff8);
                var source = incoming.SpreadsheetXml is { } xml ? SpreadsheetClipboardCodec.Read(xml, incoming.Biff8) : HtmlClipboardCodec.Read(incoming.Html!);
                _grid.Select(new(3, 2)); var before = Document.Snapshot; await PasteAsync(mode);
                var pasted = Document.Snapshot; Check(pasted.Revision > before.Revision, "Real paste handler commits external data");
                await File.WriteAllTextAsync(Path.Combine(directory, "paste-status.txt"), _status.Text);
                if (request.RootElement.TryGetProperty("expectedPasteNotice", out var expectedNotice))
                    Check(_status.Text?.Contains(expectedNotice.GetString()!, StringComparison.Ordinal) == true, "Paste handler reports the requested filter/fallback notice");
                await History(false); Check(Document.Snapshot.Sheets[0].Cells.IsEmpty && Document.Snapshot.Styles.Length == before.Styles.Length, "Real handler undo restores all data and styles"); await History(true);
                Check(Document.Snapshot.Sheets[0].Cells.Count == pasted.Sheets[0].Cells.Count, "Real handler redo restores pasted region");
                await CalculateAsync();
                _grid.SelectRange(new(3, 2, 3 + source.Rows - 1, 2 + source.Columns - 1), new(3, 2));
                await CopyAsync(false, externalFormulas: mode == "formulas");
            }
            if (operation != "object")
            {
            var snapshot = Document.Snapshot;
            foreach (var ext in new[] { "xlsx", "xlsq" })
            {
                var file = Path.Combine(directory, "quickexcel." + ext); var storage = new TypedWorkbookStorage(); storage.Save(file, snapshot); var reopened = storage.Open(file);
                Check(snapshot.Sheets[0].Cells.All(p => reopened.Sheets[0].GetCell(p.Key.Row, p.Key.Column)?.Value == p.Value.Value), "QuickExcel " + ext + " save/reopen preserves every typed value");
            }
            var outgoing = await _clipboard.ReadAsync();
            await File.WriteAllBytesAsync(Path.Combine(directory, "outgoing.xml"), outgoing.SpreadsheetXml!);
            await File.WriteAllBytesAsync(Path.Combine(directory, "outgoing.cfhtml"), outgoing.Html!);
            await File.WriteAllTextAsync(Path.Combine(directory, "outgoing.txt"), outgoing.Text);
            Check(outgoing.SpreadsheetXml is not null && outgoing.Html is not null, "Published XML and HTML remain available after FlushAsync");
            await File.WriteAllTextAsync(Path.Combine(directory, "model.json"), JsonSerializer.Serialize(snapshot.Sheets[0].Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column).Select(p => new
            { row = p.Key.Row, column = p.Key.Column, kind = p.Value.Value.Kind.ToString(), value = CellInput.DisplayText(p.Value.Value), formula = p.Value.Content.Kind == CellContentKind.Formula ? p.Value.Content.FormulaText : null,
                style = snapshot.Styles[p.Value.Content.StyleId] }), new JsonSerializerOptions { WriteIndented = true }));
            }
        }
        catch (Exception ex) { error = ex; }
        _calculationTimer.Stop(); _working = false;
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = error is null, checks, error = error?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}
