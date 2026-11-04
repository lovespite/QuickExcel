using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunDrawingClipboardChecksAsync(Action<bool, string> check)
    {
        _workspace.New(); ShowBook(); var memory = new MemoryClipboard(); _clipboard = memory;
        using var bitmap = new SkiaSharp.SKBitmap(8, 4); bitmap.Erase(SkiaSharp.SKColors.CornflowerBlue);
        using var encoded = bitmap.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100); var png = encoded.ToArray();
        memory.Value = new(null, Png: png); _grid.Select(new(2, 3)); await PasteAsync("all");
        check(Document.Snapshot.Sheets[0].Images is [{ Anchor.Row: 2, Anchor.Column: 3 }] && Document.Snapshot.Sheets[0].Cells.IsEmpty, "Image paste anchors at the selected cell without overwriting it");
        await PasteAsync("image"); check(Document.Snapshot.Sheets[0].Images.Length == 2, "Repeated image paste works while the preceding image is selected");
        await History(false); check(Document.Snapshot.Sheets[0].Images.Length == 1, "One image paste is one undo"); await History(true);
        _grid.Select(new(3, 4)); var before = Document.Snapshot;
        memory.Value = new(null, Png: png, OfficeDrawing: [1, 2, 3]);
        await PasteAsync("chart");
        check(Document.Snapshot.Sheets[0].Images.Length == 3 && _status.Text!.Contains("自动降级为图片"), "Unsupported editable chart automatically uses its picture representation");
        await History(false); check(Document.Snapshot.Sheets[0].Images.Length == 2 && Document.Snapshot.Sheets.Length == 1, "Chart-to-picture fallback has one undo and creates no stray data sheet");
        await PasteAsync("all"); check(Document.Snapshot.Sheets[0].Images.Length == 3 && _status.Text!.Contains("自动降级为图片"), "Default paste also falls back from unsupported objects to PNG");
        await PasteAsync("image"); check(Document.Snapshot.Sheets[0].Images.Length == 4, "Explicit picture selection can use PNG independently of unsupported chart data");
        await CopyAsync(false); check(memory.Value.Png is not null && memory.Value.Token is null && memory.Value.SpreadsheetXml is null, "Selected image copies as an image without a stale cell token");
        _grid.Select(new(4, 5)); before = Document.Snapshot; memory.Value = new(null, Png: [1, 2, 3]);
        await PasteAsync("image");
        check(ReferenceEquals(before, Document.Snapshot) && _status.Text!.Contains("保持不变"), "Unreadable image with no alternative leaves history unchanged without a refusal dialog");
        memory.Value = new("object description", Png: [1, 2, 3], OfficeDrawing: [1, 2, 3]);
        await PasteAsync("all");
        check(Document.Snapshot.Sheets[0].GetCell(4, 5)!.Value == CellValue.FromText("object description") && _status.Text!.Contains("降级"), "Unreadable object and image fall back to available text");
        before = Document.Snapshot;
        _clipboard = new DelayedClipboard(new(null, Png: png), () => _grid.Select(new(8, 8)));
        try { await PasteAsync("image"); throw new Exception("Stale image target accepted"); } catch (InvalidOperationException) { }
        check(before.Revision == Document.Snapshot.Revision && before.Sheets[0].Images.SequenceEqual(Document.Snapshot.Sheets[0].Images), "Image paste rejects a changed destination during clipboard read");
        _clipboard = memory; memory.Value = new(null, Png: png);
        await ExecuteAsync(new SetProtectionCommand(Document.Snapshot.ActiveSheetId, new() { Sheet = true, Objects = true })); before = Document.Snapshot;
        try { await PasteAsync("image"); throw new Exception("Protected image target accepted"); } catch (InvalidOperationException) { }
        check(ReferenceEquals(before, Document.Snapshot), "Object protection rejects the complete image insertion");
    }
    private async Task RunDrawingClipboardExchangeAsync(string directory, string mode, Action<bool, string> check)
    {
        var incoming = await _clipboard.ReadAsync();
        if (incoming.OfficeDrawing is { } drawing) await File.WriteAllBytesAsync(Path.Combine(directory, "incoming.gvml"), drawing);
        if (incoming.Png is { } png) await File.WriteAllBytesAsync(Path.Combine(directory, "incoming.png"), png);
        _grid.Select(new(3, 2)); var before = Document.Snapshot; await PasteAsync(mode);
        var pasted = Document.Snapshot; var sheet = pasted.Sheets[0];
        check(pasted.Revision == before.Revision + 1, "Object paste is one transaction");
        check(sheet.Cells.IsEmpty, "Object paste does not overwrite destination cells");
        var isChart = sheet.Data.Charts.Length == 1;
        check(isChart ? sheet.Images.IsEmpty && pasted.Sheets.Length == 2 : sheet.Images.Length == 1 && pasted.Sheets.Length == 1, "Only the selected object representation is inserted");
        await File.WriteAllTextAsync(Path.Combine(directory, "paste-status.txt"), _status.Text);
        await History(false); check(Document.Snapshot.Sheets.Length == 1 && Document.Snapshot.Sheets[0].Images.IsEmpty && Document.Snapshot.Sheets[0].Data.Charts.IsEmpty, "Undo removes the object and accompanying data sheet together");
        await History(true); check(Document.Snapshot.Sheets.Length == pasted.Sheets.Length, "Redo restores object data");
        _grid.SelectDrawingObject(new(isChart ? DrawingObjectKind.Chart : DrawingObjectKind.Image, 0));
        await CopyAsync(false); var outgoing = await _clipboard.ReadAsync();
        check(outgoing.Png is not null, "Copied object publishes a real OS PNG");
        await File.WriteAllBytesAsync(Path.Combine(directory, "outgoing.png"), outgoing.Png!);
        if (isChart)
        {
            check(outgoing.ChartPackage is not null, "Copied chart also publishes private editable drawing data");
            await File.WriteAllBytesAsync(Path.Combine(directory, "outgoing-chart.xlsx"), outgoing.ChartPackage!);
            if (outgoing.OfficeDrawing is { } native) await File.WriteAllBytesAsync(Path.Combine(directory, "outgoing.gvml"), native);
            await File.WriteAllTextAsync(Path.Combine(directory, "chart-data.json"), JsonSerializer.Serialize(ChartData.Read(pasted, sheet.Data.Charts[0]), new JsonSerializerOptions { WriteIndented = true }));
            var decoded = QuickExcel.Clipboard.OfficeDrawingClipboardCodec.ReadPrivateChart(outgoing.ChartPackage!);
            check(JsonSerializer.Serialize(ChartData.Read(decoded.Book, decoded.Chart)) == JsonSerializer.Serialize(ChartData.Read(pasted, sheet.Data.Charts[0])), "Private editable clipboard roundtrip preserves all chart values and labels");
        }
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "quickexcel." + extension); var storage = new TypedWorkbookStorage(); storage.Save(path, Document.Snapshot); var loaded = storage.Open(path);
            check(loaded.FileState?.FormulaReadOnly != true && loaded.Sheets[0].Images.Length == sheet.Images.Length && loaded.Sheets[0].Data.Charts.Length == sheet.Data.Charts.Length, extension + " save and reopen preserve editable objects without preservation-only diagnostics");
            if (isChart) check(JsonSerializer.Serialize(ChartData.Read(loaded, loaded.Sheets[0].Data.Charts[0])) == JsonSerializer.Serialize(ChartData.Read(pasted, sheet.Data.Charts[0])), extension + " preserves every chart series value and label");
        }
    }
}
