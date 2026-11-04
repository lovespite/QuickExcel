using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
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
