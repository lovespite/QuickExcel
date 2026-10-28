using System.IO;
using System.IO.Compression;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunStorageSmokeAsync(string folder)
    {
        var path = Path.Combine(folder, "typed-ui.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), path, "typed-ui.xlsq");
        ApplyValues([(0, 0, "123"), (0, 1, "00123"), (0, 2, "TRUE"), (0, 3, "'"), (0, 4, "=A1*2")]);
        if (!await SaveAsync(false)) throw new Exception("Typed UI save failed.");
        using var reopened = WorkbookDocument.Open(path, _services);
        var kinds = new[] { CellValueKind.Number, CellValueKind.Text, CellValueKind.Boolean, CellValueKind.Text };
        for (var i = 0; i < kinds.Length; i++)
            if (reopened.Snapshot.Sheets[0].GetCell(0, i)?.Value.Kind != kinds[i]) throw new Exception($"Cell {i} lost its type through the UI save path.");
        if (reopened.Snapshot.DocumentId != _document.Snapshot.DocumentId || reopened.Snapshot.Sheets[0].Id != _document.Snapshot.Sheets[0].Id || reopened.ScriptsEnabled)
            throw new Exception("Typed UI save lost identity or restored script authorization.");
        using (var zip = ZipFile.OpenRead(path))
        using (var manifest = JsonDocument.Parse(zip.GetEntry("manifest.json")!.Open()))
            if (manifest.RootElement.GetProperty("formatVersion").GetInt32() != 2) throw new Exception("UI still writes v1.");
        var v1 = Path.Combine(folder, "migration-source.xlsq");
        new LegacyWorkbookStorage().Save(v1, _document.Snapshot);
        var before = File.ReadAllBytes(v1);
        InstallDocument(WorkbookDocument.Open(v1, _services), v1, "migration-source.xlsq", fileVersion: FileVersions.Read(v1));
        if (!_document.IsDirty || !LogText.Text.Contains("另存") || _document.ScriptsEnabled)
            throw new Exception("Migration UI does not identify the unsaved v1 upgrade.");
        _filePath = Path.Combine(folder, "migration-upgraded.xlsq");
        if (!await SaveAsync(false) || _document.IsDirty || !before.SequenceEqual(File.ReadAllBytes(v1)))
            throw new Exception("Migration save changed its source or saved revision.");
        File.WriteAllText(Path.Combine(folder, "storage-smoke-result.txt"), "PASS: real desktop typed save (number/text/boolean/empty text), XLSQ v2 manifest, stable IDs, scripts disabled on open, v1 migration dirty/report, distinct-path upgrade preserves original.");
    }
}
