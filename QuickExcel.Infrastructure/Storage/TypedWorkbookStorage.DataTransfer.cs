using System.Collections.Immutable;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    /// <summary>Writes an explicit data-only XLSX projection without changing its importing workbook's API version.</summary>
    public FileVersion ExportDataValuesChecked(string path, WorkbookSnapshot values, FileVersion expected,
        CancellationToken cancellation = default)
    {
        // Validate the values-only boundary before using XLSX's canonical script-free representation.
        // This discards no executable source and does not migrate the importing workbook's modules.
        _ = new ImportWorksheetDataCommand(values);
        if (!Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Data workbook exports require .xlsx.", nameof(path));
        var output = values.WithScriptModules(ScriptModuleSet.Single(""));
        output = output.WithSheets(output.Sheets.Select(sheet => sheet.WithCells(sheet.Cells.ToImmutableDictionary(pair =>
            pair.Key, pair => NormalizeHostError(pair.Value)))));
        return SaveCore(path, output, false, cancellation, expected).Version;
    }

    /// <summary>Checked copy export that can be cancelled before the atomic file replacement.</summary>
    public FileVersion SaveChecked(string path, WorkbookSnapshot snapshot, bool discardUnsupported,
        FileVersion expected, CancellationToken cancellation) =>
        SaveCore(path, snapshot, discardUnsupported, cancellation, expected).Version;
}
