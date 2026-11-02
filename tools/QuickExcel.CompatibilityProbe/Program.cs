using System.Text;
using System.Text.Json;
using QuickExcel.CompatibilityProbe;

try
{
    if (args is ["e4-report", var reportSource, var reportDirectory]) return E4Probe.Run(reportSource, reportDirectory);
    if (args is ["e4-capacity", var capacityDirectory]) return E4Probe.Capacity(capacityDirectory);
    if (args is ["outline-subtotal", var outlineOracle, var outlineDirectory]) return OutlineSubtotalProbe.Run(outlineOracle, outlineDirectory);
    if (args is ["data-cleaning", var cleaningOracle, var cleaningDirectory]) return DataCleaningProbe.Run(cleaningOracle, cleaningDirectory);
    if (args is ["verify-data-cleaning", var cleaningOutput, var cleaningVerdict]) return DataCleaningProbe.Verify(cleaningOutput, cleaningVerdict);
    if (args is ["conditional-formatting", var conditionalFormatOracle, var conditionalFormatDirectory]) return ConditionalFormattingProbe.Run(conditionalFormatOracle, conditionalFormatDirectory);
    if (args is ["verify-conditional-formatting", var cfExpected, var cfActual, var cfSource, var cfVerdict]) return ConditionalFormattingProbe.Verify(cfExpected, cfActual, cfSource, cfVerdict);
    if (args is ["verify-conditional-model", var cfOutput, var cfWorkbook, var cfModelVerdict]) return ConditionalFormattingProbe.VerifyModel(cfOutput, cfWorkbook, cfModelVerdict);
    if (args is ["data-validation", var validationOracle, var validationDirectory]) return DataValidationProbe.Run(validationOracle, validationDirectory);
    if (args is ["verify-data-validation", var validationExpected, var validationActual, var validationVerdict]) return DataValidationProbe.Verify(validationExpected, validationActual, validationVerdict);
    if (args is ["structured-tables", var structuredOracle, var structuredDirectory]) return StructuredTablesProbe.Run(structuredOracle, structuredDirectory);
    if (args is ["table-basics", var tableOracle, var tableDirectory]) return TableBasicsProbe.Run(tableOracle, tableDirectory);
    if (args is ["table-edits", var tableSource, var tableEdits, var tableEditDirectory]) return TableBasicsProbe.Run(tableSource, tableEditDirectory, tableEdits);
    if (args is ["filter-persistence", var filterOracle, var filterDirectory]) return FilterPersistenceProbe.Run(filterOracle, filterDirectory);
    if (args is ["format-efficiency", var formatSpec, var formatOracle, var formatDirectory]) return FormatEfficiencyProbe.Run(formatSpec, formatOracle, formatDirectory);
    if (args is ["scalar-functions", var scalarSpec, var scalarOracle, var scalarDirectory]) return ScalarFunctionsProbe.Run(scalarSpec, scalarOracle, scalarDirectory);
    if (args is ["verify-scalar-functions", var scalarExpected, var scalarInitial, var scalarEdited, var scalarModel, var scalarVerdict]) return ScalarFunctionsProbe.Verify(scalarExpected, scalarInitial, scalarEdited, scalarModel, scalarVerdict);
    if (args is ["conditional-functions", var conditionalOracle, var conditionalDirectory]) return ConditionalFunctionsProbe.Run(conditionalOracle, conditionalDirectory);
    if (args is ["verify-conditional-functions", var conditionalExpected, var conditionalInitial, var conditionalEdited, var conditionalModel, var conditionalVerdict]) return ConditionalFunctionsProbe.Verify(conditionalExpected, conditionalInitial, conditionalEdited, conditionalModel, conditionalVerdict);
    if (args is ["sheet-names", var sheetNamesOracle, var sheetNamesDirectory]) return SheetNamesProbe.Run(sheetNamesOracle, sheetNamesDirectory);
    if (args is ["verify-sheet-names", var sheetNamesExpected, var sheetNamesActual, var sheetNamesModel, var sheetNamesVerdict]) return SheetNamesProbe.Verify(sheetNamesExpected, sheetNamesActual, sheetNamesModel, sheetNamesVerdict);
    if (args is ["fill", var fillOracle, var fillDirectory]) return FillProbe.Run(fillOracle, fillDirectory);
    if (args is ["verify-fill", var fillExpected, var fillActual, var fillModel, var fillVerdict]) return FillProbe.Verify(fillExpected, fillActual, fillModel, fillVerdict);
    if (args is ["clipboard", var clipboardOracle, var clipboardDirectory]) return ClipboardProbe.Run(clipboardOracle, clipboardDirectory);
    if (args is ["verify-clipboard", var clipboardExpected, var clipboardActual, var clipboardModel, var clipboardVerdict]) return ClipboardProbe.Verify(clipboardExpected, clipboardActual, clipboardModel, clipboardVerdict);
    if (args is ["business-baseline", var businessReference, var businessDirectory]) return BusinessBaselineProbe.Run(businessReference, businessDirectory);
    if (args is ["edit-matrix", var matrixSources, var matrixOracle, var matrixDirectory]) return EditMatrixProbe.Run(matrixSources, matrixOracle, matrixDirectory);
    if (args is ["verify-edit-matrix", var expectedEdits, var observedEdits, var modelEdits, var editVerdict]) return EditMatrixVerification.Run(expectedEdits, observedEdits, modelEdits, editVerdict);
    if (args is ["number-formats", var reference, var numberDirectory]) return NumberFormatProbe.Run(reference, numberDirectory);
    if (args is ["styles", var styleSource, var styleDirectory]) return StyleProbe.Run(styleSource, styleDirectory);
    if (args is ["metadata", var metadataSource, var metadataDirectory]) return MetadataProbe.Run(metadataSource, metadataDirectory);
    if (args is ["views", var viewSource, var viewDirectory]) return ViewProbe.Run(viewSource, viewDirectory);
    if (args is ["printing", var printSource, var printDirectory]) return PrintingProbe.Run(printSource, printDirectory);
    if (args is ["formulas", var formulaSource, var formulaDirectory]) return FormulaProbe.Run(formulaSource, formulaDirectory);
    if (args is ["cell-text", var textSource, var textDirectory]) return CellTextProbe.Run(textSource, textDirectory);
    if (args is ["baseline", var directory])
    {
        directory = Path.GetFullPath(directory);
        if (Directory.Exists(directory) || File.Exists(directory)) throw new IOException("Output directory must be new; previous evidence is never overwritten.");
        Directory.CreateDirectory(directory);
        var reports = FixtureCatalog.Create().Select(f => ProbeRunner.Run(f, Path.Combine(directory, f.Id))).ToArray();
        var summary = new StringBuilder("# XLSX compatibility baseline\n\nSynthetic fixtures; real Excel comparison pending. Differences include intended edits and representation changes; no automatic fidelity pass. XLSX outputs are extracted from saved XLSQ, not the desktop XLSX save flow.\n\n| Fixture | Input validation errors | Outputs | Output validation errors | Normal save blocked | Rejected paths |\n| --- | ---: | ---: | ---: | ---: | ---: |\n");
        foreach (var report in reports)
        {
            summary.AppendLine($"| {report.Id} | {report.Source.Issues.Count} | {report.Paths.Count(p => p.Output is not null)} | {report.Paths.Sum(p => p.Output?.Issues.Count ?? 0)} | {report.Paths.Count(p => p.Status.StartsWith("normal-save-blocked", StringComparison.Ordinal))} | {report.Paths.Count(p => p.Output is null)} |");
            Console.WriteLine($"{report.Id}: {report.Paths.Count} paths, source unchanged={report.SourceUnchanged}, Excel=pending");
        }
        ProbeRunner.WriteNew(Path.Combine(directory, "summary.md"), Encoding.UTF8.GetBytes(summary.ToString()));
        // A diagnostic run may expose existing output defects. Only broken input fixtures or source mutation fail baseline generation.
        return reports.All(r => r.SourceUnchanged && r.Source.Issues.Count == 0) ? 0 : 2;
    }
    if (args is ["compare", var before, var after, var output])
    {
        byte[] ReadBounded(string path) => new FileInfo(path).Length <= 64 * 1024 * 1024 ? File.ReadAllBytes(path) : throw new InvalidDataException("Input exceeds 64 MiB.");
        var left = OpenXmlAudit.Inspect(ReadBounded(before));
        var right = OpenXmlAudit.Inspect(ReadBounded(after));
        var differences = OpenXmlAudit.Compare(left, right);
        ProbeRunner.WriteNew(output, JsonSerializer.SerializeToUtf8Bytes(new { Before = left, After = right, Differences = differences, ExcelVerification = "not-performed" }, ProbeRunner.JsonOptions));
        Console.WriteLine($"{differences.Count} differences; {left.Issues.Count + right.Issues.Count} validation issues.");
        return differences.Count == 0 && left.Issues.Count == 0 && right.Issues.Count == 0 ? 0 : 2;
    }
    Console.Error.WriteLine("Usage: baseline <new-output-directory> | compare <before.xlsx> <after.xlsx> <new-report.json> | business-baseline <excel-reference.json> <new-directory> | edit-matrix <source-directory> <excel-oracle.json> <new-directory> | verify-edit-matrix <excel-oracle.json> <excel-observed.json> <model-report.json> <new-verdict.json>");
    return 1;
}
catch (Exception ex) when (ex is IOException or InvalidOperationException or ArgumentException or System.Xml.XmlException)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}
