using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public sealed record ProbePath(string Mode, string Operation, string Status, string[] CompatibilityIssues,
    string? Error, string? OutputFile, PackageAudit? Output, IReadOnlyList<PackageDifference> Differences,
    IReadOnlyList<CompatibilityDiagnostic> Diagnostics);
public sealed record FixtureReport(string Id, string Scenario, string[] RoadmapItems, string Origin,
    string ExcelVerification, string? ExcelVersion, string? ExcelBuild, string? ExcelLanguage, string? CalculationMode,
    string DateSystem, PackageAudit Source, bool SourceUnchanged, IReadOnlyList<ProbePath> Paths);

public static class ProbeRunner
{
    public static readonly string[] Modes = ["unchanged", "value", "style", "insert-row", "delete-row", "insert-column", "delete-column", "rename"];
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new JsonStringEnumConverter() }
    };

    public static FixtureReport Run(CompatibilityFixture fixture, string directory)
    {
        Directory.CreateDirectory(directory);
        var sourcePath = Path.Combine(directory, "source.xlsx");
        WriteNew(sourcePath, fixture.Bytes);
        var source = OpenXmlAudit.Inspect(fixture.Bytes);
        var paths = new List<ProbePath>();
        string dateSystem = "unknown";
        foreach (var mode in Modes)
        {
            var issues = Array.Empty<string>();
            IReadOnlyList<CompatibilityDiagnostic> diagnostics = [];
            var operation = Describe(mode);
            var stage = "open";
            try
            {
                var storage = new TypedWorkbookStorage();
                var snapshot = storage.Open(sourcePath);
                dateSystem = snapshot.DateSystem.ToString();
                issues = snapshot.FileState!.CompatibilityIssues.ToArray();
                using var session = new DocumentSession(snapshot);
                diagnostics = TypedWorkbookStorage.InspectCompatibility(snapshot).Diagnostics;
                stage = "edit";
                Apply(session, mode);
                diagnostics = TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics;
                stage = "save";
                var packagePath = Path.Combine(directory, mode + ".xlsq");
                string status;
                string? error = null;
                try
                {
                    storage.Save(packagePath, session.Snapshot);
                    status = mode == "unchanged" ? "preserved-original" : "saved-modeled-subset";
                }
                catch (WorkbookCompatibilityException ex)
                {
                    // Diagnostic copies only, never an implicit user save. Preserve the normal-save rejection in the report.
                    error = ex.Message;
                    packagePath = Path.Combine(directory, mode + ".diagnostic-lossy-copy.xlsq");
                    storage.Save(packagePath, session.Snapshot, discardUnsupported: true);
                    status = "normal-save-blocked;diagnostic-lossy-copy";
                }
                var outputBytes = OpenXmlAudit.ReadParts(File.ReadAllBytes(packagePath))["workbook.xlsx"];
                var outputPath = Path.ChangeExtension(packagePath, ".xlsx");
                WriteNew(outputPath, outputBytes);
                var audit = OpenXmlAudit.Inspect(outputBytes);
                paths.Add(new(mode, operation, status, issues, error, Path.GetFileName(outputPath), audit, OpenXmlAudit.Compare(source, audit), diagnostics));
            }
            catch (Exception ex) when (ex is InvalidDataException or InvalidOperationException or ArgumentException or NotSupportedException)
            {
                if (ex is WorkbookFeatureRejectedException rejected) diagnostics = [rejected.Diagnostic];
                paths.Add(new(mode, operation, stage + "-rejected", issues, ex.Message, null, null, [], diagnostics));
            }
        }
        var report = new FixtureReport(fixture.Id, fixture.Scenario, fixture.RoadmapItems, "synthetic-inline-spreadsheetml",
            "pending-real-excel", null, null, null, null, dateSystem, source, OpenXmlAudit.Hash(File.ReadAllBytes(sourcePath)) == source.Sha256, paths);
        WriteNew(Path.Combine(directory, "report.json"), JsonSerializer.SerializeToUtf8Bytes(report, JsonOptions));
        return report;
    }

    public static string Describe(string mode) => mode switch
    {
        "unchanged" => "无编辑，保存 XLSQ 后提取 workbook.xlsx；要求原 XLSX 字节不变",
        "value" => "首表 A2 改为数值 99，保留该格样式",
        "style" => "首表 A2 切换粗体，其他样式分量保持",
        "insert-row" => "首表第 2 行前插入 1 行",
        "delete-row" => "首表删除第 2 行",
        "insert-column" => "首表 B 列前插入 1 列",
        "delete-column" => "首表删除 B 列",
        "rename" => "首表改名为 改名后",
        _ => throw new ArgumentException("Unknown probe mode.")
    };

    internal static void Apply(DocumentSession session, string mode)
    {
        var sheet = session.Snapshot.Sheets[0];
        var address = new CellAddress(sheet.Id, 1, 0);
        DocumentCommand? command = mode switch
        {
            "unchanged" => null,
            "value" => new SetCellsCommand([new(address, CellContent.FromLiteral(CellValue.FromNumber(99), sheet.GetStyleId(1, 0)))]),
            "style" => new SetCellStyleCommand([address], new(Bold: !session.Snapshot.Styles[sheet.GetStyleId(1, 0)].Bold)),
            "insert-row" => new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1),
            "delete-row" => new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1, true),
            "insert-column" => new ShiftAxisCommand(sheet.Id, SheetAxis.Column, 1, 1),
            "delete-column" => new ShiftAxisCommand(sheet.Id, SheetAxis.Column, 1, 1, true),
            "rename" => new RenameSheetCommand(sheet.Id, "改名后"),
            _ => throw new ArgumentException("Unknown probe mode.")
        };
        if (command is not null) session.Execute(command);
    }

    public static void WriteNew(string path, byte[] bytes)
    {
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
        file.Write(bytes);
    }
}
