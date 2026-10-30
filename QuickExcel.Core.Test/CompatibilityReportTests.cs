using QuickExcel.Desktop;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class CompatibilityReportTests
{
    [Fact]
    public void CurrentBlockedDependenciesAreLocatedWithoutCalculatingDuringInspection()
    {
        var snapshot = Model("=UNKNOWN_FN(1)", "=A1+1", "=SUM(1,2)"); using var session = new DocumentSession(snapshot);
        var result = new QuickExcel.Calculation.WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation());
        Assert.True(session.TryCommitCalculation(result)); var before = session.Snapshot;
        var report = TypedWorkbookStorage.InspectCompatibility(before);
        Assert.Contains(report.Diagnostics, d => d.Code == "formula.blocked-dependency" && d.Range == "A2" && d.Kind == CompatibilityKind.CannotCalculate);
        Assert.DoesNotContain(report.Diagnostics, d => d.Kind == CompatibilityKind.CannotCalculate && d.Range == "A3");
        Assert.Same(before, session.Snapshot); Assert.False(session.ScriptsEnabled);
    }
    [Fact]
    public void ImportRejectionsRetainOriginalExceptionTypeAndLocateFile()
    {
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "rejected-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "bad.xlsx"); File.WriteAllText(path, "not a ZIP");
        var error = Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
        var diagnostic = Assert.IsType<CompatibilityDiagnostic>(TypedWorkbookStorage.DescribeRejection(error));
        Assert.Equal(CompatibilityKind.Rejected, diagnostic.Kind); Assert.Equal("bad.xlsx", diagnostic.Part);
        Assert.Contains("【拒绝】", CompatibilityReportText.Rejection(diagnostic)); Assert.Equal("not a ZIP", File.ReadAllText(path));
    }
    private static WorkbookSnapshot Model(params string[] formulas) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new WorksheetSnapshot(SheetId.New(), "当前表", formulas.Select((f, i) => new KeyValuePair<CellPosition, CellSnapshot>(
            new(i, 0), new(CellContent.FromFormula(f), new(CellValue.FromNumber(7), CalculationState.Stale, 0, 0, 0)))))],
        "throw new Error('inspection must not execute scripts');");

    [Fact]
    public void CapabilityInspectionKeepsCachesAndSeparatesUnknownFunctionsFromSaveLosses()
    {
        var snapshot = Model("=_xlfn.XLOOKUP(1,B1:B2,C1:C2)", "=SUM(1,2)", "=MY_CUSTOM(7)", "=LEN(\"XLOOKUP(1)\")");
        var report = TypedWorkbookStorage.InspectCompatibility(snapshot);
        var warnings = report.Diagnostics.Where(d => d.Kind == CompatibilityKind.CannotCalculate).ToArray();
        Assert.Equal(new[] { "A1", "A3" }, warnings.Select(d => d.Range));
        Assert.All(warnings, d => { Assert.Equal("当前表", d.SheetName); Assert.Contains("缓存", d.Message); });
        Assert.Contains(warnings, d => d.Message.Contains("MY_CUSTOM") && d.Message.Contains("本报告不执行"));
        Assert.Empty(TypedWorkbookStorage.SaveIssues(snapshot));
        Assert.All(snapshot.Sheets[0].Cells.Values, c => {
            Assert.Equal(CalculationState.Stale, c.Calculation!.State);
            Assert.Equal(7, c.Value.NumberValue);
        });
        Assert.Contains("4 个公式", Assert.Single(report.Diagnostics, d => d.Code == "formula.cache-summary").Message);
    }

    [Fact]
    public void NamesAreExpandedAndCurrentEditsRemoveOutdatedFormulaWarnings()
    {
        var snapshot = Model("=PendingResult");
        snapshot = snapshot.WithDefinedNames([new DefinedName("PendingResult", "=_xlfn.XLOOKUP(1,'当前表'!$B$1:$B$2,'当前表'!$C$1:$C$2)")]);
        Assert.Contains(TypedWorkbookStorage.InspectCompatibility(snapshot).Diagnostics, d => d.Kind == CompatibilityKind.CannotCalculate);
        using var session = new DocumentSession(snapshot);
        session.Execute(new SetCellsCommand([new(new(snapshot.Sheets[0].Id, 0, 0), CellContent.FromFormula("=1+2"))]));
        Assert.DoesNotContain(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics, d => d.Kind == CompatibilityKind.CannotCalculate);
    }

    [Theory]
    [InlineData("=Table1[Amount]")]
    [InlineData("=SUM(")]
    public void UnsupportedSyntaxHasCurrentCellLocationAndDoesNotGrantFreshness(string formula)
    {
        var snapshot = Model(formula);
        var diagnostic = Assert.Single(TypedWorkbookStorage.InspectCompatibility(snapshot).Diagnostics, d => d.Code == "formula.unsupported-syntax");
        Assert.Equal("A1", diagnostic.Range);
        Assert.Equal(CompatibilityKind.CannotCalculate, diagnostic.Kind);
        Assert.Equal(CalculationState.Stale, snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
    }

    [Fact]
    public void FormulaDiagnosticsAreBoundedAndExplicitAboutUninspectedCells()
    {
        var snapshot = Model(Enumerable.Repeat("=UNKNOWN_FN(1)", 1030).ToArray());
        var report = TypedWorkbookStorage.InspectCompatibility(snapshot);
        Assert.Equal(1024, report.Diagnostics.Count(d => d.Code == "formula.unsupported-function"));
        Assert.Contains(report.Diagnostics, d => d.Code == "report.formulas-truncated");
        Assert.Contains(report.Diagnostics, d => d.Code == "formula.cache-summary" && d.Message.Contains("1030 个公式"));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => TypedWorkbookStorage.InspectCompatibility(snapshot, cancelled.Token));
    }

    [Fact]
    public void DialogTextDistinguishesAllFiveKindsAndLocatesEachFeature()
    {
        var report = new WorkbookCompatibilityReport(true, Enum.GetValues<CompatibilityKind>().Select(kind =>
            new CompatibilityDiagnostic("test", kind, "worksheet/pageSetup", "诊断说明", "xl/worksheets/sheet1.xml", "模板", "A2")).ToArray());
        var text = CompatibilityReportText.Format(report);
        foreach (var kind in Enum.GetValues<CompatibilityKind>()) Assert.Contains("【" + CompatibilityReportText.Label(kind) + "】", text);
        Assert.Contains("模板 · A2 · xl/worksheets/sheet1.xml · worksheet/pageSetup", text);
        Assert.Contains("仅保留不代表可编辑写回", text);
        Assert.Contains("导入时", text);
        Assert.Contains("潜在损失会阻止普通保存", CompatibilityReportText.Format(report with { OriginalPackageUnchanged = false }));
        var supported = CompatibilityReportText.Format(TypedWorkbookStorage.InspectCompatibility(Model("=1+2")));
        Assert.Contains("未列出格式损失阻止项", supported);
        Assert.DoesNotContain("须明确另存受支持副本", supported);
    }
}
