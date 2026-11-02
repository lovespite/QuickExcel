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
        var snapshot = Model("=_xlfn.FUTURELOOKUP(1,B1:B2,C1:C2)", "=SUM(1,2)", "=MY_CUSTOM(7)", "=LEN(\"XLOOKUP(1)\")");
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
        snapshot = snapshot.WithDefinedNames([new DefinedName("PendingResult", "=_xlfn.FUTURELOOKUP(1,'当前表'!$B$1:$B$2,'当前表'!$C$1:$C$2)")]);
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

    [Fact]
    public void SavePromptPromotesPreservedContentAndExplainsPerItemImpact()
    {
        var report = new WorkbookCompatibilityReport(false, new[]
        {
            new CompatibilityDiagnostic("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "part/xl/charts/chart1.xml",
                "未建模部件：xl/charts/chart1.xml", "xl/charts/chart1.xml", "模板", "A1"),
            new CompatibilityDiagnostic("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "font/outline · cellXfs[0]",
                "未建模属性 font/outline", "xl/styles.xml", "模板", "B2")
        });
        var view = CompatibilityReportPresentation.Create(report, saveMode: true);
        Assert.Contains("2 处内容无法在编辑后保留", view.Headline);
        Assert.Contains("原文件不会被修改", view.Summary);
        var group = Assert.Single(view.Groups);
        Assert.Equal(CompatibilityKind.PotentialLoss, group.Kind);
        Assert.True(group.ExpandedByDefault);
        Assert.All(group.Items, item => Assert.Contains("不会写入受支持副本", item.Impact));
        Assert.Contains(group.Items, item => item.Impact.Contains("未建模部件"));
        Assert.Contains(group.Items, item => item.Impact.Contains("字体"));
        Assert.False(CompatibilityReportPresentation.Create(report, saveMode: false).HasLosses);
    }

    [Fact]
    public void ReportMergesRepeatedDiagnosticsAndKeepsEveryLocation()
    {
        var diagnostics = Enumerable.Range(0, 2).Select(index => new CompatibilityDiagnostic(
            "xlsx.affected-location", CompatibilityKind.PreservedOnly, "font/outline · cellXfs[0]",
            "未建模属性 font/outline", "xl/styles.xml", "诊断", index == 0 ? "A1" : "B2")).ToArray();
        var item = Assert.Single(Assert.Single(CompatibilityReportPresentation
            .Create(new WorkbookCompatibilityReport(true, diagnostics), saveMode: false).Groups).Items);
        Assert.Equal(2, item.Count);
        Assert.Contains("诊断 · A1 · xl/styles.xml", item.Location);
        Assert.Contains("诊断 · B2 · xl/styles.xml", item.Location);
        Assert.Contains("不会按原样写入", item.Impact);
    }

    [Fact]
    public void ImpactTextSeparatesCalculationRejectionAndEditableSubsets()
    {
        string Impact(CompatibilityKind kind, string feature) => CompatibilityReportPresentation.Impact(
            new CompatibilityDiagnostic("test", kind, feature, "诊断说明", "xl/worksheets/sheet1.xml", "模板", "A2"), saveMode: false);
        Assert.Contains("已有缓存", Impact(CompatibilityKind.CannotCalculate, "formula/function"));
        Assert.Contains("只读", Impact(CompatibilityKind.CannotCalculate, "formula/preserved-readonly"));
        Assert.Contains("阻止打开", Impact(CompatibilityKind.Rejected, "worksheet/pageSetup"));
        Assert.Contains("可正常编辑", Impact(CompatibilityKind.Editable, "worksheet/data"));
        var notice = "XLSQ v1 → v2：首次保存为新副本，原文件保留。";
        var view = CompatibilityReportPresentation.Create(new WorkbookCompatibilityReport(true, []), saveMode: false, [notice]);
        Assert.Equal(new[] { notice }, view.Notices);
        Assert.Equal("工作簿与导入时一致，可原样保存", view.Headline);
        Assert.Empty(view.Groups);
        Assert.Contains("【其他提示】\n" + notice, CompatibilityReportPresentation.PlainText(
            new WorkbookCompatibilityReport(true, []), view.Notices));
    }
}
