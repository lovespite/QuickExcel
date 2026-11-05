using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private enum CompatibilitySaveDecision { Cancel, SaveCopy, Overwrite }

    private void InstallFileCommands()
    {
        RibbonGroup(UiText.Get("审阅"), UiText.Get("文件兼容性"), RibbonButton(UiText.Get("兼容性报告"), CompatibilityAsync, false), RibbonButton(UiText.Get("导出值快照"), () => ExportAsync(false), false), RibbonButton(UiText.Get("导出可移植公式"), () => ExportAsync(true), false));
    }
    private async Task RecoveryAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var file = await _files.OpenAsync(new(UiText.Get("打开恢复副本"), [new(UiText.Get("QuickExcel 工作簿副本"), ["*.xlsq"])]));
        if (file is null || !await ConfirmDiscardAsync()) return;
        await _workspace.OpenAsync(file); _workspace.DetachFile(); ShowBook(); _status.Text = UiText.Get("已打开恢复副本，请另存。脚本未授权。");
    }
    private async Task CompatibilityAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var report = await _workspace.InspectAsync(snapshot => TypedWorkbookStorage.InspectCompatibility(snapshot));
        await ShowCompatibilityReportAsync(report, saveMode: false);
    }
    private Task<CompatibilitySaveDecision> ShowCompatibilityReportAsync(WorkbookCompatibilityReport report, bool saveMode, bool canOverwrite = false)
    {
        var view = CompatibilityReportPresentation.Create(report, saveMode, Document.Snapshot.FileState?.RequiresMigration == true ? [UiText.Get("XLSQ v1 首次保存需要另存 v2 副本。")] : [], canOverwrite);
        var window = new Window { Name = "CompatibilityReportDialog", Title = view.Title, Width = 820, Height = 620, MinWidth = 540, MinHeight = 400, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var root = new DockPanel { Margin = new(20) }; var header = new StackPanel { Spacing = 10 }; header.Children.Add(new TextBlock { Text = view.Headline, FontSize = 20, TextWrapping = TextWrapping.Wrap });
        header.Children.Add(new TextBlock { Text = view.Summary + "\n" + string.Join("\n", view.Notices), TextWrapping = TextWrapping.Wrap }); DockPanel.SetDock(header, Dock.Top); root.Children.Add(header);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, Margin = new(0, 10, 0, 0) };
        var copy = new Button { Content = UiText.Get("复制报告") }; copy.Click += async (_, _) => { try { await _clipboard.WriteAsync(new(CompatibilityReportPresentation.PlainText(report, view.Notices))); } catch (Exception ex) { Error(ex); } };
        buttons.Children.Add(copy);
        if (saveMode)
        {
            var save = new Button { Content = UiText.Get("另存受支持副本") };
            save.Click += (_, _) => window.Close(CompatibilitySaveDecision.SaveCopy); buttons.Children.Add(save);
            var overwrite = new Button { Content = UiText.Get("覆盖"), IsEnabled = canOverwrite,
                Background = new SolidColorBrush(Color.Parse("#FDE7E7")), Foreground = new SolidColorBrush(Color.Parse("#C62828")) };
            ToolTip.SetTip(overwrite, canOverwrite ? UiText.Get("覆盖当前文件，移除报告中列出的潜在损失内容。") : UiText.Get("当前文件需要另存副本，无法直接覆盖。"));
            overwrite.Click += (_, _) => { if (canOverwrite) window.Close(CompatibilitySaveDecision.Overwrite); }; buttons.Children.Add(overwrite);
        }
        var close = new Button { Content = saveMode ? UiText.Get("取消") : UiText.Get("关闭"), IsCancel = true, IsDefault = true };
        close.Click += (_, _) => window.Close(CompatibilitySaveDecision.Cancel); buttons.Children.Add(close); DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons);
        window.Opened += (_, _) => close.Focus();
        var groups = new StackPanel { Spacing = 10 };
        foreach (var group in view.Groups)
        {
            var content = new StackPanel { Spacing = 12 }; content.Children.Add(new TextBlock { Text = group.Summary, TextWrapping = TextWrapping.Wrap });
            content.Children.Add(new ListBox { MaxHeight = 260, ItemsSource = group.Items,
                ItemTemplate = new FuncDataTemplate<CompatibilityIssueView>((item, _) => new TextBlock
                { Text = item is null ? "" : UiText.Format($"{item.Feature} · {item.Location}\n{item.Detail}\n{item.Impact} · {item.Count} 处"), TextWrapping = TextWrapping.Wrap, Margin = new(0, 4) }) });
            groups.Children.Add(new Expander { Header = UiText.Format($"{group.Label}（{group.Items.Sum(item => item.Count)} 处）"), IsExpanded = group.ExpandedByDefault, Content = content });
        }
        root.Children.Add(new ScrollViewer { Content = groups }); window.Content = root; return ShowOwnedAsync<CompatibilitySaveDecision>(window);
    }
    private async Task ExportAsync(bool portable)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).ToArray();
        var missing = formulas.Count(c => c.Calculation?.Value is null);
        var stale = formulas.Count(c => c.Calculation?.Value is not null && (c.Calculation.InputRevision != snapshot.Revision || c.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)));
        var stalePivots=snapshot.Sheets.Sum(s=>s.Data.Pivots.Count(p=>p.Stale));
        var losses = snapshot.FileState?.CompatibilityIssues.ToArray() ?? [];
        var hostErrors = formulas.Count(c => c.Value.Kind == CellValueKind.Error && c.Value.ErrorCode is CellErrorCode.Script or CellErrorCode.Cycle);
        var message = UiText.Format($"所有公式将转为类型化结果值，不包含脚本。\n旧缓存 {stale} 个；宿主错误 {hostErrors} 个将映射为 #VALUE!。");
        if (portable)
        {
            if (snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(UiText.Get("只读公式工作簿的依赖范围未知，请使用明确的已有缓存值快照导出。"));
            var report = await _workspace.InspectPortableExportAsync(); missing = report.MissingResults; stale = report.StaleResults;
            message = UiText.Format($"保留 {report.RetainedFormulas} 个标准公式；冻结 {report.FrozenFormulas} 个脚本、不支持公式或其依赖为值；名称展开为引用。\n其中 {stale} 个结果可能过期。");
        }
        if (missing > 0) throw new InvalidOperationException(UiText.Format($"{missing} 个待导出的公式缺少结果，请先完成计算。"));
        if (stalePivots>0) message+=UiText.Format($"\n另有 {stalePivots} 个透视结果未刷新，将保留已有结果。");
        if (!portable && snapshot.Sheets.Any(s=>!s.Data.Pivots.IsEmpty)) message+=UiText.Get("\n透视表将转为普通结果值，不保留透视定义。");
        if (losses.Length > 0) message += UiText.Get("\n副本将移除以下未支持内容：\n") + string.Join("\n", losses);
        if (!await Dialogs.Confirm(this, portable ? UiText.Get("导出可移植公式") : UiText.Get("导出当前结果"), message + UiText.Get("\n是否导出？"))) return;
        var selected = await _files.SaveAsync(new(UiText.Get("导出工作簿副本"), [new(UiText.Get("Excel 工作簿"), ["*.xlsx"])],
            Path.GetFileNameWithoutExtension(_workspace.FilePath ?? UiText.Get("未命名")) + (portable ? UiText.Get("-公式.xlsx") : UiText.Get("-结果.xlsx")), ".xlsx"));
        if (selected is null) return;
        await _workspace.ExportAsync(selected, portable, stale > 0 || stalePivots>0, losses.Length > 0); _status.Text = UiText.Get("已导出 ") + selected.Path;
    }
}
