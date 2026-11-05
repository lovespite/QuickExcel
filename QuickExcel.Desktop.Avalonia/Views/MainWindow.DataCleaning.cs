using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallCleaning()
    {
        InstallOutline();
        RibbonGroup(UiText.Get("数据"), UiText.Get("数据工具"), RibbonButton(RibbonIcon.SplitText, UiText.Get("分列"), SplitTextAsync), RibbonButton(RibbonIcon.FlashFill, UiText.Get("快速填充  Ctrl+E"), FlashFillAsync), RibbonButton(RibbonIcon.RemoveDuplicates, UiText.Get("删除重复值"), RemoveDuplicatesAsync),
            DataToolDropdown(RibbonIcon.Validation, UiText.Get("数据验证"), CommandMenu(UiText.Get("数据验证…"), EditValidationAsync), CommandMenu(UiText.Get("清除验证"), RemoveValidationAsync)), RibbonButton(RibbonIcon.ConvertNumbers, UiText.Get("文本转数字"), ConvertNumbersAsync));
        RibbonGroup(UiText.Get("数据"), UiText.Get("预测与模拟"), DataToolDropdown(RibbonIcon.WhatIf, UiText.Get("模拟分析"), CommandMenu(UiText.Get("单变量求解…"), GoalSeekAsync), CommandMenu(UiText.Get("模拟运算表…"), DataTableAsync), CommandMenu(UiText.Get("刷新模拟运算表"), RefreshDataTablesAsync), CommandMenu(UiText.Get("方案分析…"), ScenariosAsync)),
            RibbonButton(RibbonIcon.Forecast, UiText.Get("预测"), ForecastAsync));
        RibbonGroup(UiText.Get("数据"), UiText.Get("条件格式"), RibbonButton(RibbonIcon.ConditionalRules, UiText.Get("条件格式管理"), ConditionalFormatsAsync));
        RibbonGroup(UiText.Get("数据"), UiText.Get("分组与汇总"), RibbonButton(RibbonIcon.Group, UiText.Get("分组"), () => GroupAsync(false)), RibbonButton(RibbonIcon.Ungroup, UiText.Get("取消分组"), () => GroupAsync(true)), RibbonButton(RibbonIcon.Outline, UiText.Get("展开/折叠"), OutlineAsync), RibbonButton(RibbonIcon.Subtotals, UiText.Get("分类汇总"), SubtotalsAsync));
    }
    private static int[] ColumnNumbers(string text) => text.Split([',', '，', ' '], StringSplitOptions.RemoveEmptyEntries).Select(v => CellRectangle.Parse(v.Trim() + "1").Left).ToArray();
    private async Task RemoveDuplicatesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = DataToolRange(); var sheet = snapshot.GetSheet(snapshot.ActiveSheetId);
        var choices = Enumerable.Range(range.Left, range.Columns).Select(c => new CheckBox { Content = DataColumnLabel(sheet, range.Top, c), IsChecked = true }).ToArray();
        var list = new StackPanel { Spacing = 5 }; foreach (var choice in choices) list.Children.Add(choice);
        var dialog = new FieldsDialog(UiText.Get("删除重复值")).Note(UiText.Format($"数据区域：{range}。勾选用于比较的列；保留第一次出现的记录，将先显示预览。"))
            .Preview(new ScrollViewer { Content = list, MaxHeight = 240 }).Check("header", UiText.Get("首行为标题"), true).Check("expand", UiText.Get("扩展到整条记录的数据列"), true);
        dialog.Validate = () => { if (!choices.Any(c => c.IsChecked == true)) throw new ArgumentException(UiText.Get("请至少选择一个判重列。")); };
        if (!await dialog.ShowAsync(this)) return;
        var keys = choices.Select((c, i) => (c, i)).Where(p => p.c.IsChecked == true).Select(p => range.Left + p.i).ToArray(); var header = dialog.Boolean("header"); var expand = dialog.Boolean("expand");
        var preview = await Task.Run(() => DataCleaning.RemoveDuplicates(snapshot, snapshot.ActiveSheetId, range, keys, header, expand));
        await ApplyCleaningAsync(preview);
    }
    internal static FieldsDialog NumericFields(FieldsDialog dialog, NumericTextOptions? options = null)
    {
        options ??= new();
        return dialog.Text("decimal", UiText.Get("小数点符号"), options.DecimalSeparator.ToString())
            .Text("group", UiText.Get("分组符号（留空不使用）"), options.GroupSeparator?.ToString() ?? "")
            .Check("minus", UiText.Get("识别尾部负号"), options.TrailingMinus);
    }

    internal static NumericTextOptions ReadNumeric(FieldsDialog dialog)
    {
        var point = dialog.String("decimal"); var group = dialog.String("group");
        if (point.Length != 1 || group.Length > 1) throw new ArgumentException(UiText.Get("小数点必须是一个字符，分组符号须为空或单个字符。"));
        var options = new NumericTextOptions(point[0], group.Length == 0 ? null : group[0], dialog.Boolean("minus"));
        options.Validate();
        return options;
    }
    private async Task ConvertNumbersAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange()); var dialog = NumericFields(new FieldsDialog(UiText.Get("文本转数字")));
        dialog.Validate = () => ReadNumeric(dialog); if (!await dialog.ShowAsync(this)) return;
        var options = ReadNumeric(dialog); var preview = await Task.Run(() => DataCleaning.ConvertTextNumbers(snapshot, snapshot.ActiveSheetId, range, options));
        await ApplyCleaningAsync(preview);
    }
    private async Task SplitTextAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = DataToolRange(singleColumn: true);
        var rows = await Task.Run(() => DataCleaning.ReadSplitSource(snapshot, snapshot.ActiveSheetId, range));
        var wizard = new TextSplitWizard(snapshot, range, rows);
        if (!await ShowOwnedAsync<bool>(wizard) || wizard.Result is not { } result) return;
        await ExecuteAsync(new ApplyCleaningCommand(result, wizard.AllowOverwrite));
        await CalculatePendingAsync();
    }
    private async Task<bool> ApplyCleaningAsync(CleaningPreview preview)
    {
        var window = new Window { Title = UiText.Get("数据工具预览"), Width = 820, Height = 580, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var root = new DockPanel { Margin = new(20) }; var header = new StackPanel { Spacing = 8 };
        header.Children.Add(new TextBlock { Text = UiText.Format($"变更 {preview.Changes.Length:N0} 格；重复记录 {preview.RemovedRows.Length:N0} 行；已有内容覆盖 {preview.Overwrites.Length:N0} 格。"), TextWrapping = TextWrapping.Wrap });
        header.Children.Add(new TextBlock { Text = string.Join("\n", preview.Notices), TextWrapping = TextWrapping.Wrap }); DockPanel.SetDock(header, Dock.Top); root.Children.Add(header);
        var footer = new StackPanel { Spacing = 10 }; var overwrite = new CheckBox { Content = UiText.Get("我确认覆盖预览中已有的目标内容"), IsVisible = preview.Overwrites.Length > 0 };
        footer.Children.Add(overwrite); var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 }; footer.Children.Add(buttons);
        var apply = new Button { Content = UiText.Get("应用"), IsDefault = true, IsEnabled = preview.Overwrites.IsEmpty }; var cancel = new Button { Content = UiText.Get("取消"), IsCancel = true };
        overwrite.IsCheckedChanged += (_, _) => apply.IsEnabled = preview.Overwrites.IsEmpty || overwrite.IsChecked == true;
        apply.Click += (_, _) => window.Close(true); cancel.Click += (_, _) => window.Close(false); buttons.Children.Add(apply); buttons.Children.Add(cancel);
        DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(footer);
        string Preview(CellContent? cell) { var text = cell is null ? "" : CellInput.EditText(cell); return text.Length <= 200 ? text : text[..200] + ""; }
        root.Children.Add(new ListBox { ItemsSource = preview.Changes, ItemTemplate = new FuncDataTemplate<CleaningCellPreview>((change, _) => change is null ? null : new TextBlock
        { Text = $"{QuickExcel.Formulas.FormulaSyntax.ColumnName(change.Address.Column)}{change.Address.Row + 1}   {Preview(change.Before)}  →  {Preview(change.After)}", TextTrimming = TextTrimming.CharacterEllipsis }) });
        window.Content = root;
        if (!await ShowOwnedAsync<bool>(window) || !await ExecuteAsync(new ApplyCleaningCommand(preview, overwrite.IsChecked == true))) return false;
        await CalculatePendingAsync();
        return true;
    }
}
