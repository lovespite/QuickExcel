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
        RibbonGroup("数据", "数据工具", RibbonButton("分列", SplitTextAsync), RibbonButton("快速填充  Ctrl+E", FlashFillAsync), RibbonButton("删除重复值", RemoveDuplicatesAsync),
            DataToolDropdown("数据验证", CommandMenu("数据验证…", EditValidationAsync), CommandMenu("清除验证", RemoveValidationAsync)), RibbonButton("文本转数字", ConvertNumbersAsync));
        RibbonGroup("数据", "预测与模拟", DataToolDropdown("模拟分析", CommandMenu("单变量求解…", GoalSeekAsync), CommandMenu("模拟运算表…", DataTableAsync), CommandMenu("方案分析…", ScenariosAsync)),
            RibbonButton("预测", ForecastAsync));
        RibbonGroup("数据", "条件格式", RibbonButton("条件格式管理", ConditionalFormatsAsync));
        RibbonGroup("数据", "分组与汇总", RibbonButton("分组", () => GroupAsync(false)), RibbonButton("取消分组", () => GroupAsync(true)), RibbonButton("展开/折叠", OutlineAsync), RibbonButton("分类汇总", SubtotalsAsync));
    }
    private static int[] ColumnNumbers(string text) => text.Split([',', '，', ' '], StringSplitOptions.RemoveEmptyEntries).Select(v => CellRectangle.Parse(v.Trim() + "1").Left).ToArray();
    private async Task RemoveDuplicatesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = DataToolRange(); var sheet = snapshot.GetSheet(snapshot.ActiveSheetId);
        var choices = Enumerable.Range(range.Left, range.Columns).Select(c => new CheckBox { Content = DataColumnLabel(sheet, range.Top, c), IsChecked = true }).ToArray();
        var list = new StackPanel { Spacing = 5 }; foreach (var choice in choices) list.Children.Add(choice);
        var dialog = new FieldsDialog("删除重复值").Note($"数据区域：{range}。勾选用于比较的列；保留第一次出现的记录，将先显示预览。")
            .Preview(new ScrollViewer { Content = list, MaxHeight = 240 }).Check("header", "首行为标题", true).Check("expand", "扩展到整条记录的数据列", true);
        dialog.Validate = () => { if (!choices.Any(c => c.IsChecked == true)) throw new ArgumentException("请至少选择一个判重列。"); };
        if (!await dialog.ShowAsync(this)) return;
        var keys = choices.Select((c, i) => (c, i)).Where(p => p.c.IsChecked == true).Select(p => range.Left + p.i).ToArray(); var header = dialog.Boolean("header"); var expand = dialog.Boolean("expand");
        var preview = await Task.Run(() => DataCleaning.RemoveDuplicates(snapshot, snapshot.ActiveSheetId, range, keys, header, expand));
        await ApplyCleaningAsync(preview);
    }
    internal static FieldsDialog NumericFields(FieldsDialog dialog, NumericTextOptions? options = null)
    {
        options ??= new();
        return dialog.Text("decimal", "小数点符号", options.DecimalSeparator.ToString())
            .Text("group", "分组符号（留空不使用）", options.GroupSeparator?.ToString() ?? "")
            .Check("minus", "识别尾部负号", options.TrailingMinus);
    }

    internal static NumericTextOptions ReadNumeric(FieldsDialog dialog)
    {
        var point = dialog.String("decimal"); var group = dialog.String("group");
        if (point.Length != 1 || group.Length > 1) throw new ArgumentException("小数点必须是一个字符，分组符号须为空或单个字符。");
        var options = new NumericTextOptions(point[0], group.Length == 0 ? null : group[0], dialog.Boolean("minus"));
        options.Validate();
        return options;
    }
    private async Task ConvertNumbersAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange()); var dialog = NumericFields(new FieldsDialog("文本转数字"));
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
    private async Task ApplyCleaningAsync(CleaningPreview preview)
    {
        var window = new Window { Title = "数据工具预览", Width = 820, Height = 580, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var root = new DockPanel { Margin = new(20) }; var header = new StackPanel { Spacing = 8 };
        header.Children.Add(new TextBlock { Text = $"变更 {preview.Changes.Length:N0} 格；重复记录 {preview.RemovedRows.Length:N0} 行；已有内容覆盖 {preview.Overwrites.Length:N0} 格。", TextWrapping = TextWrapping.Wrap });
        header.Children.Add(new TextBlock { Text = string.Join("\n", preview.Notices), TextWrapping = TextWrapping.Wrap }); DockPanel.SetDock(header, Dock.Top); root.Children.Add(header);
        var footer = new StackPanel { Spacing = 10 }; var overwrite = new CheckBox { Content = "我确认覆盖预览中已有的目标内容", IsVisible = preview.Overwrites.Length > 0 };
        footer.Children.Add(overwrite); var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 }; footer.Children.Add(buttons);
        var apply = new Button { Content = "应用", IsDefault = true, IsEnabled = preview.Overwrites.IsEmpty }; var cancel = new Button { Content = "取消", IsCancel = true };
        overwrite.IsCheckedChanged += (_, _) => apply.IsEnabled = preview.Overwrites.IsEmpty || overwrite.IsChecked == true;
        apply.Click += (_, _) => window.Close(true); cancel.Click += (_, _) => window.Close(false); buttons.Children.Add(apply); buttons.Children.Add(cancel);
        DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(footer);
        string Preview(CellContent? cell) { var text = cell is null ? "" : CellInput.EditText(cell); return text.Length <= 200 ? text : text[..200] + ""; }
        root.Children.Add(new ListBox { ItemsSource = preview.Changes, ItemTemplate = new FuncDataTemplate<CleaningCellPreview>((change, _) => change is null ? null : new TextBlock
        { Text = $"{QuickExcel.Formulas.FormulaSyntax.ColumnName(change.Address.Column)}{change.Address.Row + 1}   {Preview(change.Before)}  →  {Preview(change.After)}", TextTrimming = TextTrimming.CharacterEllipsis }) });
        window.Content = root;
        if (await ShowOwnedAsync<bool>(window)) { await ExecuteAsync(new ApplyCleaningCommand(preview, overwrite.IsChecked == true)); await CalculatePendingAsync(); }
    }
}
