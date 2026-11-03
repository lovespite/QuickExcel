using System.Collections.Immutable;
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
        RibbonGroup("数据", "数据清理", RibbonButton("删除重复记录", RemoveDuplicatesAsync), RibbonButton("文本分列", SplitTextAsync), RibbonButton("文本转数字", ConvertNumbersAsync));
        RibbonGroup("数据", "条件格式", RibbonButton("条件格式管理", ConditionalFormatsAsync));
        RibbonGroup("数据", "分组与汇总", RibbonButton("分组", () => GroupAsync(false)), RibbonButton("取消分组", () => GroupAsync(true)), RibbonButton("展开/折叠", OutlineAsync), RibbonButton("分类汇总", SubtotalsAsync));
    }
    private static int[] ColumnNumbers(string text) => text.Split([',', '，', ' '], StringSplitOptions.RemoveEmptyEntries).Select(v => CellRectangle.Parse(v.Trim() + "1").Left).ToArray();
    private async Task RemoveDuplicatesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange());
        var dialog = new FieldsDialog("删除重复记录").Note("同一记录的其他列随去重移动；将先显示预览。").Text("keys", "比较列（列字母，用逗号分隔）", string.Join(',', Enumerable.Range(range.Left, range.Columns).Select(QuickExcel.Formulas.FormulaSyntax.ColumnName)))
            .Check("header", "首行为标题", true).Check("expand", "扩展到整条记录的数据列");
        dialog.Validate = () => ColumnNumbers(dialog.String("keys")); if (!await dialog.ShowAsync(this)) return;
        var keys = ColumnNumbers(dialog.String("keys")); var header = dialog.Boolean("header"); var expand = dialog.Boolean("expand");
        var preview = await Task.Run(() => DataCleaning.RemoveDuplicates(snapshot, snapshot.ActiveSheetId, range, keys, header, expand));
        await ApplyCleaningAsync(preview);
    }
    private static FieldsDialog NumericFields(FieldsDialog dialog) => dialog.Text("decimal", "小数点符号", ".").Text("group", "分组符号（留空不使用）", ",").Check("minus", "识别尾部负号", true);
    private static NumericTextOptions ReadNumeric(FieldsDialog dialog)
    {
        var point = dialog.String("decimal"); var group = dialog.String("group");
        if (point.Length != 1 || group.Length > 1 || point == group) throw new ArgumentException("小数点必须是一个字符，分组符号须为空或不同的单个字符。");
        return new(point[0], group.Length == 0 ? null : group[0], dialog.Boolean("minus"));
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
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange());
        var dialog = NumericFields(new FieldsDialog("文本分列").Note("来源须为一列。未指定类型的输出列保留文本，不会自动执行公式。")
            .Text("target", "目标左上角", new CellRectangle(range.Top, range.Left, range.Top, range.Left).ToString())
            .Choice("mode", "分列方式", ["分隔符", "固定宽度"])
            .Text("delimiter", "分隔字符（输入 \\t 表示制表符）", ",").Text("breaks", "固定宽度切分位置（字符数，用逗号分隔）", "")
            .Check("collapse", "合并连续分隔符").Choice("quote", "文本限定符", ["双引号", "单引号", "无"])
            .Text("types", "各输出列类型（逗号分隔：文本、数字、日期YMD、日期MDY、日期DMY、跳过）", "文本"));
        (CellPosition Target, TextSplitOptions Split, CleaningColumnType[] Types, NumericTextOptions Numbers) Read()
        {
            var target = CellRectangle.Parse(dialog.String("target")); if (target.Count != 1) throw new ArgumentException("目标必须是单个单元格。");
            var breaks = dialog.Index("mode") == 0 ? ImmutableArray<int>.Empty : dialog.String("breaks").Split(',', StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToImmutableArray();
            string[] labels = ["文本", "数字", "日期YMD", "日期MDY", "日期DMY", "跳过"];
            var types = dialog.String("types").Split([',', '，'], StringSplitOptions.RemoveEmptyEntries).Select(v =>
            { var index = Array.IndexOf(labels, v.Trim()); if (index < 0) throw new ArgumentException("未知输出列类型：" + v); return (CleaningColumnType)index; }).ToArray();
            return (new(target.Top, target.Left), new(dialog.String("delimiter").Replace("\\t", "\t"), dialog.Boolean("collapse"), dialog.Index("quote") switch { 0 => '"', 1 => '\'', _ => null }, breaks), types, ReadNumeric(dialog));
        }
        dialog.Validate = () => Read(); if (!await dialog.ShowAsync(this)) return;
        var options = Read(); var preview = await Task.Run(() => DataCleaning.SplitText(snapshot, snapshot.ActiveSheetId, range, options.Target, options.Split, options.Types, options.Numbers));
        await ApplyCleaningAsync(preview);
    }
    private async Task ApplyCleaningAsync(CleaningPreview preview)
    {
        var window = new Window { Title = "数据清理预览", Width = 820, Height = 580, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
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
        root.Children.Add(new ListBox { ItemsSource = preview.Changes, ItemTemplate = new FuncDataTemplate<CleaningCellPreview>((change, _) => new TextBlock
        { Text = $"{QuickExcel.Formulas.FormulaSyntax.ColumnName(change!.Address.Column)}{change.Address.Row + 1}   {Preview(change.Before)}  →  {Preview(change.After)}", TextTrimming = TextTrimming.CharacterEllipsis }) });
        window.Content = root;
        if (await ShowOwnedAsync<bool>(window)) { await ExecuteAsync(new ApplyCleaningCommand(preview, overwrite.IsChecked == true)); await CalculatePendingAsync(); }
    }
}
