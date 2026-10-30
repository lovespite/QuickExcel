using System.Collections.Immutable;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal enum CleaningOperation { Duplicates, Split, Numbers }
internal sealed class CleaningDialog : Window
{
    internal CheckBox Header { get; } = new() { Content = "首行是标题，保留首行", IsChecked = true };
    internal TextBox Keys { get; } = new();
    internal ComboBox SplitKind { get; } = new() { ItemsSource = new[] { "分隔符", "固定宽度（中文双宽）" }, SelectedIndex = 0 };
    internal TextBox Delimiters { get; } = new() { Text = "," };
    internal CheckBox Collapse { get; } = new() { Content = "连续分隔符视为一个" };
    internal ComboBox Qualifier { get; } = new() { ItemsSource = new[] { "双引号", "单引号", "无" }, SelectedIndex = 0 };
    internal TextBox Breaks { get; } = new() { Text = "3,7" };
    internal TextBox Types { get; } = new();
    internal TextBox Destination { get; } = new();
    internal ComboBox Decimal { get; } = new() { ItemsSource = new[] { ".", "," }, SelectedIndex = 0 };
    internal TextBox GroupSeparator { get; } = new() { Text = ",", MaxLength = 1 };
    internal CheckBox TrailingMinus { get; } = new() { Content = "识别末尾负号（如 12-）", IsChecked = true };
    internal CheckBox Overwrite { get; } = new() { Content = "确认覆盖预览中列出的已有目标内容", IsEnabled = false };
    internal Button PreviewButton { get; } = new() { Content = "生成预览", MinWidth = 110, Margin = new(4) };
    internal Button ApplyButton { get; } = new() { Content = "应用预览", IsDefault = true, MinWidth = 110, Margin = new(4), IsEnabled = false };
    internal CleaningPreview? Preview { get; private set; }
    internal bool AllowOverwrite => Overwrite.IsChecked == true;
    private sealed record Row(string Address, string Before, string After, string Type);
    public CleaningDialog(Window owner, WorkbookSnapshot book, SheetId sheet, CellRectangle selection, CleaningOperation operation)
    {
        Owner = owner; Title = operation switch { CleaningOperation.Duplicates => "删除重复项", CleaningOperation.Split => "文本分列", _ => "文本数字转换" };
        Width = 800; Height = Math.Min(830, SystemParameters.WorkArea.Height - 32); WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Keys.Text = string.Join(",", Enumerable.Range(1, selection.Right - selection.Left + 1));
        Destination.Text = new CellRectangle(selection.Top, selection.Left, selection.Top, selection.Left).ToString();
        var root = new DockPanel { Margin = new(16) }; var footer = new StackPanel();
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(PreviewButton); buttons.Children.Add(ApplyButton); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 90, Margin = new(4) });
        footer.Children.Add(Overwrite); footer.Children.Add(buttons); DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(footer);
        var panel = new StackPanel(); var scroll = new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; root.Children.Add(scroll);
        panel.Children.Add(new TextBlock { Text = $"{book.GetSheet(sheet).Name} · {selection}，最多 20,000 格。先预览，应用后可一次撤销。", TextWrapping = TextWrapping.Wrap });
        void Field(string text, Control control)
        { panel.Children.Add(new TextBlock { Text = text, Margin = new(0, 6, 0, 2) }); panel.Children.Add(control); }
        if (operation == CleaningOperation.Duplicates)
        {
            panel.Children.Add(Header); Field("判重列序号（相对选区，从 1 开始，逗号分隔）", Keys);
            panel.Children.Add(new TextBlock { Text = "保留每组第一条记录，只压缩选区并清空尾部。文本 ID 与数值分开；按基础显示格式判重，不区分文本大小写。", TextWrapping = TextWrapping.Wrap });
        }
        if (operation == CleaningOperation.Split)
        {
            Field("方式", SplitKind); Field(@"分隔字符（可组合，\t 表示制表符）", Delimiters); panel.Children.Add(Collapse);
            Field("文本限定符", Qualifier); Field("固定宽度分界（累计宽度，如 3,7；中文宽 2、其他已验证字符宽 1）", Breaks);
            Field("各列类型：文本、数字、YMD、MDY、DMY、跳过；逗号分隔，未指定列默认文本", Types); Field("目标首格", Destination);
        }
        if (operation != CleaningOperation.Duplicates)
        {
            Field("数值小数点", Decimal); Field("数值千分位符（留空表示不使用）", GroupSeparator); panel.Children.Add(TrailingMinus);
            panel.Children.Add(new TextBlock { Text = "数字转换须明确选择。超过 15 位精度、无效数字或日期将拒绝整批操作；不猜测日期或布尔值。分列默认保留文本 ID、前导零和等号文本。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 5, 0, 5) });
        }
        var summary = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new(0, 8, 0, 5) };
        var error = new TextBlock { TextWrapping = TextWrapping.Wrap, Foreground = System.Windows.Media.Brushes.Firebrick, Margin = new(0, 6, 0, 4) };
        var grid = new DataGrid { IsReadOnly = true, AutoGenerateColumns = false, CanUserAddRows = false, Height = 230 };
        foreach (var (title, property, width) in new[] { ("位置", "Address", 65d), ("修改前", "Before", 220d), ("修改后", "After", 220d), ("结果类型", "Type", 90d) })
            grid.Columns.Add(new DataGridTextColumn { Header = title, Binding = new Binding(property), Width = width });
        panel.Children.Add(summary); panel.Children.Add(grid); panel.Children.Add(error);
        void Reset() { Preview = null; ApplyButton.IsEnabled = false; grid.ItemsSource = null; summary.Text = "设置已改变，请生成预览。"; error.Text = ""; Overwrite.IsEnabled = false; Overwrite.IsChecked = false; }
        foreach (var text in new[] { Keys, Delimiters, Breaks, Types, Destination, GroupSeparator }) text.TextChanged += (_, _) => Reset();
        foreach (var picker in new[] { SplitKind, Qualifier, Decimal }) picker.SelectionChanged += (_, _) => Reset();
        foreach (var check in new[] { Header, Collapse, TrailingMinus }) { check.Checked += (_, _) => Reset(); check.Unchecked += (_, _) => Reset(); }
        void EnableApply() => ApplyButton.IsEnabled = Preview is { } p && (p.Overwrites.IsEmpty || AllowOverwrite);
        Overwrite.Checked += (_, _) => EnableApply(); Overwrite.Unchecked += (_, _) => EnableApply();
        PreviewButton.Click += (_, _) =>
        {
            Reset();
            try
            {
                var numeric = new NumericTextOptions(Decimal.SelectedIndex == 0 ? '.' : ',', GroupSeparator.Text.Length == 0 ? null : GroupSeparator.Text[0], TrailingMinus.IsChecked == true);
                Preview = operation switch
                {
                    CleaningOperation.Duplicates => DataCleaning.RemoveDuplicates(book, sheet, selection, Keys.Text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => selection.Left + int.Parse(s, System.Globalization.CultureInfo.InvariantCulture) - 1), Header.IsChecked == true),
                    CleaningOperation.Numbers => DataCleaning.ConvertTextNumbers(book, sheet, selection, numeric),
                    _ => Split()
                };
                CleaningPreview Split()
                {
                    var target = CellRectangle.Parse(Destination.Text.Trim());
                    if (target.Top != target.Bottom || target.Left != target.Right) throw new ArgumentException("目标须为单个首格地址。");
                    var options = new TextSplitOptions(Delimiters.Text.Replace(@"\t", "\t"), Collapse.IsChecked == true, Qualifier.SelectedIndex switch { 0 => '"', 1 => '\'', _ => null },
                        SplitKind.SelectedIndex == 0 ? default : Breaks.Text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToImmutableArray());
                    var types = Types.Text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).Select(t => t.ToUpperInvariant() switch
                    { "文本" or "TEXT" => CleaningColumnType.Text, "数字" or "NUMBER" => CleaningColumnType.Number, "YMD" => CleaningColumnType.DateYmd,
                        "MDY" => CleaningColumnType.DateMdy, "DMY" => CleaningColumnType.DateDmy, "跳过" or "SKIP" => CleaningColumnType.Skip, _ => throw new ArgumentException("未知列类型：" + t) });
                    return DataCleaning.SplitText(book, sheet, selection, new(target.Top, target.Left), options, types, numeric);
                }
                string Display(CellContent? content, WorkbookSnapshot workbook) => content is null ? "" : content.Kind == CellContentKind.Formula ? content.FormulaText :
                    CellDisplayFormatter.Format(content.LiteralValue, workbook.Styles[content.StyleId], workbook.DateSystem);
                grid.ItemsSource = Preview.Changes.Take(200).Select(c => new Row(new CellRectangle(c.Address.Row, c.Address.Column, c.Address.Row, c.Address.Column).ToString(),
                    Display(c.Before, book), Display(c.After, Preview.Result), c.After.Kind == CellContentKind.Formula ? "公式" : c.After.LiteralValue.Kind switch
                    { CellValueKind.Text => "文本", CellValueKind.Number => "数值", CellValueKind.Blank => "空白", CellValueKind.Boolean => "布尔", _ => "错误" })).ToArray();
                summary.Text = $"修改 {Preview.Changes.Length} 格；删除重复记录 {Preview.RemovedRows.Length} 行；已有目标内容 {Preview.Overwrites.Length} 格。显示前 200 项修改。\n" +
                    string.Join("\n", Preview.Notices);
                if (!Preview.RemovedRows.IsEmpty) summary.Text += "\n删除的原行号：" + string.Join(", ", Preview.RemovedRows.Take(30).Select(r => r + 1));
                Overwrite.IsEnabled = !Preview.Overwrites.IsEmpty; EnableApply();
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(scroll.ScrollToEnd));
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException or FormatException or OverflowException)
            { Preview = null; error.Text = ex.Message.Length <= 400 ? ex.Message : ex.Message[..400] + "…"; }
        };
        ApplyButton.Click += (_, _) => { if (Preview is not null && ApplyButton.IsEnabled) DialogResult = true; };
        Content = root;
    }
}
