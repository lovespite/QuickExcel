using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

internal enum DataTransferFormat { Xlsx, Text, Odbc, Sqlite }
internal sealed record DataTransferRequest(bool Export, DataTransferFormat Format, int Mode);

/// <summary>The common entry point for workbook, delimited text and database transfers.</summary>
internal sealed class DataTransferWindow : Window
{
    internal ComboBox Direction { get; } = new() { ItemsSource = new[] { UiText.Get("导入"), UiText.Get("导出") }, SelectedIndex = 0, HorizontalAlignment = HorizontalAlignment.Stretch };
    internal ComboBox Format { get; } = new() { ItemsSource = new[] { "Excel (.xlsx)", "CSV / TSV", UiText.Get("SQL 数据库（ODBC）"), UiText.Get("本地 SQLite 文件") }, SelectedIndex = 0, HorizontalAlignment = HorizontalAlignment.Stretch };
    internal ComboBox Mode { get; } = new() { HorizontalAlignment = HorizontalAlignment.Stretch };
    internal DataTransferRequest Request => new(Direction.SelectedIndex == 1, (DataTransferFormat)Format.SelectedIndex, Mode.SelectedIndex);
    private readonly TextBlock _description = new() { TextWrapping = TextWrapping.Wrap, MinHeight = 88 };
    private readonly StackPanel _modePanel = new() { Spacing = 8 };

    internal DataTransferWindow(bool hasWorkbook)
    {
        Name = "DataTransferWindow"; Title = UiText.Get("数据导入 / 导出"); Width = 620; SizeToContent = SizeToContent.Height;
        CanResize = false; ShowInTaskbar = false; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new StackPanel { Margin = new(24), Spacing = 16 };
        root.Children.Add(new TextBlock { Text = Title, FontSize = 23, FontWeight = FontWeight.SemiBold });
        root.Children.Add(new TextBlock { Text = UiText.Get("在同一处交换工作簿、文本和数据库中的数据。"), TextWrapping = TextWrapping.Wrap });
        AddField(root, UiText.Get("操作"), Direction); AddField(root, UiText.Get("格式 / 数据源"), Format);
        AddField(_modePanel, UiText.Get("方式"), Mode); root.Children.Add(_modePanel); root.Children.Add(_description);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 10 };
        var next = new Button { Content = UiText.Get("下一步"), IsDefault = true };
        var cancel = new Button { Content = UiText.Get("取消"), IsCancel = true };
        next.Click += (_, _) => Close(Request); cancel.Click += (_, _) => Close(null);
        buttons.Children.Add(next); buttons.Children.Add(cancel); root.Children.Add(buttons); Content = root;
        Direction.SelectionChanged += (_, _) => { if (!hasWorkbook && Direction.SelectedIndex == 1) Direction.SelectedIndex = 0; UpdateOptions(); };
        Format.SelectionChanged += (_, _) => UpdateOptions();
        UpdateOptions();
    }

    private static void AddField(StackPanel root, string label, Control input)
    {
        AutomationProperties.SetName(input, label); root.Children.Add(new TextBlock { Text = label }); root.Children.Add(input);
    }

    private void UpdateOptions()
    {
        var export = Direction.SelectedIndex == 1; var format = (DataTransferFormat)Format.SelectedIndex;
        _modePanel.IsVisible = format == DataTransferFormat.Xlsx;
        Mode.ItemsSource = (export
            ? new[] { "整个工作簿：结果值", "整个工作簿：可移植公式", "当前工作表：结果值", "当前选区：结果值" }
            : new[] { "选择工作表，导入结果值", "完整工作簿副本（保留受支持内容）" }).Select(UiText.Get).ToArray();
        Mode.SelectedIndex = 0;
        _description.Text = UiText.Get(format switch
        {
            DataTransferFormat.Xlsx when export => "可导出整个工作簿、当前工作表或矩形选区。结果值导出不包含脚本；可移植公式保留可支持的 Excel 公式。",
            DataTransferFormat.Xlsx => "可选择多个工作表、修改目标名称，并在导入前处理同名冲突。结果值可导入当前工作簿或新工作簿；完整副本保留原有的兼容性限制。",
            DataTransferFormat.Text when export => "导出当前工作表或选区为 CSV / TSV。可选择原始值或显示文本，以及 UTF-8 BOM。",
            DataTransferFormat.Text => "预览 CSV / TSV，选择编码、分隔符与各列类型，再导入当前工作簿或新工作簿。默认按文本保留前导零。",
            DataTransferFormat.Odbc => "使用已安装的 ODBC 驱动连接 SQL 数据库。导入表或视图；导出可新建表或追加行。连接信息仅用于本次操作。",
            _ => "直接读取本地 SQLite 文件中的表或视图，或将当前工作表 / 选区导出到新表、追加到已有表。"
        });
    }
}
