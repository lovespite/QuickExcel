using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private static readonly (string Name, DifferentialStyle Style)[] ConditionalStyles =
    [
        ("浅红填充色深红色文本", new(FillColor: "FFFFC7CE", FontColor: "FF9C0006")),
        ("黄填充色深黄色文本", new(FillColor: "FFFFEB9C", FontColor: "FF9C6500")),
        ("绿填充色深绿色文本", new(FillColor: "FFC6EFCE", FontColor: "FF006100")),
        ("浅红色填充", new(FillColor: "FFFFC7CE")),
        ("红色文本", new(FontColor: "FF9C0006")),
        ("红色边框", new(Left: new(1, "FF9C0006"), Right: new(1, "FF9C0006"), Top: new(1, "FF9C0006"), Bottom: new(1, "FF9C0006")))
    ];

    private DropDownButton ConditionalFormatDropdown()
    {
        MenuItem Rule(string title, ConditionalPreset preset) => CommandMenu(title, () => QuickConditionalFormatAsync(title.TrimEnd('…'), preset));
        MenuItem Group(string title, int icon, params Control[] children) => new()
        { Header = title, Icon = new ConditionalFormatIcon(icon), ItemsSource = children, MinHeight = 44, MinWidth = 270 };
        var highlight = Group("突出显示单元格规则(H)", 0,
            Rule("大于…", ConditionalPreset.GreaterThan), Rule("小于…", ConditionalPreset.LessThan),
            Rule("介于…", ConditionalPreset.Between), Rule("等于…", ConditionalPreset.Equal), new Separator(),
            Rule("文本包含…", ConditionalPreset.ContainsText),
            new MenuItem { Header = "发生日期…（暂不支持）", IsEnabled = false },
            Rule("重复值…", ConditionalPreset.DuplicateValues), new Separator(), CommandMenu("其他规则…", NewConditionalFormatAsync));
        var ranking = Group("最前/最后规则(T)", 1,
            Rule("前 10 项…", ConditionalPreset.TopItems), Rule("前 10%…", ConditionalPreset.TopPercent),
            Rule("后 10 项…", ConditionalPreset.BottomItems), Rule("后 10%…", ConditionalPreset.BottomPercent), new Separator(),
            Rule("高于平均值…", ConditionalPreset.AboveAverage), Rule("低于平均值…", ConditionalPreset.BelowAverage));
        var create = CommandMenu("新建规则(N)…", NewConditionalFormatAsync); create.Icon = new ConditionalFormatIcon(0);
        var clear = Group("清除规则(C)", 5,
            CommandMenu("清除所选单元格的规则", () => ClearConditionalFormatsAsync(false)),
            CommandMenu("清除整个工作表的规则", () => ClearConditionalFormatsAsync(true)));
        clear.MinHeight = 34;
        var manage = CommandMenu("管理规则(R)…", ConditionalFormatsAsync); manage.Icon = RibbonIcons.Create(RibbonIcon.CreateTable);
        var content = new StackPanel { Spacing = 6, HorizontalAlignment = HorizontalAlignment.Center };
        content.Children.Add(new Viewbox { Width = 30, Height = 30, Child = new ConditionalFormatIcon(0) });
        content.Children.Add(new TextBlock { Text = "条件格式", HorizontalAlignment = HorizontalAlignment.Center });
        var button = new DropDownButton
        {
            Name = "ConditionalFormatDropdown", Content = content, Height = 82, MinWidth = 100, Padding = new(8, 6), Margin = new(2, 1),
            Background = Brushes.Transparent, BorderThickness = new(0),
            HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Center,
            Flyout = new MenuFlyout { ItemsSource = new Control[] { highlight, ranking, new Separator(),
                DataBarMenu(), ColorScaleMenu(), IconSetMenu(), new Separator(), create, clear, manage } }
        };
        button.Classes.Add("ribbon-command"); AutomationProperties.SetName(button, "条件格式");
        ToolTip.SetTip(button, "为所选单元格快速设置条件格式"); _mutationControls.Add(button); return button;
    }

    private void RequireConditionalFormatEditing()
    {
        var snapshot = Document.Snapshot;
        if (snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException("当前工作簿只读，无法编辑条件格式。");
        if (snapshot.GetSheet(snapshot.ActiveSheetId).Data.Protection is { Sheet: true, FormatCells: true }) throw new InvalidOperationException("工作表受保护，无法编辑条件格式。");
        if (snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException("工作簿含尚未支持的条件格式，需保留原始规则，无法编辑。");
    }

    private async Task QuickConditionalFormatAsync(string title, ConditionalPreset preset)
    {
        if (!await FlushEditsAsync()) return;
        RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange());
        var dialog = new FieldsDialog(title).Note("应用于 " + range);
        var ranking = preset is ConditionalPreset.TopItems or ConditionalPreset.TopPercent or ConditionalPreset.BottomItems or ConditionalPreset.BottomPercent;
        var percent = preset is ConditionalPreset.TopPercent or ConditionalPreset.BottomPercent;
        var duplicate = preset is ConditionalPreset.DuplicateValues or ConditionalPreset.UniqueValues;
        if (ranking) dialog.Text("rank", percent ? "百分比（1–100）" : "项数（1–1000）", "10").Note("仅比较数字；边界处相同数值一并突出显示。");
        else if (duplicate) dialog.Choice("duplicates", "突出显示", ["重复值", "唯一值"], preset == ConditionalPreset.UniqueValues ? 1 : 0);
        else if (preset is ConditionalPreset.AboveAverage or ConditionalPreset.BelowAverage) dialog.Note("按所选区域内数字的平均值判断，随数据变化更新。");
        else
        {
            dialog.Text("first", preset == ConditionalPreset.ContainsText ? "包含以下文本" : "比较值（数字、文本或以 = 开头的公式）", preset == ConditionalPreset.ContainsText ? "" : "0");
            if (preset == ConditionalPreset.Between) dialog.Text("second", "上限", "10");
        }
        dialog.Choice("style", "设置为", ConditionalStyles.Select(s => s.Name).ToArray());
        var sample = new Border { Padding = new(14, 8), CornerRadius = new(3), BorderThickness = new(1), Child = new TextBlock { Text = "格式预览  Aa 123", HorizontalAlignment = HorizontalAlignment.Center } };
        dialog.Note("确认后应用，可撤销。").Preview(sample);
        void Preview()
        {
            var style = ConditionalStyles[dialog.Index("style")].Style;
            sample.Background = SolidColorBrush.Parse("#" + (style.FillColor ?? "FFFFFFFF"));
            sample.BorderBrush = SolidColorBrush.Parse("#" + (style.Left?.Color ?? "FFD0D0D0"));
            ((TextBlock)sample.Child!).Foreground = SolidColorBrush.Parse("#" + (style.FontColor ?? "FF202020"));
        }
        ((ComboBox)dialog.Input("style")).SelectionChanged += (_, _) => Preview(); Preview();
        ConditionalFormatRule Read() => ConditionalFormatPresets.Create(range,
            duplicate ? dialog.Index("duplicates") == 0 ? ConditionalPreset.DuplicateValues : ConditionalPreset.UniqueValues : preset,
            ConditionalStyles[dialog.Index("style")].Style,
            ranking || duplicate || preset is ConditionalPreset.AboveAverage or ConditionalPreset.BelowAverage ? "0" : dialog.String("first"),
            preset == ConditionalPreset.Between ? dialog.String("second") : "0", ranking ? dialog.Integer("rank", 1, percent ? 100 : 1000) : 10);
        dialog.Validate = () => Read().ValidateBinding(snapshot, snapshot.ActiveSheetId);
        if (await dialog.ShowAsync(this)) await AddConditionalFormatAsync(snapshot, Read());
    }

    private async Task NewConditionalFormatAsync()
    {
        if (!await FlushEditsAsync()) return;
        RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; var dialog = RuleFields(Rectangle(SelectedRange()), null, out var read);
        dialog.Validate = () => read().ValidateBinding(snapshot, snapshot.ActiveSheetId);
        if (await dialog.ShowAsync(this)) await AddConditionalFormatAsync(snapshot, read());
    }

    private async Task AddConditionalFormatAsync(WorkbookSnapshot snapshot, ConditionalFormatRule rule)
    {
        var rules = snapshot.GetSheet(snapshot.ActiveSheetId).Data.ConditionalFormats.OrderBy(r => r.Priority).Prepend(rule).Select((r, i) => r.WithPriority(i + 1));
        await ExecuteAsync(new SetConditionalFormatsCommand(snapshot.ActiveSheetId, rules), snapshot.Revision);
    }

    private async Task ClearConditionalFormatsAsync(bool entireSheet)
    {
        if (!await FlushEditsAsync()) return;
        RequireConditionalFormatEditing();
        if (_grid.SelectedDrawingObject is not null) throw new InvalidOperationException("请先选择单元格。");
        var snapshot = Document.Snapshot;
        var ranges = _grid.SelectionRanges.Count == 0 ? new[] { Rectangle(SelectedRange()) } : _grid.SelectionRanges.Select(Rectangle).ToArray();
        await ExecuteAsync(new ClearConditionalFormatsCommand(snapshot.ActiveSheetId, entireSheet ? null : ranges), snapshot.Revision);
    }
}
