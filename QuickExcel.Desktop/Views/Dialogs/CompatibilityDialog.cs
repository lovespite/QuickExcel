using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>
/// Impact-first compatibility dialog shared by the report command and the save-time loss prompt.
/// Visual language matches the rest of the shell: white surface, green accent, quiet borders.
/// </summary>
internal sealed class CompatibilityDialog : Window
{
    private const int MaxRowsPerGroup = 40;
    private static readonly Brush Ink = new SolidColorBrush(Color.FromRgb(0x22, 0x35, 0x2E));
    private static readonly Brush Muted = new SolidColorBrush(Color.FromRgb(0x6E, 0x7C, 0x74));
    private static readonly Brush Soft = new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x52));
    private static readonly Brush Accent = new SolidColorBrush(Color.FromRgb(0x18, 0x75, 0x53));
    private static readonly Brush Border = new SolidColorBrush(Color.FromRgb(0xE1, 0xE9, 0xE4));
    private static readonly Brush Card = new SolidColorBrush(Color.FromRgb(0xFB, 0xFC, 0xFB));
    private static readonly Brush Surface = new SolidColorBrush(Color.FromRgb(0xF6, 0xF9, 0xF7));
    private static readonly Brush Pill = new SolidColorBrush(Color.FromRgb(0xED, 0xF3, 0xF0));
    private static readonly Brush WarningBackground = new SolidColorBrush(Color.FromRgb(0xFF, 0xF8, 0xEC));
    private static readonly Brush WarningInk = new SolidColorBrush(Color.FromRgb(0x7A, 0x52, 0x00));
    private static readonly Brush DangerInk = new SolidColorBrush(Color.FromRgb(0x9B, 0x2C, 0x24));
    private static readonly Brush InfoInk = new SolidColorBrush(Color.FromRgb(0x25, 0x5B, 0x86));

    internal CompatibilityReportView View { get; }
    internal string PlainText { get; }
    internal bool SaveMode { get; }

    public CompatibilityDialog(Window owner, WorkbookCompatibilityReport report, bool saveMode,
        IEnumerable<string>? notices = null)
        : this(owner, CompatibilityReportPresentation.Create(report, saveMode, notices),
            CompatibilityReportPresentation.PlainText(report, notices), saveMode)
    { }

    private CompatibilityDialog(Window owner, CompatibilityReportView view, string plainText, bool saveMode)
    {
        View = view;
        PlainText = plainText;
        SaveMode = saveMode;
        Owner = owner;
        Title = saveMode ? "保存兼容性检查" : "兼容性报告";
        Width = 780;
        Height = saveMode ? 620 : 660;
        MinWidth = 620;
        MinHeight = 460;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ShowInTaskbar = false;
        Background = Brushes.White;
        FontFamily = owner.FontFamily;
        FontSize = 13;
        UseLayoutRounding = true;
        AutomationProperties.SetName(this, Title);
        Content = BuildLayout();
    }

    private Grid BuildLayout()
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = BuildHeader();
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        var body = new StackPanel { Margin = new Thickness(24, 16, 24, 10) };
        if (View.Groups.Count == 0)
            body.Children.Add(BuildEmptyState());
        foreach (var group in View.Groups)
            body.Children.Add(BuildGroup(group));
        if (View.Notices.Count > 0)
            body.Children.Add(BuildNotices());
        var scroll = new ScrollViewer
        {
            Content = body,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Focusable = false
        };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        var footer = BuildFooter();
        Grid.SetRow(footer, 2);
        root.Children.Add(footer);
        return root;
    }

    private UIElement BuildHeader()
    {
        var content = new Grid();
        content.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        content.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var badge = new Border
        {
            Width = 36, Height = 36, CornerRadius = new CornerRadius(18),
            Background = SaveMode ? WarningInk : InfoInk,
            VerticalAlignment = VerticalAlignment.Top
        };
        badge.Child = new TextBlock
        {
            Text = SaveMode ? "!" : "i",
            Foreground = Brushes.White,
            FontSize = 19,
            FontWeight = FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        content.Children.Add(badge);

        var text = new StackPanel { Margin = new Thickness(14, 0, 0, 0) };
        var eyebrow = new TextBlock
        {
            Text = SaveMode ? "兼容性检查" : "导入与保存能力",
            FontSize = 11,
            FontWeight = FontWeights.SemiBold,
            Foreground = SaveMode ? WarningInk : Accent,
            Margin = new Thickness(0, 0, 0, 3)
        };
        var headline = new TextBlock
        {
            Text = View.Headline,
            FontSize = 19,
            FontWeight = FontWeights.SemiBold,
            Foreground = Ink,
            TextWrapping = TextWrapping.Wrap
        };
        var summary = new TextBlock
        {
            Text = View.Summary,
            FontSize = 12.5,
            Foreground = Soft,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 6, 0, 0),
            MaxWidth = 640
        };
        text.Children.Add(eyebrow);
        text.Children.Add(headline);
        text.Children.Add(summary);
        Grid.SetColumn(text, 1);
        content.Children.Add(text);

        return new Border
        {
            Background = SaveMode ? WarningBackground : Surface,
            BorderBrush = Border,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(24, 20, 24, 16),
            Child = content
        };
    }

    private UIElement BuildGroup(CompatibilityIssueGroup group)
    {
        var header = new StackPanel();
        var titleRow = new StackPanel { Orientation = Orientation.Horizontal };
        var dot = new Border
        {
            Width = 9, Height = 9, CornerRadius = new CornerRadius(5),
            Background = KindBrush(group.Kind),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 9, 0)
        };
        var label = new TextBlock
        {
            Text = group.Label,
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            Foreground = Ink,
            VerticalAlignment = VerticalAlignment.Center
        };
        var count = new Border
        {
            Background = Pill,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(7, 1, 7, 1),
            Margin = new Thickness(9, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = $"{group.Items.Sum(item => item.Count)} 处",
                FontSize = 11,
                Foreground = Muted
            }
        };
        titleRow.Children.Add(dot);
        titleRow.Children.Add(label);
        titleRow.Children.Add(count);
        header.Children.Add(titleRow);
        header.Children.Add(new TextBlock
        {
            Text = group.Summary,
            FontSize = 12,
            Foreground = Muted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(18, 4, 0, 0)
        });

        var items = new StackPanel { Margin = new Thickness(18, 8, 0, 0) };
        foreach (var item in group.Items.Take(MaxRowsPerGroup))
            items.Children.Add(BuildItem(item, group.Kind));
        if (group.Items.Count > MaxRowsPerGroup)
        {
            var hidden = group.Items.Skip(MaxRowsPerGroup).Sum(item => item.Count);
            items.Children.Add(new TextBlock
            {
                Text = $"另有 {hidden} 处同类问题未在此展开；可用“复制报告”查看全部。",
                FontSize = 11.5,
                Foreground = Muted,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(2, 2, 0, 4)
            });
        }

        var expander = new Expander
        {
            Header = header,
            Content = items,
            IsExpanded = group.ExpandedByDefault,
            Margin = new Thickness(0, 0, 0, 12),
            Padding = new Thickness(0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            HorizontalContentAlignment = HorizontalAlignment.Stretch
        };
        AutomationProperties.SetName(expander, $"{group.Label}，{group.Items.Sum(item => item.Count)} 处");
        return expander;
    }

    private UIElement BuildItem(CompatibilityIssueView item, CompatibilityKind kind)
    {
        var panel = new StackPanel();
        var titleRow = new Grid();
        titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var feature = new TextBlock
        {
            Text = item.Feature,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = Ink,
            TextWrapping = TextWrapping.Wrap
        };
        titleRow.Children.Add(feature);
        if (item.Count > 1)
        {
            var count = new Border
            {
                Background = Pill,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(7, 1, 7, 1),
                Margin = new Thickness(10, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Top,
                Child = new TextBlock { Text = $"{item.Count} 处", FontSize = 11, Foreground = Muted }
            };
            Grid.SetColumn(count, 1);
            titleRow.Children.Add(count);
        }
        panel.Children.Add(titleRow);

        var impact = new TextBlock { Margin = new Thickness(0, 5, 0, 0), TextWrapping = TextWrapping.Wrap };
        impact.Inlines.Add(new Run("可能影响  ")
        {
            Foreground = kind == CompatibilityKind.PotentialLoss ? DangerInk : Accent,
            FontWeight = FontWeights.SemiBold,
            FontSize = 11.5
        });
        impact.Inlines.Add(new Run(item.Impact) { Foreground = Ink, FontSize = 12.5 });
        panel.Children.Add(impact);

        var details = new StackPanel { Margin = new Thickness(0, 2, 0, 0) };
        details.Children.Add(DetailLine("位置", item.Location));
        details.Children.Add(DetailLine("说明", item.Detail));
        details.Children.Add(DetailLine("诊断代码", item.Code));
        var expander = new Expander
        {
            Header = new TextBlock { Text = "位置与原始说明", FontSize = 12, Foreground = Muted },
            Content = details,
            Margin = new Thickness(0, 6, 0, 0),
            Padding = new Thickness(0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent
        };
        panel.Children.Add(expander);

        var card = new Border
        {
            Background = Card,
            BorderBrush = Border,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 10, 12, 10),
            Margin = new Thickness(0, 0, 0, 7),
            Child = panel
        };
        AutomationProperties.SetName(card, $"{item.Feature}。可能影响：{item.Impact}");
        return card;
    }

    private static UIElement DetailLine(string label, string value) => new TextBlock
    {
        Text = label + "  " + value,
        FontSize = 12,
        Foreground = label == "诊断代码" ? Muted : Soft,
        TextWrapping = TextWrapping.Wrap,
        Margin = new Thickness(0, 2, 0, 0)
    };

    private UIElement BuildNotices()
    {
        var panel = new StackPanel();
        panel.Children.Add(new TextBlock
        {
            Text = "其他提示",
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = Ink,
            Margin = new Thickness(0, 2, 0, 6)
        });
        foreach (var notice in View.Notices)
            panel.Children.Add(new TextBlock
            {
                Text = "· " + notice,
                FontSize = 12,
                Foreground = Soft,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 4)
            });
        return new Border
        {
            Background = Surface,
            BorderBrush = Border,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 10, 12, 8),
            Margin = new Thickness(0, 2, 0, 12),
            Child = panel
        };
    }

    private UIElement BuildEmptyState() => new Border
    {
        Background = Surface,
        BorderBrush = Border,
        BorderThickness = new Thickness(1),
        CornerRadius = new CornerRadius(6),
        Padding = new Thickness(16, 14, 16, 14),
        Margin = new Thickness(0, 0, 0, 12),
        Child = new TextBlock
        {
            Text = "当前工作簿未发现兼容性问题。单元格类型、普通公式和基础格式都在已支持范围内。",
            FontSize = 12.5,
            Foreground = Soft,
            TextWrapping = TextWrapping.Wrap
        }
    };

    private UIElement BuildFooter()
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var copy = new Button { Content = "复制报告", MinWidth = 90, VerticalAlignment = VerticalAlignment.Center };
        copy.Click += (_, _) =>
        {
            try
            {
                Clipboard.SetText(PlainText);
                copy.Content = "已复制";
            }
            catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException or InvalidOperationException)
            {
                copy.Content = "无法复制";
            }
        };
        copy.Margin = new Thickness(0);
        grid.Children.Add(copy);

        var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        if (SaveMode)
        {
            var cancel = new Button { Content = "取消保存", IsCancel = true, MinWidth = 96, Margin = new Thickness(8, 0, 0, 0) };
            var proceed = new Button { Content = "另存受支持副本…", IsDefault = true, MinWidth = 150, Margin = new Thickness(8, 0, 0, 0) };
            ApplyPrimary(proceed);
            proceed.Click += (_, _) => DialogResult = true;
            actions.Children.Add(cancel);
            actions.Children.Add(proceed);
            Loaded += (_, _) => proceed.Focus();
        }
        else
        {
            var close = new Button { Content = "关闭", IsDefault = true, IsCancel = true, MinWidth = 90, Margin = new Thickness(8, 0, 0, 0) };
            ApplyPrimary(close);
            close.Click += (_, _) => Close();
            actions.Children.Add(close);
            Loaded += (_, _) => close.Focus();
        }
        Grid.SetColumn(actions, 1);
        grid.Children.Add(actions);

        return new Border
        {
            Background = Brushes.White,
            BorderBrush = Border,
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(20, 12, 20, 14),
            Child = grid
        };
    }

    private static void ApplyPrimary(Button button)
    {
        if (Application.Current?.TryFindResource("Primary") is Style style)
            button.Style = style;
        else
        {
            button.Background = Accent;
            button.Foreground = Brushes.White;
            button.BorderBrush = Accent;
        }
    }

    private static Brush KindBrush(CompatibilityKind kind) => kind switch
    {
        CompatibilityKind.PotentialLoss => DangerInk,
        CompatibilityKind.Rejected => DangerInk,
        CompatibilityKind.CannotCalculate => WarningInk,
        CompatibilityKind.PreservedOnly => Muted,
        _ => Accent
    };
}
