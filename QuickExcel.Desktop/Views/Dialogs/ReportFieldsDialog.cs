using System.Windows;
using System.Windows.Controls;

namespace QuickExcel.Desktop;

internal sealed class ReportFieldsDialog : Window
{
    private readonly StackPanel _fields = new();
    private readonly TextBlock _error = new() { TextWrapping = TextWrapping.Wrap, Margin = new(0, 12, 0, 4), Foreground = System.Windows.Media.Brushes.Firebrick };
    internal Action? Apply { get; set; }
    internal Dictionary<string, Control> Fields { get; } = [];
    internal Button ApplyButton { get; } = new() { Content = "应用", IsDefault = true, MinWidth = 85, Margin = new(4) };
    internal ReportFieldsDialog(Window owner, string title)
    {
        Owner = owner; Title = title; Width = 520; Height = 620; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new DockPanel { Margin = new(18) }; var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var apply = ApplyButton;
        buttons.Children.Add(apply); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 85, Margin = new(4) });
        DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons); DockPanel.SetDock(_error, Dock.Bottom); root.Children.Add(_error);
        root.Children.Add(new ScrollViewer { Content = _fields, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }); Content = root;
        apply.Click += (_, _) => { try { Apply?.Invoke(); DialogResult = true; } catch (Exception ex) { _error.Text = ex.Message; } };
    }
    internal T Field<T>(string label, T control) where T : Control
    {
        Fields[label.Length == 0 && control is CheckBox check ? (string)check.Content : label] = control;
        _fields.Children.Add(new TextBlock { Text = label, TextWrapping = TextWrapping.Wrap, Margin = new(0, 9, 0, 3) }); _fields.Children.Add(control); return control;
    }
    internal TextBox Text(string label, string? value) => Field(label, new TextBox { Text = value ?? "" });
    internal ComboBox Choice(string label, string[] choices, int selected) => Field(label, new ComboBox { ItemsSource = choices, SelectedIndex = selected });
    internal CheckBox Check(string label, bool value) => Field("", new CheckBox { Content = label, IsChecked = value });
}
