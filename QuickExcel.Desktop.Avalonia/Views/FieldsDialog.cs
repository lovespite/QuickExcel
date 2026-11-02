using System.Globalization;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Native labelled fields with validation that keeps invalid input in the dialog.</summary>
internal sealed class FieldsDialog : Window
{
    private readonly StackPanel _fields = new() { Spacing = 10 };
    private readonly Dictionary<string, Control> _inputs = [];
    private readonly TextBlock _error = new() { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };
    public FieldsDialog(string title)
    {
        Title = title; Width = 520; MaxHeight = 760; SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Spacing = 14, Margin = new(20) };
        panel.Children.Add(new ScrollViewer { MaxHeight = 550, Content = _fields }); panel.Children.Add(_error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var ok = new Button { Content = "确定", IsDefault = true }; var cancel = new Button { Content = "取消", IsCancel = true };
        ok.Click += (_, _) => { try { Validate?.Invoke(); Close(true); } catch (Exception ex) { _error.Text = ex.Message; } };
        cancel.Click += (_, _) => Close(false); buttons.Children.Add(ok); buttons.Children.Add(cancel); panel.Children.Add(buttons); Content = panel;
        Opened += (_, _) => _inputs.Values.FirstOrDefault()?.Focus();
    }
    public Action? Validate { get; set; }
    public FieldsDialog Note(string text) { _fields.Children.Add(new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap }); return this; }
    private void Add(string key, string label, Control input)
    {
        _inputs.Add(key, input); AutomationProperties.SetName(input, label);
        _fields.Children.Add(new TextBlock { Text = label }); _fields.Children.Add(input);
    }
    public FieldsDialog Text(string key, string label, string initial = "", bool multiline = false)
    {
        Add(key, label, new TextBox { Text = initial, AcceptsReturn = multiline, TextWrapping = multiline ? TextWrapping.Wrap : TextWrapping.NoWrap,
            MinHeight = multiline ? 90 : 30 }); return this;
    }
    public FieldsDialog Choice(string key, string label, string[] choices, int selected = 0)
    { Add(key, label, new ComboBox { ItemsSource = choices, SelectedIndex = Math.Clamp(selected, 0, choices.Length - 1), HorizontalAlignment = HorizontalAlignment.Stretch }); return this; }
    public FieldsDialog Check(string key, string label, bool value = false)
    { Add(key, label, new CheckBox { IsChecked = value, Content = label }); return this; }
    public string String(string key) => ((TextBox)_inputs[key]).Text ?? "";
    public int Index(string key) => ((ComboBox)_inputs[key]).SelectedIndex;
    public bool Boolean(string key) => ((CheckBox)_inputs[key]).IsChecked == true;
    public double Number(string key, double minimum, double maximum)
    {
        if (!double.TryParse(String(key), NumberStyles.Float, CultureInfo.CurrentCulture, out var value) || !double.IsFinite(value) || value < minimum || value > maximum)
            throw new ArgumentException($"请输入 {minimum} 到 {maximum} 之间的数字。");
        return value;
    }
    public int Integer(string key, int minimum, int maximum)
    { var number = Number(key, minimum, maximum); if (Math.Truncate(number) != number) throw new ArgumentException("请输入整数。"); return (int)number; }
    internal Control Input(string key) => _inputs[key];
    public Task<bool> ShowAsync(Window owner) => owner is MainWindow main ? main.ShowOwnedAsync<bool>(this) : ShowDialog<bool>(owner);
}
