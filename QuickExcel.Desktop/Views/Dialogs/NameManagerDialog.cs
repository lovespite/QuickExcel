using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Desktop;

internal sealed class NameManagerDialog : Window
{
    internal sealed record ScopeChoice(SheetId? Id, string Label);
    internal sealed record NameChoice(DefinedName Definition, string Label);
    internal ListBox Names { get; } = new() { DisplayMemberPath = "Label", Height = 150, Margin = new(0, 6, 0, 12) };
    internal TextBox NameInput { get; } = new() { Margin = new(0, 4, 0, 8) };
    internal ComboBox ScopePicker { get; } = new() { DisplayMemberPath = "Label", Margin = new(0, 4, 0, 8) };
    internal TextBox FormulaInput { get; } = new() { Margin = new(0, 4, 0, 8), MinHeight = 60, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap };
    internal Button NewButton { get; } = new() { Content = "新建", MinWidth = 70 };
    internal Button ApplyButton { get; } = new() { Content = "应用", IsDefault = true, MinWidth = 70, Margin = new(8, 0, 0, 0) };
    internal Button DeleteButton { get; } = new() { Content = "删除", MinWidth = 70, Margin = new(8, 0, 0, 0), IsEnabled = false };
    public DocumentCommand? Command { get; private set; }
    public NameManagerDialog(Window owner, WorkbookSnapshot snapshot, string initialFormula)
    {
        Owner = owner; Title = "名称管理"; SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize;
        var readOnly = snapshot.FileState?.FormulaReadOnly == true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Width = 540, Margin = new(20) };
        panel.Children.Add(new TextBlock { Text = "名称可用于公式。单元格公式中的局部名称优先于同名全局名称；修改一次可撤销。", TextWrapping = TextWrapping.Wrap });
        Names.ItemsSource = snapshot.DefinedNames.OrderBy(n => n.Scope is null ? "" : snapshot.GetSheet(n.Scope.Value).Name).ThenBy(n => n.Name)
            .Select(n => new NameChoice(n, n.Name + "  ·  " + (n.Scope is null ? "工作簿" : snapshot.GetSheet(n.Scope.Value).Name))).ToArray();
        ScopePicker.ItemsSource = new[] { new ScopeChoice(null, "工作簿") }.Concat(snapshot.Sheets.Select(s => new ScopeChoice(s.Id, s.Name))).ToArray();
        ScopePicker.SelectedIndex = 0; FormulaInput.Text = initialFormula;
        panel.Children.Add(Names); panel.Children.Add(new TextBlock { Text = "名称" }); panel.Children.Add(NameInput);
        panel.Children.Add(new TextBlock { Text = "作用域（现有名称的作用域保持不变）" }); panel.Children.Add(ScopePicker);
        panel.Children.Add(new TextBlock { Text = "引用位置或公式（例如 ='Sheet1'!$A$1:$A$10；引用须带表名并使用绝对地址）", TextWrapping = TextWrapping.Wrap }); panel.Children.Add(FormulaInput);
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 8) }; panel.Children.Add(error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        foreach (var button in new[] { NewButton, DeleteButton, ApplyButton, new Button { Content = "关闭", IsCancel = true, MinWidth = 70, Margin = new(8, 0, 0, 0) } }) buttons.Children.Add(button);
        panel.Children.Add(buttons);
        Names.SelectionChanged += (_, _) =>
        {
            if (Names.SelectedItem is not NameChoice choice) return;
            NameInput.Text = choice.Definition.Name; FormulaInput.Text = choice.Definition.Formula;
            ScopePicker.SelectedItem = ScopePicker.Items.Cast<ScopeChoice>().Single(s => s.Id == choice.Definition.Scope);
            ScopePicker.IsEnabled = false; DeleteButton.IsEnabled = !readOnly; error.Text = "";
        };
        NewButton.Click += (_, _) => { Names.SelectedIndex = -1; NameInput.Text = ""; FormulaInput.Text = initialFormula; ScopePicker.IsEnabled = true; ScopePicker.SelectedIndex = 0; DeleteButton.IsEnabled = false; error.Text = ""; NameInput.Focus(); };
        ApplyButton.Click += (_, _) =>
        {
            try
            {
                var scope = ((ScopeChoice)ScopePicker.SelectedItem).Id; var original = (Names.SelectedItem as NameChoice)?.Definition;
                var definition = new DefinedName(NameInput.Text.Trim(), FormulaInput.Text.Trim(), scope);
                if (snapshot.DefinedNames.Any(n => n != original && n.Scope == scope && n.Name.Equals(definition.Name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("该作用域已有同名名称。");
                _ = snapshot.WithDefinedNames(snapshot.DefinedNames.Where(n => n != original).Append(definition));
                Command = original is null ? new DefineNameCommand(definition) : new EditDefinedNameCommand(original.Name, original.Scope, definition.Name, definition.Formula);
                DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or FormatException) { error.Text = ex.Message; }
        };
        DeleteButton.Click += (_, _) => { if (Names.SelectedItem is NameChoice choice) { Command = new DeleteDefinedNameCommand(choice.Definition.Name, choice.Definition.Scope); DialogResult = true; } };
        NameInput.IsReadOnly = FormulaInput.IsReadOnly = readOnly;
        NewButton.IsEnabled = ApplyButton.IsEnabled = !readOnly; ScopePicker.IsEnabled = !readOnly;
        Content = panel;
    }
}
