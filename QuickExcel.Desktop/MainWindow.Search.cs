using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private Action<Window>? _smokeSearchDriver;
    private void Find_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        ShowWorkbookPage();
        var sheetId = _document.Snapshot.Sheets[_sheetIndex].Id;
        // Capture exact discontiguous targets before search navigation changes the selection.
        var selection = GridSurface.Selection.Select(p => new CellAddress(sheetId, p.Row, p.Column)).ToHashSet();
        var activeSheet = _sheetIndex;
        var panel = new DockPanel { Margin = new(20) };
        var fields = new StackPanel(); DockPanel.SetDock(fields, Dock.Top); panel.Children.Add(fields);
        TextBox Field(string label)
        {
            fields.Children.Add(new TextBlock { Text = label, Margin = new(0, 0, 0, 4) });
            var box = new TextBox { Margin = new(0, 0, 0, 10) }; AutomationProperties.SetName(box, label); fields.Children.Add(box); return box;
        }
        var query = Field("查找"); var replacement = Field("替换为");
        var selectors = new WrapPanel { Margin = new(0, 0, 0, 10) }; fields.Children.Add(selectors);
        var scope = new ComboBox { ItemsSource = new[] { "当前工作表", "整个工作簿", "选区" }, SelectedIndex = 0, MinWidth = 140, Margin = new(0, 0, 10, 0) };
        var target = new ComboBox { ItemsSource = new[] { "显示值", "公式源码" }, SelectedIndex = 0, MinWidth = 120 };
        AutomationProperties.SetName(scope, "查找范围"); AutomationProperties.SetName(target, "查找对象"); selectors.Children.Add(scope); selectors.Children.Add(target);
        var options = new WrapPanel { Margin = new(0, 0, 0, 12) }; fields.Children.Add(options);
        var matchCase = new CheckBox { Content = "区分大小写", Margin = new(0, 0, 14, 0) };
        var whole = new CheckBox { Content = "整格匹配", Margin = new(0, 0, 14, 0) };
        var hidden = new CheckBox { Content = "包含隐藏行列" };
        options.Children.Add(matchCase); options.Children.Add(whole); options.Children.Add(hidden);
        var actions = new WrapPanel { Margin = new(0, 0, 0, 12) }; fields.Children.Add(actions);
        var find = new Button { Content = "查找全部", IsDefault = true };
        var replace = new Button { Content = "替换所选" }; var all = new Button { Content = "全部替换" };
        actions.Children.Add(find); actions.Children.Add(replace); actions.Children.Add(all);
        var status = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 10) }; fields.Children.Add(status);
        AutomationProperties.SetLiveSetting(status, AutomationLiveSetting.Polite);
        var results = new ListBox { DisplayMemberPath = "Location" }; AutomationProperties.SetName(results, "查找结果，Enter 定位"); panel.Children.Add(results);
        var dialog = ToolDialog("查找与替换", panel, 620, 580);
        SearchOptions Options() => new(query.Text, (SearchScope)scope.SelectedIndex, (SearchTarget)target.SelectedIndex, matchCase.IsChecked == true, whole.IsChecked == true, hidden.IsChecked == true);
        IReadOnlyList<SearchHit> FindMatches() => WorkbookSearch.Find(_document.Snapshot, activeSheet, Options(), selection,
            (s, r) => _sheetViews.TryGetValue(_document.Snapshot.Sheets[s].Id, out var view) && view.IsFilteredOut(_document.Snapshot, s, r));
        void Search()
        {
            try { var matches = FindMatches(); results.ItemsSource = matches; status.Text = $"{matches.Count:N0} 个匹配"; }
            catch (Exception ex) { results.ItemsSource = null; status.Text = ex.Message; }
        }
        void Replace(bool everything)
        {
            try
            {
                var matches = FindMatches();
                if (!everything) matches = results.SelectedItem is SearchHit selected ? matches.Where(hit => hit.Address == selected.Address).ToArray() : [];
                if (matches.Count == 0) { status.Text = "没有可替换的匹配。"; return; }
                var command = WorkbookSearch.Replace(matches, Options(), replacement.Text);
                var changed = _editor.Execute(command);
                _calculationPending |= changed.Changed; RefreshSheet(); UpdateChrome();
                results.ItemsSource = null; status.Text = $"已替换 {matches.Count:N0} 格，可一次撤销。";
            }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        void Locate()
        {
            if (results.SelectedItem is not SearchHit hit) return;
            try { NavigateToCell(hit.SheetIndex, hit.Address.Row, hit.Address.Column); status.Text = hit.Location; }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        find.Click += (_, _) => Search(); replace.Click += (_, _) => Replace(false); all.Click += (_, _) => Replace(true);
        results.MouseDoubleClick += (_, _) => Locate();
        results.KeyDown += (_, args) => { if (args.Key == Key.Enter) { args.Handled = true; Locate(); } };
        dialog.Loaded += (_, _) =>
        {
            query.Focus();
            if (_smokeMode && _smokeSearchDriver is { } driver) Dispatcher.BeginInvoke(new Action(() => driver(dialog)));
        };
        dialog.ShowDialog(); QueueRecalculation();
    }
}
