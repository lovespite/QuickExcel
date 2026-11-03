using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallSearch()
    {
        RibbonGroup("开始", "查找", RibbonButton("查找与替换", SearchAsync, false), RibbonButton("定位", NavigateAsync, false));
        var edit = (MenuItem)_menu.Items[1]!; edit.Items.Add(CommandMenu("查找与替换  Ctrl+F", SearchAsync)); edit.Items.Add(CommandMenu("定位  Ctrl+G", NavigateAsync));
    }
    private async Task NavigateAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var text = await Dialogs.Text(this, "定位", "单元格、区域或名称（例如 A1、B2:D10、'工作表 2'!A1）", FormulaEditing.Address(SelectedRange()));
        if (text is null) return;
        var expression = FormulaSyntax.Parse("=" + text.TrimStart('='));
        expression = FormulaBinding.Bind(expression, Document.Snapshot, Document.Snapshot.ActiveSheetId, _grid.Current);
        if (expression is not ReferenceNode reference || reference.WholeColumns) throw new ArgumentException("请输入单元格或矩形区域引用。");
        var sheet = reference.Sheet is null ? Document.Snapshot.Sheets[SheetIndex] : Document.Snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase)) ?? throw new ArgumentException("工作表不存在。");
        await NavigateToAsync(sheet.Id, new(reference.Row, reference.Column, reference.EndRow, reference.EndColumn));
    }
    private async Task NavigateToAsync(SheetId sheet, GridRange range)
    {
        if (!await CommitEditorsAsync()) return;
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        var snapshot = Document.Snapshot; var target = snapshot.GetSheet(sheet);
        if (target.Visibility != SheetVisibility.Visible) throw new InvalidOperationException("不能定位到隐藏工作表。");
        if (target.Rows.GetValueOrDefault(range.Top)?.Hidden == true || target.Columns.GetValueOrDefault(range.Left)?.Hidden == true ||
            WorksheetFilterEngine.Evaluate(snapshot, sheet, WorksheetFilterEngine.CanUseSourceCaches(snapshot)).Contains(range.Top))
            throw new InvalidOperationException("目标单元格位于隐藏或筛除的行列，请先取消隐藏或筛选。");
        if (!ProtectionSafety.CanSelect(snapshot, target, range.Top, range.Left)) throw new InvalidOperationException("工作表保护禁止选择该单元格。");
        Document.Session.UpdateNavigation(sheet, new(range.Top, range.Left)); Synchronize();
        _grid.Select(new(range.Top, range.Left)); _grid.SelectRange(range, new(range.Top, range.Left));
    }
    private async Task SearchAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var selection = _grid.SelectionRanges.ToArray(); var selectedSheet = Document.Snapshot.ActiveSheetId; var activeSheet = SheetIndex;
        if (selection.Length > SelectionStatistics.MaxRanges) throw new ArgumentException("请将选区数量缩小到 64 个以内。");
        var window = new Window { Title = "查找与替换", Width = 660, Height = 620, MinWidth = 500, MinHeight = 400,
            WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var panel = new DockPanel { Margin = new(20) }; var fields = new StackPanel { Spacing = 8 }; DockPanel.SetDock(fields, Dock.Top); panel.Children.Add(fields);
        TextBox Field(string label)
        { fields.Children.Add(new TextBlock { Text = label }); var box = new TextBox(); AutomationProperties.SetName(box, label); fields.Children.Add(box); return box; }
        var query = Field("查找"); var replacement = Field("替换为");
        var scope = new ComboBox { ItemsSource = new[] { "当前工作表", "整个工作簿", "选区" }, SelectedIndex = 0, MinWidth = 140 };
        var target = new ComboBox { ItemsSource = new[] { "显示值", "公式源码" }, SelectedIndex = 0, MinWidth = 120 };
        AutomationProperties.SetName(scope, "查找范围"); AutomationProperties.SetName(target, "查找对象");
        var selectors = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 }; selectors.Children.Add(scope); selectors.Children.Add(target); fields.Children.Add(selectors);
        var options = new WrapPanel();
        var matchCase = new CheckBox { Content = "区分大小写", Margin = new(0, 0, 12, 0) }; var whole = new CheckBox { Content = "整格匹配", Margin = new(0, 0, 12, 0) }; var hidden = new CheckBox { Content = "包含隐藏行列" };
        options.Children.Add(matchCase); options.Children.Add(whole); options.Children.Add(hidden); fields.Children.Add(options);
        var actions = new WrapPanel(); fields.Children.Add(actions);
        var status = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 8) }; fields.Children.Add(status);
        AutomationProperties.SetLiveSetting(status, AutomationLiveSetting.Polite);
        var results = new ListBox { ItemTemplate = new FuncDataTemplate<SearchHit>((hit, _) => new TextBlock { Text = hit?.Location + "   " + hit?.Text, TextTrimming = TextTrimming.CharacterEllipsis }) };
        AutomationProperties.SetName(results, "查找结果，Enter 定位"); panel.Children.Add(results); window.Content = panel;
        SearchOptions Options() => new(query.Text ?? "", (SearchScope)scope.SelectedIndex, (SearchTarget)target.SelectedIndex, matchCase.IsChecked == true, whole.IsChecked == true, hidden.IsChecked == true);
        bool Selected(CellAddress address) => address.SheetId == selectedSheet && selection.Any(r => address.Row >= r.Top && address.Row <= r.Bottom && address.Column >= r.Left && address.Column <= r.Right);
        async Task<(IReadOnlyList<SearchHit> Hits, long Revision, SearchOptions Options)> FindAsync()
        {
            var snapshot = Document.Snapshot; var selectedOptions = Options();
            var hits = await Task.Run(() =>
            {
                var filters = snapshot.Sheets.Select(s => WorksheetFilterEngine.Evaluate(snapshot, s.Id, WorksheetFilterEngine.CanUseSourceCaches(snapshot))).ToArray();
                return WorkbookSearch.Find(snapshot, activeSheet, selectedOptions, filtered: (s, r) => filters[s].Contains(r), selected: Selected);
            });
            return (hits, snapshot.Revision, selectedOptions);
        }
        void Action(string title, Func<Task> action)
        {
            var button = new Button { Content = title, Margin = new(0, 0, 8, 0) }; actions.Children.Add(button);
            button.Click += async (_, _) =>
            {
                actions.IsEnabled = false;
                try { await action(); } catch (Exception ex) { status.Text = ex.Message; }
                finally { actions.IsEnabled = true; }
            };
        }
        Action("查找全部", async () => { var found = await FindAsync(); results.ItemsSource = found.Hits; status.Text = $"{found.Hits.Count:N0} 个匹配"; });
        async Task ReplaceAsync(bool all)
        {
            var found = await FindAsync(); var hits = all ? found.Hits : found.Hits.Where(h => results.SelectedItem is SearchHit selected && h.Address == selected.Address).ToArray();
            if (hits.Count == 0) { status.Text = "没有可替换的匹配。"; return; }
            if (await ExecuteAsync(WorkbookSearch.Replace(hits, found.Options, replacement.Text ?? ""), found.Revision))
            { results.ItemsSource = null; status.Text = $"已替换 {hits.Count:N0} 格，可一次撤销。"; }
        }
        Action("替换所选", () => ReplaceAsync(false)); Action("全部替换", () => ReplaceAsync(true));
        async Task LocateAsync()
        {
            if (results.SelectedItem is not SearchHit hit) return;
            try { await NavigateToAsync(hit.Address.SheetId, new(hit.Address.Row, hit.Address.Column, hit.Address.Row, hit.Address.Column)); status.Text = hit.Location; }
            catch (Exception ex) { status.Text = ex.Message; }
        }
        results.DoubleTapped += async (_, _) => await LocateAsync();
        results.KeyDown += async (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; await LocateAsync(); } };
        Action("定位", LocateAsync); Action("关闭", () => { window.Close(); return Task.CompletedTask; });
        window.Opened += (_, _) => query.Focus(); await window.ShowDialog(this);
    }
}
