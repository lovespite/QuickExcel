using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using QuickExcel.Calculation;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private FormulaAuditWindow? _formulaAudit;
    private void InstallFormulaAudit()
    {
        var item = new MenuItem { Header = UiText.Get("公式审核…"), Name = "FormulaAuditMenu" };
        item.Bind(IsEnabledProperty, _book.GetObservable(IsVisibleProperty));
        item.Click += async (_, _) => { if (_book.IsVisible) await PerformAsync(ShowFormulaAuditAsync); };
        ((MenuItem)_menu.Items[3]!).Items.Add(item);
    }
    private async Task ShowFormulaAuditAsync()
    {
        if (!await CommitEditorsAsync()) return;
        if (_formulaAudit is { } open) { open.Activate(); await open.RefreshAsync(true); return; }
        var window = _formulaAudit = new FormulaAuditWindow(this);
        window.Closed += (_, _) => _formulaAudit = null;
        window.Show(this); await window.RefreshAsync(true);
    }

    private sealed class FormulaAuditWindow : Window
    {
        private readonly MainWindow _owner;
        private readonly TextBlock _status = new() { TextWrapping = TextWrapping.Wrap };
        private readonly TextBlock _notices = new() { TextWrapping = TextWrapping.Wrap, Foreground = Brushes.DarkOrange };
        private readonly TextBox _formula = new() { IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MaxHeight = 100 };
        private readonly ListBox _precedents = List("AuditPrecedents"), _dependents = List("AuditDependents");
        private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromMilliseconds(500) };
        private CancellationTokenSource? _running;
        private WorkbookSnapshot? _snapshot;
        private Guid _session;
        private CellAddress _origin;
        private bool _closed;
        internal FormulaAuditResult? Result { get; private set; }
        internal FormulaAuditWindow(MainWindow owner)
        {
            _owner = owner; Title = UiText.Get("公式审核"); Width = 850; Height = 640; MinWidth = 640; MinHeight = 460;
            WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
            var root = new DockPanel { Margin = new(18) }; var top = new StackPanel { Spacing = 8 }; DockPanel.SetDock(top, Dock.Top); root.Children.Add(top);
            top.Children.Add(new TextBlock { Text = UiText.Get("显示静态值引用和直接从属。位置/尺寸引用不读取值。审核固定地址，编辑或重算后自动刷新；切换审核位置请点击“使用当前格”。"), TextWrapping = TextWrapping.Wrap });
            top.Children.Add(_status); top.Children.Add(_formula); top.Children.Add(_notices);
            var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new(0, 8) }; top.Children.Add(actions);
            void Button(string label, Func<Task> action) { var button = new Button { Content = label }; button.Click += async (_, _) => await Guard(action); actions.Children.Add(button); }
            Button(UiText.Get("使用当前格"), () => RefreshAsync(true)); Button(UiText.Get("刷新"), () => RefreshAsync(false, true));
            Button(UiText.Get("取消分析"), () => { _running?.Cancel(); return Task.CompletedTask; });
            var columns = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,*") }; root.Children.Add(columns);
            void Column(string title, ListBox list, int column)
            {
                var panel = new DockPanel { Margin = new(column == 0 ? 0 : 8, 8, 0, 0) }; global::Avalonia.Controls.Grid.SetColumn(panel, column); columns.Children.Add(panel);
                var label = new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new(0, 0, 0, 6) }; DockPanel.SetDock(label, Dock.Top); panel.Children.Add(label);
                var locate = new Button { Content = UiText.Get("定位选中项"), Margin = new(0, 8, 0, 0) }; DockPanel.SetDock(locate, Dock.Bottom); panel.Children.Add(locate); panel.Children.Add(list);
                locate.Click += async (_, _) => await Guard(() => list.SelectedItem is FormulaAuditTarget target ? NavigateAsync(target) : Task.CompletedTask);
                list.DoubleTapped += async (_, _) => await Guard(() => list.SelectedItem is FormulaAuditTarget target ? NavigateAsync(target) : Task.CompletedTask);
            }
            Column(UiText.Get("直接值引用"), _precedents, 0); Column(UiText.Get("直接从属公式"), _dependents, 1);
            Content = root; _timer.Tick += async (_, _) => { if (_running is null) await Guard(() => RefreshAsync()); };
            _timer.Start(); Closed += (_, _) => { _closed = true; _timer.Stop(); _running?.Cancel(); };
        }
        private static ListBox List(string name)
        {
            var list = new ListBox { Name = name, ItemTemplate = new FuncDataTemplate<FormulaAuditTarget>((target, _) => target is null ? null :
                new TextBlock { Text = target.Label + "\n" + target.State + " · " + target.Value, TextWrapping = TextWrapping.Wrap, Margin = new(4) }) };
            AutomationProperties.SetName(list, name == "AuditPrecedents" ? UiText.Get("直接值引用") : UiText.Get("直接从属公式")); return list;
        }
        private async Task Guard(Func<Task> action)
        { try { await action(); } catch (OperationCanceledException) { _notices.Text = UiText.Get("分析已取消；未修改工作簿。"); } catch (Exception ex) { _notices.Text = ex.Message; } }
        internal async Task RefreshAsync(bool useCurrent = false, bool force = false)
        {
            if (_closed || _running is not null) return;
            if (!_owner._book.IsVisible) { Result = null; _precedents.ItemsSource = _dependents.ItemsSource = null; _status.Text = UiText.Get("没有打开的工作簿。"); return; }
            var snapshot = _owner.Document.Snapshot; var session = _owner.Document.Session.SessionId;
            if (useCurrent || session != _session || !snapshot.Sheets.Any(s => s.Id == _origin.SheetId))
                _origin = new(snapshot.ActiveSheetId, _owner._grid.Current.Row, _owner._grid.Current.Column);
            if (!force && !useCurrent && ReferenceEquals(snapshot, _snapshot)) return;
            using var cancellation = new CancellationTokenSource(); _running = cancellation; Result = null; _precedents.ItemsSource = _dependents.ItemsSource = null;
            _status.Text = UiText.Get("正在分析公式依赖…");
            try
            {
                var result = await Task.Run(() => FormulaAudit.Inspect(snapshot, _origin, cancellation.Token), cancellation.Token);
                if (_closed || session != _owner.Document.Session.SessionId || !ReferenceEquals(snapshot, _owner.Document.Snapshot)) return;
                _snapshot = snapshot; _session = session; Result = result;
                var current = FormulaAudit.Target(snapshot, _origin, new(_origin.Row, _origin.Column, _origin.Row, _origin.Column));
                _status.Text = current.Label + " · " + result.State + " · " + result.Value;
                _formula.Text = result.Formula; _notices.Text = string.Join(Environment.NewLine, result.Notices);
                _precedents.ItemsSource = result.Precedents; _dependents.ItemsSource = result.Dependents;
            }
            catch { _snapshot = snapshot; _session = session; throw; } // Retry on an explicit refresh or a changed snapshot, not every timer tick.
            finally { _running = null; }
        }
        internal async Task NavigateAsync(FormulaAuditTarget target)
        {
            if (!await _owner.CommitEditorsAsync()) return;
            if (Result is null || _session != _owner.Document.Session.SessionId || !ReferenceEquals(_snapshot, _owner.Document.Snapshot))
            { await RefreshAsync(false, true); throw new InvalidOperationException(UiText.Get("审核结果已改变，请在刷新后的列表中重新选择。")); }
            await _owner.NavigateToAsync(target.Address.SheetId, new(target.Range.Top, target.Range.Left, target.Range.Bottom, target.Range.Right));
        }
    }
}
