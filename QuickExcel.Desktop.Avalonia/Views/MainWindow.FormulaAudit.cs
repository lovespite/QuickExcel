using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;

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
        private readonly ListBox _errors = List("AuditErrors");
        private readonly TextBlock _step = new() { TextWrapping = TextWrapping.Wrap };
        private FormulaEvaluationResult? _evaluation;
        private int _stepIndex;
        private int _arrowMode = 1;
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
            var more = new WrapPanel { Orientation = Orientation.Horizontal };
            top.Children.Add(more);
            void Tool(string label, Action action) { var button = new Button { Content = label, Margin = new(0, 0, 8, 8) }; button.Click += (_, _) => action(); more.Children.Add(button); }
            Tool(UiText.Get("引用箭头"), () => { _arrowMode = 1; UpdateArrows(); });
            Tool(UiText.Get("从属箭头"), () => { _arrowMode = 2; UpdateArrows(); });
            Tool(UiText.Get("清除箭头"), () => { _arrowMode = 0; UpdateArrows(); });
            Tool(UiText.Get("下一步求值"), NextStep);
            Tool(UiText.Get("重新查看步骤"), () => { _stepIndex = 0; NextStep(); });
            top.Children.Add(new TextBlock { Text = UiText.Get("箭头仅连接当前视图可见的同表端点（最多 50 条）；跨表和视图外目标请用列表定位。错误列表最多显示 500 项，排除旧缓存。"), TextWrapping = TextWrapping.Wrap });
            top.Children.Add(_step);
            var columns = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,*,*") }; root.Children.Add(columns);
            void Column(string title, ListBox list, int column)
            {
                var panel = new DockPanel { Margin = new(column == 0 ? 0 : 8, 8, 0, 0) }; global::Avalonia.Controls.Grid.SetColumn(panel, column); columns.Children.Add(panel);
                var label = new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new(0, 0, 0, 6) }; DockPanel.SetDock(label, Dock.Top); panel.Children.Add(label);
                var locate = new Button { Content = UiText.Get("定位选中项"), Margin = new(0, 8, 0, 0) }; DockPanel.SetDock(locate, Dock.Bottom); panel.Children.Add(locate); panel.Children.Add(list);
                locate.Click += async (_, _) => await Guard(() => list.SelectedItem is FormulaAuditTarget target ? NavigateAsync(target) : Task.CompletedTask);
                list.DoubleTapped += async (_, _) => await Guard(() => list.SelectedItem is FormulaAuditTarget target ? NavigateAsync(target) : Task.CompletedTask);
            }
            Column(UiText.Get("直接值引用"), _precedents, 0); Column(UiText.Get("直接从属公式"), _dependents, 1);
            Column(UiText.Get("当前错误单元格"), _errors, 2);
            Content = root; _timer.Tick += async (_, _) => { if (_running is null) await Guard(() => RefreshAsync()); };
            _timer.Start(); Closed += (_, _) => { _closed = true; _timer.Stop(); _running?.Cancel(); _owner._grid.SetFormulaAuditArrows(null); };
        }
        private static ListBox List(string name)
        {
            var list = new ListBox { Name = name, ItemTemplate = new FuncDataTemplate<FormulaAuditTarget>((target, _) => target is null ? null :
                new TextBlock { Text = target.Label + "\n" + target.State + " · " + target.Value, TextWrapping = TextWrapping.Wrap, Margin = new(4) }) };
            AutomationProperties.SetName(list, name == "AuditPrecedents" ? UiText.Get("直接值引用") : name == "AuditErrors" ? UiText.Get("当前错误单元格") : UiText.Get("直接从属公式")); return list;
        }
        internal void NextStep()
        {
            if (!ReferenceEquals(_snapshot, _owner.Document.Snapshot) || _session != _owner.Document.Session.SessionId)
            { _step.Text = UiText.Get("审核结果已改变，请在刷新后的列表中重新选择。"); return; }
            if (_evaluation is null) return;
            if (_evaluation.Notice is { } notice) { _step.Text = notice; return; }
            if (_stepIndex >= _evaluation.Steps.Count) return;
            var step = _evaluation.Steps[_stepIndex++];
            _step.Text = $"{_stepIndex}/{_evaluation.Steps.Count}  {step.Expression} → {step.Value}";
        }
        private void UpdateArrows()
        {
            if (Result is null || !ReferenceEquals(_snapshot, _owner.Document.Snapshot)) { _owner._grid.SetFormulaAuditArrows(null); return; }
            _owner._grid.Refresh(_snapshot!);
            _owner._grid.SetFormulaAuditArrows(_snapshot, _arrowMode switch
            {
                1 => Result.Precedents.Select(t => new FormulaAuditArrow(t.Address, Result.Origin)),
                2 => Result.Dependents.Select(t => new FormulaAuditArrow(Result.Origin, t.Address)),
                _ => []
            });
        }
        private async Task Guard(Func<Task> action)
        { try { await action(); } catch (OperationCanceledException) { _notices.Text = UiText.Get("分析已取消；未修改工作簿。"); } catch (Exception ex) { _notices.Text = ex.Message; } }
        internal async Task RefreshAsync(bool useCurrent = false, bool force = false)
        {
            if (_closed || _running is not null) return;
            if (!_owner._book.IsVisible) { Result = null; _evaluation = null; _step.Text = _formula.Text = _notices.Text = ""; _precedents.ItemsSource = _dependents.ItemsSource = _errors.ItemsSource = null; _owner._grid.SetFormulaAuditArrows(null); _status.Text = UiText.Get("没有打开的工作簿。"); return; }
            var snapshot = _owner.Document.Snapshot; var session = _owner.Document.Session.SessionId;
            if (useCurrent || session != _session || !snapshot.Sheets.Any(s => s.Id == _origin.SheetId))
                _origin = new(snapshot.ActiveSheetId, _owner._grid.Current.Row, _owner._grid.Current.Column);
            if (!force && !useCurrent && ReferenceEquals(snapshot, _snapshot)) return;
            using var cancellation = new CancellationTokenSource(); _running = cancellation; Result = null; _evaluation = null; _step.Text = _formula.Text = _notices.Text = ""; _precedents.ItemsSource = _dependents.ItemsSource = _errors.ItemsSource = null; _owner._grid.SetFormulaAuditArrows(null);
            _status.Text = UiText.Get("正在分析公式依赖…");
            try
            {
                var bundle = await Task.Run(() => (Audit: FormulaAudit.Inspect(snapshot, _origin, cancellation.Token),
                    Evaluation: FormulaEvaluation.Inspect(snapshot, _origin, cancellation.Token), Errors: FormulaEvaluation.Errors(snapshot, cancellation.Token)), cancellation.Token);
                var result = bundle.Audit;
                if (_closed || session != _owner.Document.Session.SessionId || !ReferenceEquals(snapshot, _owner.Document.Snapshot)) return;
                _snapshot = snapshot; _session = session; Result = result;
                var current = FormulaAudit.Target(snapshot, _origin, new(_origin.Row, _origin.Column, _origin.Row, _origin.Column));
                _status.Text = current.Label + " · " + result.State + " · " + result.Value;
                _formula.Text = result.Formula; _notices.Text = string.Join(Environment.NewLine, result.Notices);
                _precedents.ItemsSource = result.Precedents; _dependents.ItemsSource = result.Dependents;
                _errors.ItemsSource = bundle.Errors; _evaluation = bundle.Evaluation; _stepIndex = 0;
                _step.Text = _evaluation.Notice ?? UiText.Get("点击“下一步求值”按实际执行顺序查看子表达式；未选择的惰性分支不执行。");
                UpdateArrows();
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
