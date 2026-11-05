using Avalonia.Automation;
using Avalonia.Automation.Peers;
using Avalonia.Automation.Provider;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    protected override AutomationPeer OnCreateAutomationPeer() => new GridPeer(this);
    // One current-cell value provider plus the native editor and scroll bars; never one peer per stored cell.
    private sealed class GridPeer : ControlAutomationPeer, IValueProvider
    {
        private readonly WorkbookGrid _grid;
        public GridPeer(WorkbookGrid grid) : base(grid)
        {
            _grid = grid;
            grid.SelectionChanged += () => RaisePropertyChangedEvent(ValuePatternIdentifiers.ValueProperty, null, Value);
            grid.DraftEnded += () => RaisePropertyChangedEvent(ValuePatternIdentifiers.ValueProperty, null, Value);
        }
        public bool IsReadOnly => !_grid.IsEffectivelyEnabled || _grid.IsCellReadOnly(_grid.Current);
        public string? Value => _grid._snapshot is null ? "" :
            ProtectionSafety.HidesFormula(_grid._snapshot, _grid.Sheet, _grid.Current.Row, _grid.Current.Column) ? "受保护的隐藏公式" :
            _grid.Sheet.GetCell(_grid.Current.Row, _grid.Current.Column)?.Value.ToString() ?? "";
        public void SetValue(string? value)
        {
            if (IsReadOnly || !_grid.BeginEdit()) throw new InvalidOperationException("当前单元格不能编辑。");
            // Match a native TextBox: editing owns a draft; Enter/Tab or a host command commits it.
            _grid.CellEditor.Text = value ?? "";
        }
        protected override AutomationControlType GetAutomationControlTypeCore() => AutomationControlType.DataGrid;
        protected override IReadOnlyList<AutomationPeer> GetChildrenCore() =>
            [CreatePeerForElement(_grid.CellEditor), CreatePeerForElement(_grid._vertical), CreatePeerForElement(_grid._horizontal)];
        protected override string GetNameCore() => _grid._snapshot is null ? "工作表网格" :
            $"工作表 {_grid.Sheet.Name}，{ColumnName(_grid.Current.Column)} 列，第 {_grid.Current.Row + 1} 行，{(IsReadOnly ? "只读" : "可编辑")}";
        protected override string GetHelpTextCore() => "选区：" + string.Join("；", _grid.SelectionRanges.Take(64).Select(r =>
            $"{ColumnName(r.Left)}{r.Top + 1}:{ColumnName(r.Right)}{r.Bottom + 1}")) +
            "。当前值：" + Value + "。F2 编辑，Alt+Enter 换行，Enter 提交，Escape 取消，F6 切换网格、公式栏、工作表和菜单。";
        protected override void SetFocusCore() => _grid.CellEditor.Focus();
        protected override bool HasKeyboardFocusCore() => _grid.IsKeyboardFocusWithin;
    }
}
