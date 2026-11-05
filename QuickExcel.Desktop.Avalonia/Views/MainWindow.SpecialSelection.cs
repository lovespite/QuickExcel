using Avalonia;
using Avalonia.Controls;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallSpecialSelection()
    {
        var item = new MenuItem { Header = UiText.Get("定位条件…") };
        item.Bind(IsEnabledProperty, _book.GetObservable(IsVisibleProperty));
        item.Click += async (_, _) => { if (_book.IsVisible) await PerformAsync(GoToSpecialAsync); };
        ((MenuItem)_menu.Items[1]!).Items.Add(item);
    }

    private async Task GoToSpecialAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(snapshot.ActiveSheetId);
        var session = Document.Session.SessionId; var selection = _grid.SelectionRanges.ToArray();
        var dialog = new FieldsDialog(UiText.Get("定位条件"))
            .Note(UiText.Get("选择不会修改工作簿。空白不包含空文本或公式结果；合并格按左上角判断，仅选择范围内完整且允许选择的合并格。可见单元格排除隐藏和筛除行列。结果最多 64 个连续区域；大范围请缩小后重试。"))
            .Choice("kind", UiText.Get("定位对象"), [UiText.Get("空白"), UiText.Get("常量"), UiText.Get("公式"), UiText.Get("错误值"), UiText.Get("可见单元格")])
            .Choice("scope", UiText.Get("定位范围"), [UiText.Get("当前选区"), UiText.Get("当前表已用范围")]);
        if (!await dialog.ShowAsync(this)) return;
        var kind = (SpecialCellKind)dialog.Index("kind");
        CellRectangle[] ranges = dialog.Index("scope") == 0 ? selection.Select(Rectangle).ToArray()
            : [new(0, 0, Math.Max(0, sheet.UsedRowCount - 1), Math.Max(0, sheet.UsedColumnCount - 1))];
        using var cancellation = new CancellationTokenSource();
        SpecialCellSelectionResult result;
        using (BeginExecution(UiText.Get("正在定位单元格…"), cancellation))
            result = await Task.Run(() => SpecialCellSelection.Find(snapshot, sheet.Id, ranges, kind, cancellation.Token), cancellation.Token);
        void RequireCurrent()
        {
            if (Document.Session.SessionId != session || !ReferenceEquals(Document.Snapshot.GetSheet(sheet.Id), sheet) ||
                Document.Snapshot.Revision != snapshot.Revision || Document.Snapshot.ActiveSheetId != sheet.Id || !_grid.SelectionRanges.SequenceEqual(selection))
                throw new InvalidOperationException(UiText.Get("文档或选区已改变，请重新定位。"));
        }
        RequireCurrent();
        if (result.Cells == 0) { _status.Text = UiText.Get("没有符合定位条件的单元格。"); return; }
        var confirmation = new FieldsDialog(UiText.Get("确认定位"))
            .Note(UiText.Format($"将选择 {result.Cells:N0} 格，共 {result.Ranges.Count:N0} 个连续区域。可继续查看统计、设置格式或清除内容。批量编辑仍遵守现有容量及保护限制。"));
        if (!await confirmation.ShowAsync(this)) return;
        RequireCurrent();
        var first = result.Ranges[0];
        Document.Session.UpdateNavigation(sheet.Id, new(first.Top, first.Left)); Synchronize();
        _grid.SelectRanges(result.Ranges.Select(r => new GridRange(r.Top, r.Left, r.Bottom, r.Right)), new(first.Top, first.Left));
        _status.Text = UiText.Format($"已定位 {result.Cells:N0} 格，{result.Ranges.Count:N0} 个区域。");
    }
}
