using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private FilterPopup? _filterPopup;

    private void ClearSortState_Click(object sender, RoutedEventArgs e)
    {
        if (ExecuteWorkbookCommand(new ClearSortStateCommand(_document.Snapshot.Sheets[_sheetIndex].Id)))
            StatusText.Text = "已清除排序标记和上次排序设置，当前记录顺序保持不变，可撤销。";
    }

    private void EstablishDataRegion(GridRange range)
    {
        var region = new DataRegion(Guid.NewGuid(), _document.Snapshot.Sheets[_sheetIndex].Id, range);
        region.Validate(); if (!ExecuteWorkbookCommand(new SetAutoFilterCommand(region.SheetId, new(new(range.Top, range.Left, range.Bottom, range.Right))))) return; RefreshSheet();
        StatusText.Text = "筛选区域已建立。首行为标题；点击标题单元格右侧按钮或按 Alt+↓。";
    }

    private async void OpenHeaderFilter(int column, FrameworkElement target) => await (_workspaceInput = OpenHeaderFilterAsync(column, target));
    private async Task OpenHeaderFilterAsync(int column, FrameworkElement target)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var region = CurrentView.Region ?? throw new InvalidOperationException("请先选择含表头的矩形并建立筛选区域。");
            if (column < region.Range.Left || column > region.Range.Right) throw new ArgumentException("当前列不在筛选区域内。");
            var snapshot = _document.Snapshot; var view = CurrentView; var filters = view.Filters;
            var candidates = PersistentFiltering.Candidates(snapshot, snapshot.GetSheet(region.SheetId), column);
            _filterPopup?.Close();
            _filterPopup = new(target, new Rect(target.RenderSize), $"{ColumnName(column)} · {CellInput.DisplayText(snapshot.GetSheet(region.SheetId).GetCell(region.Range.Top, column)?.Value ?? CellValue.Blank)}", candidates, filters.GetValueOrDefault(column));
            var result = await _filterPopup.ShowAsync(); _filterPopup = null;
            if (result is null) return;
            if (!ReferenceEquals(snapshot, _document.Snapshot) || !ReferenceEquals(view, CurrentView) || view.Region != region || !ReferenceEquals(filters, view.Filters))
                throw new InvalidOperationException("面板打开期间数据或筛选已变化，请重新打开面板。");
            await ApplyFilterPopupResultAsync(column, result);
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async Task ApplyFilterPopupResultAsync(int column, FilterPopupResult result)
    {
        var region = CurrentView.Region ?? throw new InvalidOperationException("筛选区域已失效。");
        switch (result.Action)
        {
            case FilterPopupAction.SortAscending:
            case FilterPopupAction.SortDescending:
                await SortDataRegionAsync(region, [new(column, result.Action == FilterPopupAction.SortDescending)]); return;
            case FilterPopupAction.AdvancedSort: await AdvancedSortAsync(); return;
            default:
                var snapshot = _document.Snapshot; var old = snapshot.GetSheet(region.SheetId).Data.Filter ?? throw new InvalidOperationException("筛选区域已失效。");
                var columns = result.Action == FilterPopupAction.ClearAll ? old.Columns.Clear() : result.Action == FilterPopupAction.ClearColumn ? old.Columns.Remove(column) :
                    old.Columns.SetItem(column, PersistentFiltering.Convert(result.Filter ?? throw new ArgumentException("筛选条件缺失。"), snapshot.DateSystem)
                        .WithButton(old.Columns.GetValueOrDefault(column)?.ShowButton ?? true));
                if (!ExecuteWorkbookCommand(new SetAutoFilterCommand(region.SheetId, new(old.Range, columns)))) return;
                break;
        }
        RefreshSheet(); GridSurface.Select(new(region.Range.Top, column));
        StatusText.Text = "筛选已更新；区域外数据不参与判定，复制只含筛选可见行，粘贴不能跨筛选隐藏行。";
    }

    private async Task SortDataRegionAsync(DataRegion region, IReadOnlyList<SortKey> keys)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var range = region.Range;
        if (!ExecuteWorkbookCommand(new SortRangeCommand(new SortDescriptor(region.SheetId, new(range.Top, range.Left), new(range.Bottom, range.Right), keys, region.HasHeader)),
            new(range.Top, keys[0].Column))) return;
        // Sorting invalidates formula caches. Complete calculation before publishing a filtered projection.
        if (_calculationPending) await RecalculateAsync();
        RefreshSheet();
        StatusText.Text = CurrentView.FilterError ?? "排序完成（含筛选和手动隐藏记录）；筛选已重新应用，可一步撤销。";
    }

    private async void AdvancedSort_Click(object sender, RoutedEventArgs e)
    {
        try { await (_workspaceInput = AdvancedSortAsync()); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async Task AdvancedSortAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        var region = CurrentView.Region;
        if (region is null && CurrentView.FilterRange is not null) throw new InvalidOperationException("请清除筛选后建立明确的数据区域。");
        var range = region?.Range ?? SelectedRange();
        var snapshot = _document.Snapshot;
        var dialog = new SortDialog(this, snapshot, snapshot.Sheets[_sheetIndex].Id, range, CurrentView.SortKeys, region?.HasHeader ?? true, region is not null);
        if (dialog.ShowDialog() != true) return;
        if (!ReferenceEquals(snapshot, _document.Snapshot)) throw new InvalidOperationException("数据已变化，请重新发起排序。");
        await SortDataRegionAsync(region ?? new(Guid.NewGuid(), snapshot.Sheets[_sheetIndex].Id, range, dialog.HasHeader), dialog.Keys);
    }
}
