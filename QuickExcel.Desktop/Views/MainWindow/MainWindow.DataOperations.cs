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
        if (ExecuteWorkbookCommand(new ClearSortStateCommand(_document.Snapshot.Sheets[_sheetIndex].Id, CurrentView.TableId)))
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
            var candidates = PersistentFiltering.Candidates(snapshot, snapshot.GetSheet(region.SheetId), column, view.TableId);
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
                await SortDataRegionAsync(region, [new(column, result.Action == FilterPopupAction.SortDescending)], expandColumns: CurrentView.TableId is null); return;
            case FilterPopupAction.AdvancedSort: await AdvancedSortAsync(); return;
            default:
                var snapshot = _document.Snapshot; var data = snapshot.GetSheet(region.SheetId).Data; var tableId = CurrentView.TableId;
                var old = (tableId.HasValue ? data.Tables.Single(t => t.Id == tableId).Filter : data.Filter) ?? throw new InvalidOperationException("筛选区域已失效。");
                var columns = result.Action == FilterPopupAction.ClearAll ? old.Columns.Clear() : result.Action == FilterPopupAction.ClearColumn ? old.Columns.Remove(column) :
                    old.Columns.SetItem(column, PersistentFiltering.Convert(result.Filter ?? throw new ArgumentException("筛选条件缺失。"), snapshot.DateSystem)
                        .WithButton(old.Columns.GetValueOrDefault(column)?.ShowButton ?? true));
                if (!ExecuteWorkbookCommand(tableId is { } id ? new SetTableFilterCommand(region.SheetId, id, new(old.Range, columns)) : new SetAutoFilterCommand(region.SheetId, new(old.Range, columns)))) return;
                break;
        }
        RefreshSheet(); GridSurface.Select(new(region.Range.Top, column));
        StatusText.Text = "筛选已更新；区域外数据不参与判定，复制只含筛选可见行，粘贴不能跨筛选隐藏行。";
    }

    private async Task SortDataRegionAsync(DataRegion region, IReadOnlyList<SortKey> keys, bool expandColumns = false)
    {
        if (_busy) { StatusText.Text = "正在处理其他操作，请稍后再排序。"; return; }
        if (!await FlushEditsAsync()) { StatusText.Text = "排序未开始：当前编辑没有提交成功。"; return; }
        var range = region.Range;
        var snapshot = _document.Snapshot;
        var sheet = snapshot.GetSheet(region.SheetId);
        // The key columns must hold current results; sorting otherwise rejects half-way with a cache error.
        if (SortKeysNeedCalculation(snapshot, sheet, region, keys))
        {
            if (!await RecalculateAsync() || SortKeysNeedCalculation(_document.Snapshot, _document.Snapshot.GetSheet(region.SheetId), region, keys))
                throw new InvalidOperationException("排序关键列包含待重算公式，重算未完成；排序已取消，工作簿未修改。");
        }
        if (!ExecuteWorkbookCommand(new SortRangeCommand(new SortDescriptor(region.SheetId, new(range.Top, range.Left), new(range.Bottom, range.Right),
                keys, region.HasHeader, expandColumns)), new(range.Top, keys[0].Column))) return;
        // Sorting invalidates formula caches. Complete calculation before publishing a filtered projection.
        if (_calculationPending) await RecalculateAsync();
        RefreshSheet();
        var data = _document.Snapshot.GetSheet(region.SheetId).Data;
        var table = CurrentView.TableId is { } tableId ? data.Tables.FirstOrDefault(t => t.Id == tableId) : null;
        var applied = (table is null ? data.Sort : table.Sort)?.Range;
        var where = applied is { } sorted ? $"{ColumnName(sorted.Left)}{range.Top + 1}:{ColumnName(sorted.Right)}{range.Bottom + 1}" : ColumnName(keys[0].Column) + " 列";
        StatusText.Text = CurrentView.FilterError ?? $"排序完成：{where}{(region.HasHeader ? "（首行为标题）" : "")}{(expandColumns ? "，相邻数据列同步移动" : "")}；筛选已重新应用，可一步撤销。";
    }

    private static bool SortKeysNeedCalculation(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, DataRegion region, IReadOnlyList<SortKey> keys)
    {
        for (var row = region.FirstDataRow; row <= region.Range.Bottom; row++)
            foreach (var key in keys)
            {
                var cell = sheet.GetCell(row, key.Column);
                if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation?.Value is null ||
                    cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.InputRevision != snapshot.Revision))
                    return true;
            }
        return false;
    }

    private async void AdvancedSort_Click(object sender, RoutedEventArgs e)
    {
        try { await (_workspaceInput = AdvancedSortAsync()); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async Task AdvancedSortAsync()
    {
        if (_busy) { StatusText.Text = "正在处理其他操作，请稍后再排序。"; return; }
        if (!await FlushEditsAsync()) { StatusText.Text = "排序未开始：当前编辑没有提交成功。"; return; }
        var region = CurrentView.Region ?? TableSortRegion();
        if (region is null && CurrentView.FilterRange is not null) throw new InvalidOperationException("请清除筛选后建立明确的数据区域。");
        var snapshot = _document.Snapshot;
        var sheet = snapshot.Sheets[_sheetIndex];
        var block = region is null ? SortPlanning.Resolve(sheet, SelectionBounds(), new(_selectedRow, _selectedColumn)) : null;
        var range = region?.Range ?? new GridRange(block!.Range.Top, block.Range.Left, block.Range.Bottom, block.Range.Right);
        if (range.Rows < 2) throw new ArgumentException("当前位置没有可排序的多行数据；请点击数据区内的单元格或选出包含记录的矩形。");
        var table = CurrentView.TableId is not null;
        var dialog = new SortDialog(this, snapshot, sheet.Id, range, CurrentView.SortKeys, region?.HasHeader ?? block!.HasHeader,
            region is not null, region is not null && !table,
            table ? "表内排序范围固定" : region is null ? "已包含相邻数据列" : null);
        if (dialog.ShowDialog() != true) return;
        if (!ReferenceEquals(snapshot, _document.Snapshot)) throw new InvalidOperationException("数据已变化，请重新发起排序。");
        // An auto-detected block is already as wide as its records; a persisted filter region may still grow sideways.
        await SortDataRegionAsync(region ?? new(Guid.NewGuid(), sheet.Id, range, dialog.HasHeader), dialog.Keys, dialog.ExpandColumns);
    }

    private CellRectangle? SelectionBounds()
    {
        var ranges = GridSurface.SelectionRanges;
        if (ranges.Count == 0) return null;
        return new(ranges.Min(r => r.Top), ranges.Min(r => r.Left), ranges.Max(r => r.Bottom), ranges.Max(r => r.Right));
    }

    /// <summary>A table without a filter still owns an exact sortable rectangle (header and totals rows excluded).</summary>
    private DataRegion? TableSortRegion()
    {
        if (CurrentView.TableId is not { } id) return null;
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var table = sheet.Data.Tables.FirstOrDefault(t => t.Id == id);
        if (table is null) return null;
        var range = new GridRange(table.Range.Top, table.Range.Left, table.Range.Bottom - (table.TotalsRow ? 1 : 0), table.Range.Right);
        if (range.Rows < 2) return null;
        return new DataRegion(Guid.NewGuid(), sheet.Id, range, table.HeaderRow);
    }
}
