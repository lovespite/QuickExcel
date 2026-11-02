using System.Collections.Immutable;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallDataCommands()
    {
        RibbonGroup("数据", "排序和筛选", RibbonButton("升序", () => SortAsync(false)), RibbonButton("降序", () => SortAsync(true)),
            RibbonButton("多级排序…", () => SortAsync(null)), RibbonButton("筛选区域", EstablishFilterAsync), RibbonButton("当前列筛选…", () => FilterAsync(_grid.Current.Column)),
            RibbonButton("清除全部筛选条件", ClearFiltersAsync), RibbonButton("取消筛选区域", RemoveFilterAsync), RibbonButton("清除排序标记", ClearSortAsync));
        RibbonGroup("插入", "表格", RibbonButton("创建表…", CreateTableAsync), RibbonButton("调整表范围…", ResizeTableAsync), RibbonButton("重命名表…", () => EditTableAsync(0)),
            RibbonButton("重命名表列…", () => EditTableAsync(1)), RibbonButton("设置计算列…", () => EditTableAsync(2)));
        _grid.FilterRequested += column => _ = PerformAsync(() => FilterAsync(column));
        AddHandler(KeyDownEvent, async (_, e) =>
        { if (!e.Handled && e.Key == Key.Down && e.KeyModifiers == KeyModifiers.Alt && !_grid.IsEditing && !_formula.IsComposing) { e.Handled = true; await PerformAsync(() => FilterAsync(_grid.Current.Column)); } }, RoutingStrategies.Tunnel);
    }
    private WorksheetTable? CurrentTable => Document.Snapshot.Sheets[SheetIndex].Data.Tables.FirstOrDefault(t => t.Range.Contains(_grid.Current.Row, _grid.Current.Column));
    private async Task SortAsync(bool? descending)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var table = CurrentTable; var filter = table?.Filter ?? (table is null ? sheet.Data.Filter : null);
        var block = table is not null ? new SortBlock(table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) }, table.HeaderRow) :
            filter is not null ? new SortBlock(filter.Range, true) : SortPlanning.Resolve(sheet, Rectangle(SelectedRange()), _grid.Current);
        if (block.Range.Rows < 2) throw new ArgumentException("请选择包含多行记录的数据区域。");
        SortDescriptor descriptor;
        if (descending is { } order) descriptor = new(sheet.Id, new(block.Range.Top, block.Range.Left), new(block.Range.Bottom, block.Range.Right),
            [new(Math.Clamp(_grid.Current.Column, block.Range.Left, block.Range.Right), order)], block.HasHeader, filter is not null && table is null);
        else
        {
            var dialog = new SortDialog(snapshot, sheet, block.Range, block.HasHeader, CurrentView.SortKeys, table is not null || filter is not null, filter is not null && table is null);
            if (!await ShowOwnedAsync<bool>(dialog) || dialog.Descriptor is null) return; descriptor = dialog.Descriptor;
        }
        if (await ExecuteAsync(new SortRangeCommand(descriptor), snapshot.Revision)) await CalculateAsync();
    }
    private async Task EstablishFilterAsync()
    {
        if (!await FlushEditsAsync()) return;
        if (CurrentTable is not null) throw new InvalidOperationException("表格已有独立筛选；请使用当前列筛选。");
        var range = SelectedRange(); if (range.Count == 1) { var block = SortPlanning.CurrentRegion(Document.Snapshot.Sheets[SheetIndex], range.Top, range.Left); range = new(block.Top, block.Left, block.Bottom, block.Right); }
        new DataRegion(Guid.NewGuid(), Document.Snapshot.ActiveSheetId, range).Validate();
        await ExecuteAsync(new SetAutoFilterCommand(Document.Snapshot.ActiveSheetId, new(Rectangle(range))));
    }
    private async Task FilterAsync(int column)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var table = CurrentTable;
        var filter = table is null ? sheet.Data.Filter : table.Filter;
        if (filter is null || column < filter.Range.Left || column > filter.Range.Right) throw new InvalidOperationException("请先建立筛选区域，并选择区域内的一列。");
        var candidates = await Task.Run(() => PersistentFiltering.Candidates(snapshot, sheet, column, table?.Id));
        var dialog = new FilterDialog(WorkbookGrid.ColumnName(column) + " 列筛选", candidates, filter.Columns.TryGetValue(column, out var current) ? PersistentFiltering.Restore(current) : null);
        var result = await ShowOwnedAsync<FilterChoice?>(dialog); if (result is null) return;
        if (snapshot.Revision != Document.Snapshot.Revision) throw new InvalidOperationException("筛选面板打开期间文档已改变，请重新打开。");
        if (result.Action is "ascending" or "descending")
        {
            var range = filter.Range;
            if (await ExecuteAsync(new SortRangeCommand(new(sheet.Id, new(range.Top, range.Left), new(range.Bottom, range.Right), [new(column, result.Action == "descending")], true, table is null)), snapshot.Revision)) await CalculateAsync();
            return;
        }
        var columns = result.Action == "clear-all" ? filter.Columns.Clear() : result.Action == "clear" ? filter.Columns.Remove(column) :
            filter.Columns.SetItem(column, PersistentFiltering.Convert(result.Filter!, snapshot.DateSystem).WithButton(current?.ShowButton ?? true));
        await ExecuteAsync(table is null ? new SetAutoFilterCommand(sheet.Id, new(filter.Range, columns)) : new SetTableFilterCommand(sheet.Id, table.Id, new(filter.Range, columns)), snapshot.Revision);
        await CalculatePendingAsync(); _grid.Select(new(filter.Range.Top, column));
    }
    private async Task ClearFiltersAsync()
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable; var sheet = Document.Snapshot.Sheets[SheetIndex]; var filter = table is null ? sheet.Data.Filter : table.Filter;
        if (filter is null) return;
        await ExecuteAsync(table is null ? new SetAutoFilterCommand(sheet.Id, new(filter.Range)) : new SetTableFilterCommand(sheet.Id, table.Id, new(filter.Range)));
    }
    private async Task RemoveFilterAsync()
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable; var sheet = Document.Snapshot.Sheets[SheetIndex];
        await ExecuteAsync(table is null ? new SetAutoFilterCommand(sheet.Id, null) : new SetTableFilterCommand(sheet.Id, table.Id, null));
    }
    private async Task ClearSortAsync()
    { if (await CommitEditorsAsync()) await ExecuteAsync(new ClearSortStateCommand(Document.Snapshot.ActiveSheetId, CurrentTable?.Id)); }
    private async Task CreateTableAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = SelectedRange(); var suffix = 1;
        while (snapshot.Sheets.SelectMany(s => s.Data.Tables).Any(t => t.Name.Equals("QE_Table" + suffix, StringComparison.OrdinalIgnoreCase))) suffix++;
        var dialog = new FieldsDialog("创建表").Note("所选区域首行须为非空、唯一的文本表头。").Text("name", "表名", "QE_Table" + suffix)
            .Choice("style", "表样式", ["蓝色", "绿色", "橙色", "灰色"]);
        if (!await dialog.ShowAsync(this)) return;
        var style = TableStyles.Resolve(snapshot, new[] { "TableStyleMedium2", "TableStyleMedium4", "TableStyleMedium9", "TableStyleMedium1" }[dialog.Index("style")]);
        await ExecuteAsync(new CreateTableCommand(snapshot.ActiveSheetId, Rectangle(range), dialog.String("name"), style), snapshot.Revision);
    }
    private async Task ResizeTableAsync()
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable ?? throw new InvalidOperationException("请先选择表内单元格。"); var revision = Document.Snapshot.Revision;
        var dialog = new FieldsDialog("调整表范围").Note($"表 {table.Name}：{table.Range}。现有内容保留，新行续填计算列。").Text("bottom", "新的末行行号", (table.Range.Bottom + 1).ToString());
        dialog.Validate = () => dialog.Integer("bottom", 1, WorkbookLimits.MaxRows);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new ResizeTableCommand(Document.Snapshot.ActiveSheetId, table.Id, dialog.Integer("bottom", 1, WorkbookLimits.MaxRows) - 1), revision);
    }
    private async Task EditTableAsync(int action)
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable ?? throw new InvalidOperationException("请先选择表内单元格。"); var column = table.Columns[_grid.Current.Column - table.Range.Left]; var revision = Document.Snapshot.Revision;
        var answer = await Dialogs.Text(this, action == 0 ? "重命名表" : action == 1 ? "重命名表列" : "设置计算列",
            action == 2 ? "以 = 开头的公式；将覆盖整列数据区，包括例外值，可撤销。" : "新名称；所有绑定的公式同步更新。", action == 0 ? table.Name : action == 1 ? column.Name : column.CalculatedFormula ?? "=");
        if (answer is null) return;
        await ExecuteAsync(action == 0 ? new RenameTableCommand(Document.Snapshot.ActiveSheetId, table.Id, answer) : action == 1 ? new RenameTableColumnCommand(Document.Snapshot.ActiveSheetId, table.Id, column.Id, answer) :
            new SetTableCalculatedColumnCommand(Document.Snapshot.ActiveSheetId, table.Id, column.Id, answer), revision);
    }
}
