using System.Collections.Immutable;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Controls.Primitives.PopupPositioning;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Popup? _filterPopup;
    private Func<bool>? _filterContextCurrent;
#if QUICKEXCEL_DIAGNOSTICS
    private Func<FilterPanel, Task>? _filterDriver;
#endif
    private void CloseFilterPanel() { if (_filterPopup is { } popup) popup.IsOpen = false; }
    private void InstallDataCommands()
    {
        RibbonGroup(UiText.Get("数据"), UiText.Get("排序和筛选"), RibbonButton(RibbonIcon.SortAscending, UiText.Get("升序"), () => SortAsync(false)), RibbonButton(RibbonIcon.SortDescending, UiText.Get("降序"), () => SortAsync(true)),
            RibbonButton(RibbonIcon.SortLevels, UiText.Get("多级排序"), () => SortAsync(null)), RibbonButton(RibbonIcon.Filter, UiText.Get("筛选区域"), EstablishFilterAsync), RibbonButton(RibbonIcon.FilterColumn, UiText.Get("当前列筛选"), () => FilterAsync(_grid.Current.Column)),
            RibbonButton(RibbonIcon.ClearFilters, UiText.Get("清除全部筛选条件"), ClearFiltersAsync), RibbonButton(RibbonIcon.RemoveFilter, UiText.Get("取消筛选区域"), RemoveFilterAsync), RibbonButton(RibbonIcon.ClearSort, UiText.Get("清除排序标记"), ClearSortAsync));
        RibbonGroup(UiText.Get("插入"), UiText.Get("表格"), RibbonButton(RibbonIcon.CreateTable, UiText.Get("创建表"), CreateTableAsync), RibbonButton(RibbonIcon.ResizeTable, UiText.Get("调整表范围"), ResizeTableAsync), RibbonButton(RibbonIcon.RenameTable, UiText.Get("重命名表"), () => EditTableAsync(0)),
            RibbonButton(RibbonIcon.RenameColumn, UiText.Get("重命名表列"), () => EditTableAsync(1)), RibbonButton(RibbonIcon.CalculatedColumn, UiText.Get("设置计算列"), () => EditTableAsync(2)));
        RibbonGroup(UiText.Get("数据"),UiText.Get("透视表"),RibbonButton(RibbonIcon.CreatePivot, UiText.Get("创建透视表"),CreatePivotAsync),RibbonButton(RibbonIcon.RefreshPivot, UiText.Get("刷新透视表"),RefreshPivotAsync),RibbonButton(RibbonIcon.PivotValues, UiText.Get("透视表转换为值"),ConvertPivotToValuesAsync));
        _grid.FilterRequested += column => _ = PerformAsync(() => FilterAsync(column));
        AddHandler(KeyDownEvent, async (_, e) =>
        { if (!e.Handled && e.Key == Key.Down && e.KeyModifiers == KeyModifiers.Alt && !_grid.IsEditing && !_formula.IsComposing &&
              !AutomationPickerOwnsKeyboard)
            { e.Handled = true; await PerformAsync(() => FilterAsync(_grid.Current.Column)); } }, RoutingStrategies.Tunnel);
    }
    private WorksheetTable? CurrentTable => Document.Snapshot.Sheets[SheetIndex].Data.Tables.FirstOrDefault(t => t.Range.Contains(_grid.Current.Row, _grid.Current.Column));
    private async Task SortAsync(bool? descending)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var table = CurrentTable;
        var filter = table?.Filter ?? (table is null && sheet.Data.Filter?.Range.Contains(_grid.Current.Row, _grid.Current.Column) == true ? sheet.Data.Filter : null);
        var selected = Rectangle(SelectedRange());
        if (table is null && filter is null && (SelectionPlanning.IsWholeColumn(selected) || SelectionPlanning.IsWholeRow(selected)))
        {
            var records = SortPlanning.WholeAxisRegion(sheet, selected, _grid.Current);
            // A sheet-header selection may keep its active cell above the table. Recover its explicit
            // header/total-row contract only when the entire inferred block belongs to that table/filter.
            table = sheet.Data.Tables.FirstOrDefault(t => t.Range == records);
            filter = table?.Filter ?? (sheet.Data.Filter?.Range == records ? sheet.Data.Filter : null);
        }
        var block = table is not null ? new SortBlock(table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) }, table.HeaderRow) :
            filter is not null ? new SortBlock(filter.Range, true) : SortPlanning.Resolve(sheet, selected, _grid.Current);
        if (block.Range.Rows < 2) throw new ArgumentException(UiText.Get("请选择包含多行记录的数据区域。"));
        SortDescriptor descriptor;
        if (descending is { } order) descriptor = new(sheet.Id, new(block.Range.Top, block.Range.Left), new(block.Range.Bottom, block.Range.Right),
            [new(Math.Clamp(_grid.Current.Column, block.Range.Left, block.Range.Right), order)], block.HasHeader, filter is not null && table is null);
        else
        {
            var dialog = new SortDialog(snapshot, sheet, block.Range, block.HasHeader, CurrentView.SortKeys, table is not null || filter is not null, filter is not null && table is null);
            if (!await ShowOwnedAsync<bool>(dialog) || dialog.Descriptor is null) return; descriptor = dialog.Descriptor;
        }
        if (await ExecuteAsync(new SortRangeCommand(descriptor), snapshot.Revision)) await CalculatePendingAsync();
    }
    private async Task EstablishFilterAsync()
    {
        if (!await FlushEditsAsync()) return;
        if (CurrentTable is not null) throw new InvalidOperationException(UiText.Get("表格已有独立筛选；请使用当前列筛选。"));
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        var block = PersistentFiltering.ResolveRange(sheet, Rectangle(SelectedRange()), _grid.Current);
        var range = new GridRange(block.Top, block.Left, block.Bottom, block.Right);
        new DataRegion(Guid.NewGuid(), Document.Snapshot.ActiveSheetId, range).Validate();
        await ExecuteAsync(new SetAutoFilterCommand(Document.Snapshot.ActiveSheetId, new(Rectangle(range))));
    }
    private async Task FilterAsync(int column)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var sheet = snapshot.Sheets[SheetIndex]; var table = CurrentTable;
        var filter = table is null ? sheet.Data.Filter : table.Filter;
        if (filter is null || column < filter.Range.Left || column > filter.Range.Right) throw new InvalidOperationException(UiText.Get("请先建立筛选区域，并选择区域内的一列。"));
        bool ContextCurrent() => session == Document.Session.SessionId && sheet.Id == Document.Snapshot.ActiveSheetId && snapshot.Revision == Document.Snapshot.Revision;
        IReadOnlyList<FilterCandidate> candidates = []; string? candidateError = null;
        try { candidates = await Task.Run(() => PersistentFiltering.Candidates(snapshot, sheet, column, table?.Id)); }
        catch (Exception ex) { candidateError = ex.Message; }
        if (!ContextCurrent() || _windowClosed) return;
        var panel = new FilterPanel(WorkbookGrid.ColumnName(column) + UiText.Get(" 列筛选"), candidates,
            filter.Columns.TryGetValue(column, out var current) ? PersistentFiltering.Restore(current) : null, candidateError);
        var completion = new TaskCompletionSource<FilterChoice?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var anchor = _grid.GetCellBounds(new(filter.Range.Top, column)) ?? _grid.GetCellBounds(_grid.Current) ?? new Rect(50, 0, 100, 28);
        var popup = new Popup { Child = panel, PlacementTarget = _grid, Placement = PlacementMode.AnchorAndGravity,
            PlacementRect = anchor, PlacementAnchor = PopupAnchor.BottomRight, PlacementGravity = PopupGravity.BottomLeft,
            IsLightDismissEnabled = true, MaxHeight = Math.Max(280, Math.Min(640, ClientSize.Height - 40)) };
        panel.MaxHeight = popup.MaxHeight;
        panel.Completed += choice =>
        {
            // Conversion can reject unsupported combinations too; keep those errors inside the menu.
            if (choice?.Filter is { } value) _ = PersistentFiltering.Convert(value, snapshot.DateSystem);
            completion.TrySetResult(choice); popup.IsOpen = false;
        };
        popup.Closed += (_, _) => completion.TrySetResult(null);
        _filterPopup = popup; _filterContextCurrent = ContextCurrent;
        // Keep the popup in the logical tree so Fluent templates, resources and input routing apply.
        _book.Children.Insert(0, popup);
        FilterChoice? result;
        try
        {
            popup.IsOpen = true; panel.FocusSearch();
#if QUICKEXCEL_DIAGNOSTICS
            if (Program.SmokeDirectory is not null && _filterDriver is { } driver) { _filterDriver = null; await driver(panel); }
#endif
            result = await completion.Task;
        }
        finally { popup.IsOpen = false; popup.Child = null; _book.Children.Remove(popup); _filterPopup = null; _filterContextCurrent = null; }
        if (result is null) return;
        if (!ContextCurrent()) throw new InvalidOperationException(UiText.Get("筛选面板打开期间文档已改变，请重新打开。"));
        if (result.Action is "ascending" or "descending")
        {
            var range = table is null ? filter.Range : table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) };
            if (range.Rows < 2) throw new ArgumentException(UiText.Get("请选择包含多行记录的数据区域。"));
            if (await ExecuteAsync(new SortRangeCommand(new(sheet.Id, new(range.Top, range.Left), new(range.Bottom, range.Right), [new(column, result.Action == "descending")], table?.HeaderRow ?? true, table is null)), snapshot.Revision)) await CalculatePendingAsync();
            return;
        }
        if (result.Action == "color") { await ColorFilterAsync(snapshot, sheet, table, filter, column); return; }
        if (result.Action == "top") { await TopFilterAsync(snapshot, sheet, table, filter, column); return; }
        if (result.Action == "dates") { await DateGroupFilterAsync(snapshot, sheet, table, filter, column); return; }
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
        var dialog = new FieldsDialog(UiText.Get("创建表")).Note(UiText.Get("所选区域首行须为非空、唯一的文本表头。")).Text("name", UiText.Get("表名"), "QE_Table" + suffix)
            .Choice("style", UiText.Get("表样式"), [UiText.Get("蓝色"), UiText.Get("绿色"), UiText.Get("橙色"), UiText.Get("灰色")]);
        if (!await dialog.ShowAsync(this)) return;
        var style = TableStyles.Resolve(snapshot, new[] { "TableStyleMedium2", "TableStyleMedium4", "TableStyleMedium9", "TableStyleMedium1" }[dialog.Index("style")]);
        await ExecuteAsync(new CreateTableCommand(snapshot.ActiveSheetId, Rectangle(range), dialog.String("name"), style), snapshot.Revision);
    }
    private async Task ResizeTableAsync()
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable ?? throw new InvalidOperationException(UiText.Get("请先选择表内单元格。")); var revision = Document.Snapshot.Revision;
        var dialog = new FieldsDialog(UiText.Get("调整表范围")).Note(UiText.Format($"表 {table.Name}：{table.Range}。现有内容保留，新行续填计算列。")).Text("bottom", UiText.Get("新的末行行号"), (table.Range.Bottom + 1).ToString());
        dialog.Validate = () => dialog.Integer("bottom", 1, WorkbookLimits.MaxRows);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new ResizeTableCommand(Document.Snapshot.ActiveSheetId, table.Id, dialog.Integer("bottom", 1, WorkbookLimits.MaxRows) - 1), revision);
    }
    private async Task EditTableAsync(int action)
    {
        if (!await FlushEditsAsync()) return;
        var table = CurrentTable ?? throw new InvalidOperationException(UiText.Get("请先选择表内单元格。")); var column = table.Columns[_grid.Current.Column - table.Range.Left]; var revision = Document.Snapshot.Revision;
        var answer = await Dialogs.Text(this, action == 0 ? UiText.Get("重命名表") : action == 1 ? UiText.Get("重命名表列") : UiText.Get("设置计算列"),
            action == 2 ? UiText.Get("以 = 开头的公式；将覆盖整列数据区，包括例外值，可撤销。") : UiText.Get("新名称；所有绑定的公式同步更新。"), action == 0 ? table.Name : action == 1 ? column.Name : column.CalculatedFormula ?? "=");
        if (answer is null) return;
        await ExecuteAsync(action == 0 ? new RenameTableCommand(Document.Snapshot.ActiveSheetId, table.Id, answer) : action == 1 ? new RenameTableColumnCommand(Document.Snapshot.ActiveSheetId, table.Id, column.Id, answer) :
            new SetTableCalculatedColumnCommand(Document.Snapshot.ActiveSheetId, table.Id, column.Id, answer), revision);
    }
}
