using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallSheetCommands()
    {
        var sheet = (MenuItem)_menu.Items[2]!;
        foreach (var (name, action) in new (string, Func<Task>)[]
        {
            ("复制工作表", CopySheetAsync), ("向前移动", () => MoveSheetAsync(-1)), ("向后移动", () => MoveSheetAsync(1)),
            ("隐藏工作表", HideSheetAsync), ("取消隐藏", UnhideSheetAsync), ("保护工作表", ProtectSheetAsync)
        }) sheet.Items.Add(CommandMenu(name, action));
        sheet.Items.Add(SheetColorMenu());
        sheet.Items.Add(new Separator());
        sheet.Items.Add(CommandMenu("工作表管理器", ShowSheetManagerAsync));
        var context = new ContextMenu();
        foreach (var (name, action) in new (string, Func<Task>)[] { ("新建工作表", AddSheetAsync), ("重命名", RenameSheetAsync), ("复制", CopySheetAsync), ("向前移动", () => MoveSheetAsync(-1)), ("向后移动", () => MoveSheetAsync(1)), ("隐藏", HideSheetAsync), ("取消隐藏", UnhideSheetAsync), ("删除", DeleteSheetAsync) })
            context.Items.Add(CommandMenu(name, action));
        context.Items.Add(SheetColorMenu());
        context.Items.Add(new Separator());
        context.Items.Add(CommandMenu("工作表管理器", ShowSheetManagerAsync));
        _sheets.ContextMenu = context;
        _sheets.Styles.Add(new Style(s => s.OfType<ListBoxItem>()) { Setters =
        {
            new Setter(PaddingProperty, new Thickness(8, 4)), new Setter(MinHeightProperty, 0d),
            new Setter(MarginProperty, new Thickness(0)), new Setter(BorderThicknessProperty, new Thickness(0))
        } });
        _sheets.ItemTemplate = new FuncDataTemplate<SheetItem>((item, _) =>
        {
            if (item is null) return null;
            var label = new TextBlock { Name = "SheetTabLabel", Text = item.Name, VerticalAlignment = VerticalAlignment.Center };
            var tab = new global::Avalonia.Controls.Grid(); tab.Children.Add(label);
            // The color line occupies the existing bottom padding instead of adding height.
            tab.Children.Add(new Border { Height = 2, VerticalAlignment = VerticalAlignment.Bottom, Margin = new(0, 0, 0, -3), BorderThickness = new(0, 0, 0, 2),
                BorderBrush = item.TabColor is { } color && Color.TryParse("#" + color, out var parsed) ? new SolidColorBrush(parsed) : Brushes.Transparent });
            AutomationProperties.SetName(tab, item.Name + (item.TabColor is { } value ? $"，标签颜色 #{value}" : "，默认标签颜色"));
            return tab;
        });
        // Resolve the tab under the context request before committing the original
        // cell draft. Keyboard requests use the selected tab; failed commits do not open a menu.
        _sheets.AddHandler(ContextRequestedEvent, async (_, e) =>
        {
            e.Handled = true;
            var source = e.Source as Visual;
            var container = source as ListBoxItem ?? source?.FindAncestorOfType<ListBoxItem>();
            var target = (container?.DataContext as SheetItem)?.Id ?? Document.Snapshot.ActiveSheetId;
            await PerformAsync(async () =>
            {
                if (!await CommitEditorsAsync()) return;
                var targetSheet = Document.Snapshot.GetSheet(target);
                Document.Session.UpdateNavigation(target, targetSheet.View.ActiveCell); Synchronize();
                context.Open(_sheets);
            });
        }, RoutingStrategies.Tunnel);
        RibbonGroup("视图", "窗口", RibbonButton(RibbonIcon.FreezePanes, "冻结至当前格", () => FreezeAsync(_grid.Current.Row, _grid.Current.Column)),
            RibbonButton(RibbonIcon.FreezeRow, "冻结首行", () => FreezeAsync(1, 0)), RibbonButton(RibbonIcon.FreezeColumn, "冻结首列", () => FreezeAsync(0, 1)), RibbonButton(RibbonIcon.Unfreeze, "取消冻结", () => FreezeAsync(0, 0)),
            RibbonButton(RibbonIcon.Zoom, "缩放", ZoomDialogAsync));
        RibbonGroup("开始", "行列", RibbonButton(RibbonIcon.InsertRow, "插入行", () => ShiftAxisAsync(SheetAxis.Row, false)), RibbonButton(RibbonIcon.DeleteRow, "删除行", () => ShiftAxisAsync(SheetAxis.Row, true)),
            RibbonButton(RibbonIcon.InsertColumn, "插入列", () => ShiftAxisAsync(SheetAxis.Column, false)), RibbonButton(RibbonIcon.DeleteColumn, "删除列", () => ShiftAxisAsync(SheetAxis.Column, true)),
            RibbonButton("隐藏/显示行列", AxisVisibilityAsync), RibbonButton("移动行列", MoveAxisAsync));
        RibbonGroup("公式", "名称", RibbonButton("名称管理", NamesAsync));
        RibbonGroup("审阅", "保护", RibbonButton("工作表保护", ProtectSheetAsync), RibbonButton("单元格保护", ProtectCellsAsync));
        _grid.ViewportChanged += () =>
        {
            if (!_synchronizing && !_workspace.IsBusy && _grid.ScrollingCell is { } cell)
                Document.Session.UpdateNavigation(Document.Snapshot.ActiveSheetId, _grid.Current, cell);
        };
    }
    private MenuItem SheetColorMenu()
    {
        var menu = new MenuItem { Header = "标签颜色" };
        foreach (var item in ColorMenuItems("恢复默认颜色", SetSheetTabColorAsync, CustomSheetTabColorAsync)) menu.Items.Add(item);
        return menu;
    }
    private async Task SetSheetTabColorAsync(string? color)
    {
        if (await CommitEditorsAsync())
            await ExecuteAsync(new SetSheetTabColorCommand(Document.Snapshot.ActiveSheetId, color));
    }
    private Task CustomSheetTabColorAsync() => CustomSheetTabColorAsync(this, Document.Snapshot.ActiveSheetId);
    private async Task CustomSheetTabColorAsync(Window owner, SheetId target)
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(target);
        var color = await new CustomColorDialog(sheet.TabColor, "自定义标签颜色").ShowAsync(owner);
        if (color is not null) await ExecuteAsync(new SetSheetTabColorCommand(sheet.Id, color), snapshot.Revision);
    }
    private Task RenameSheetAsync() => RenameSheetAsync(this, Document.Snapshot.ActiveSheetId);
    private async Task RenameSheetAsync(Window owner, SheetId target)
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(target);
        var name = await Dialogs.Text(owner, "重命名工作表", "工作表名称", sheet.Name);
        if (name is not null) await ExecuteAsync(new RenameSheetCommand(target, name), snapshot.Revision);
    }
    private Task DeleteSheetAsync() => DeleteSheetAsync(this, Document.Snapshot.ActiveSheetId);
    private async Task DeleteSheetAsync(Window owner, SheetId target)
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(target);
        if (await Dialogs.Confirm(owner, "删除工作表", $"删除“{sheet.Name}”？引用此表的公式将变为 #REF!，可撤销。"))
            await ExecuteAsync(new DeleteSheetCommand(target), snapshot.Revision);
    }
    private Task CopySheetAsync() => CopySheetAsync(this, Document.Snapshot.ActiveSheetId);
    private async Task<SheetId?> CopySheetAsync(Window owner, SheetId target)
    {
        if (!await CommitEditorsAsync()) return null;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(target);
        var name = await Dialogs.Text(owner, "复制工作表", "新工作表名称", NextCopyName(snapshot, sheet.Name));
        if (name is null) return null;
        var copy = new CopySheetCommand(target, name, snapshot.Sheets.IndexOf(sheet) + 1);
        return await ExecuteAsync(copy, snapshot.Revision) ? copy.Id : null;
    }
    private static string NextCopyName(WorkbookSnapshot snapshot, string name)
    {
        for (int i = 2; ; i++)
        {
            var suffix = $" ({i})";
            var length = Math.Min(name.Length, WorkbookLimits.MaxSheetNameLength - suffix.Length);
            if (char.IsHighSurrogate(name[length - 1])) length--;
            var candidate = name[..length] + suffix;
            if (!snapshot.Sheets.Any(s => s.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase))) return candidate;
        }
    }
    private async Task MoveSheetAsync(int direction)
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var index = SheetIndex + direction;
        while (index >= 0 && index < snapshot.Sheets.Length && snapshot.Sheets[index].Visibility != SheetVisibility.Visible) index += direction;
        if (index >= 0 && index < snapshot.Sheets.Length) await ExecuteAsync(new ReorderSheetCommand(snapshot.ActiveSheetId, index));
    }
    private async Task HideSheetAsync()
    { if (await CommitEditorsAsync()) await ExecuteAsync(new SetSheetVisibilityCommand(Document.Snapshot.ActiveSheetId, SheetVisibility.Hidden)); }
    private async Task UnhideSheetAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var hidden = snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Hidden).ToArray();
        if (hidden.Length == 0) { _status.Text = "没有可取消隐藏的普通工作表。"; return; }
        var dialog = new FieldsDialog("取消隐藏工作表").Choice("sheet", "工作表", hidden.Select(s => s.Name).ToArray());
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetSheetVisibilityCommand(hidden[dialog.Index("sheet")].Id, SheetVisibility.Visible), snapshot.Revision);
    }
    private async Task FreezeAsync(int rows, int columns)
    {
        if (!await CommitEditorsAsync()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        await ExecuteAsync(new SetSheetViewCommand(sheet.Id, sheet.View with { FrozenRows = rows, FrozenColumns = columns, PaneTopLeftCell = new(rows, columns) }));
    }
    private async Task ZoomDialogAsync()
    {
        var dialog = new FieldsDialog("缩放").Text("zoom", "缩放比例（10–400%）", Document.Snapshot.Sheets[SheetIndex].View.Zoom.ToString());
        dialog.Validate = () => dialog.Integer("zoom", 10, 400);
        if (await dialog.ShowAsync(this)) await ZoomAsync(dialog.Integer("zoom", 10, 400));
    }
    private async Task ShiftAxisAsync(SheetAxis axis, bool delete)
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var revision = Document.Snapshot.Revision;
        var first = axis == SheetAxis.Row ? range.Top : range.Left; var count = axis == SheetAxis.Row ? range.Rows : range.Columns;
        if (delete && !await Dialogs.Confirm(this, "删除行列", $"删除所选 {count} {(axis == SheetAxis.Row ? "行" : "列")}？可撤销。")) return;
        await ExecuteAsync(new ShiftAxisCommand(Document.Snapshot.ActiveSheetId, axis, first, count, delete), revision);
    }
    private async Task MoveAxisAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var range = SelectedRange();
        var dialog = new FieldsDialog("移动行列").Choice("axis", "移动对象", ["行", "列"]).Text("destination", "目标起点（从 1 开始的序号）", "1");
        dialog.Validate = () => dialog.Integer("destination", 1, dialog.Index("axis") == 0 ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new MoveAxisCommand(Document.Snapshot.ActiveSheetId, (SheetAxis)dialog.Index("axis"),
            dialog.Index("axis") == 0 ? range.Top : range.Left, dialog.Index("axis") == 0 ? range.Rows : range.Columns,
            dialog.Integer("destination", 1, dialog.Index("axis") == 0 ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns) - 1), revision);
    }
    private async Task AxisVisibilityAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var revision = Document.Snapshot.Revision;
        var dialog = new FieldsDialog("隐藏或显示行列").Choice("axis", "对象", ["行", "列"]).Choice("action", "操作", ["隐藏所选", "显示所选", "显示当前表全部"]);
        if (!await dialog.ShowAsync(this)) return;
        var row = dialog.Index("axis") == 0; var all = dialog.Index("action") == 2;
        await ExecuteAsync(new SetAxisMetadataCommand(Document.Snapshot.ActiveSheetId, row ? SheetAxis.Row : SheetAxis.Column,
            all ? 0 : row ? range.Top : range.Left, all ? row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns : row ? range.Rows : range.Columns, hidden: dialog.Index("action") == 0), revision);
    }
    private Task ProtectSheetAsync() => ProtectSheetAsync(this, Document.Snapshot.ActiveSheetId);
    private async Task ProtectSheetAsync(Window owner, SheetId target)
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.GetSheet(target); var protection = sheet.Data.Protection;
        if (protection?.HasPassword == true) throw new NotSupportedException("此工作表带密码保护，请在 Excel 中取消保护后再修改保护设置。");
        var previous = protection ?? new WorksheetProtection();
        var dialog = new FieldsDialog("保护工作表").Note("工作表保护限制编辑操作，不加密文件。")
            .Check("enabled", "启用工作表保护", true)
            .Check("locked", "允许选择锁定单元格", !previous.SelectLockedCells).Check("unlocked", "允许选择未锁定单元格", !previous.SelectUnlockedCells)
            .Check("format", "允许设置单元格格式", !previous.FormatCells).Check("rows", "允许设置行格式", !previous.FormatRows).Check("columns", "允许设置列格式", !previous.FormatColumns)
            .Check("insertRows", "允许插入行", !previous.InsertRows).Check("insertColumns", "允许插入列", !previous.InsertColumns)
            .Check("deleteRows", "允许删除行", !previous.DeleteRows).Check("deleteColumns", "允许删除列", !previous.DeleteColumns)
            .Check("objects", "允许编辑对象和批注", !previous.Objects).Check("sort", "允许排序", !previous.Sort).Check("filter", "允许筛选", !previous.AutoFilter);
        if (await dialog.ShowAsync(owner)) await ExecuteAsync(new SetProtectionCommand(sheet.Id, !dialog.Boolean("enabled") ? null : previous with
        {
            SelectLockedCells = !dialog.Boolean("locked"), SelectUnlockedCells = !dialog.Boolean("unlocked"), FormatCells = !dialog.Boolean("format"),
            FormatRows = !dialog.Boolean("rows"), FormatColumns = !dialog.Boolean("columns"), Sort = !dialog.Boolean("sort"), AutoFilter = !dialog.Boolean("filter"),
            InsertRows = !dialog.Boolean("insertRows"), InsertColumns = !dialog.Boolean("insertColumns"), DeleteRows = !dialog.Boolean("deleteRows"), DeleteColumns = !dialog.Boolean("deleteColumns"), Objects = !dialog.Boolean("objects")
        }), snapshot.Revision);
    }
    private async Task ProtectCellsAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var revision = Document.Snapshot.Revision; var style = CurrentStyle;
        var dialog = new FieldsDialog("单元格保护").Check("locked", "锁定", style.Locked).Check("hidden", "隐藏公式", style.FormulaHidden)
            .Note("这些限制在工作表保护启用后生效。");
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetCellProtectionCommand(Document.Snapshot.ActiveSheetId, Rectangle(range), dialog.Boolean("locked"), dialog.Boolean("hidden")), revision);
    }
    private async Task NamesAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var names = snapshot.DefinedNames.ToArray(); var range = SelectedRange(); var sheet = snapshot.Sheets[SheetIndex];
        var choose = new FieldsDialog("名称管理").Choice("name", "名称 / 作用域", new[] { "新建名称" }.Concat(names.Select(n => n.Name + " · " + (n.Scope is { } id ? snapshot.GetSheet(id).Name : "工作簿"))).ToArray())
            .Choice("action", "操作", ["新建 / 编辑", "删除"]);
        if (!await choose.ShowAsync(this)) return;
        var selected = choose.Index("name") == 0 ? null : names[choose.Index("name") - 1];
        if (choose.Index("action") == 1)
        { if (selected is not null) await ExecuteAsync(new DeleteDefinedNameCommand(selected.Name, selected.Scope), snapshot.Revision); return; }
        var formula = "=" + FormulaSyntax.Print(new ReferenceNode(sheet.Name, range.Top, range.Left, range.Bottom, range.Right, true, true, true, true, range.Count > 1));
        var edit = new FieldsDialog(selected is null ? "新建名称" : "编辑名称").Text("name", "名称", selected?.Name ?? "").Text("formula", "引用位置 / 公式", selected?.Formula ?? formula);
        if (selected is null) edit.Choice("scope", "作用域", ["工作簿", "当前工作表"]);
        else edit.Note("作用域：" + (selected.Scope is { } id ? snapshot.GetSheet(id).Name : "工作簿"));
        if (!await edit.ShowAsync(this)) return;
        DocumentCommand command = selected is null ? new DefineNameCommand(new(edit.String("name"), edit.String("formula"), edit.Index("scope") == 0 ? null : sheet.Id)) :
            new EditDefinedNameCommand(selected.Name, selected.Scope, edit.String("name"), edit.String("formula"));
        if (await ExecuteAsync(command, snapshot.Revision)) await CalculatePendingAsync();
    }
}
