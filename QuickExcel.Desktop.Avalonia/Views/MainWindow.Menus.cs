using System.Collections.Immutable;
using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private static void SeparateBefore(ItemsControl menu, params string[] titles)
    {
        foreach (var title in titles)
        {
            var item = menu.Items.OfType<MenuItem>().FirstOrDefault(i => i.Header as string == title);
            if (item is null) continue;
            var index = menu.Items.IndexOf(item);
            if (index > 0 && menu.Items[index - 1] is not Separator) menu.Items.Insert(index, new Separator());
        }
    }
    private void GroupMainMenus()
    {
        var file = (MenuItem)_menu.Items[0]!;
        var exit = file.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("退出")));
        file.Items.Remove(exit); file.Items.Add(new Separator()); file.Items.Add(exit);
        file.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("另存为"))).InputGesture = new(global::Avalonia.Input.Key.F12);
        file.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("打印预览 / PDF"))).InputGesture = new(global::Avalonia.Input.Key.P, global::Avalonia.Input.KeyModifiers.Control);
        file.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("关闭工作簿"))).InputGesture = new(global::Avalonia.Input.Key.W, global::Avalonia.Input.KeyModifiers.Control);
        SeparateBefore((MenuItem)_menu.Items[1]!, UiText.Get("复制结果值（对外）"), UiText.Get("粘贴  Ctrl+V"), UiText.Get("查找与替换  Ctrl+F"));
        SeparateBefore((MenuItem)_menu.Items[2]!, UiText.Get("向前移动"), UiText.Get("隐藏工作表"), UiText.Get("保护工作表"), UiText.Get("标签颜色"));
        SeparateBefore((MenuItem)_menu.Items[3]!, UiText.Get("标签在底部"));
        var script = _menu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header as string == UiText.Get("脚本"));
        if (script is not null) SeparateBefore(script, UiText.Get("应用脚本并重算"), UiText.Get("签名详情"));
        if (_sheets.ContextMenu is { } tabs) SeparateBefore(tabs, UiText.Get("向前移动"), UiText.Get("隐藏"), UiText.Get("删除"), UiText.Get("标签颜色"));
    }
    private void InstallGridContextMenu()
    {
        var menu = new ContextMenu();
        menu.Opening += (_, e) =>
        {
            if (_grid.IsContextSelectionBlocked || _working || _execution.IsRunning || !CommitEditors()) { e.Cancel = true; return; }
            var ranges = _grid.SelectionRanges;
            var row = ranges.Count == 1 && SelectionPlanning.IsWholeRow(Rectangle(ranges[0]));
            var column = ranges.Count == 1 && SelectionPlanning.IsWholeColumn(Rectangle(ranges[0]));
            BuildGridContextMenu(menu, row != column ? row ? SheetAxis.Row : SheetAxis.Column : null);
        };
        _grid.ContextMenu = menu;
    }
    private void BuildGridContextMenu(ContextMenu menu, SheetAxis? axis)
    {
        menu.Items.Clear();
        if (Document.IsReadOnly)
        {
            menu.Items.Add(CommandMenu(UiText.Get("复制"), () => CopyAsync(false), true, mutation: false));
            menu.Items.Add(CommandMenu(UiText.Get("复制结果值（对外）"), () => CopyAsync(false, valuesOnly: true), true, mutation: false));
            menu.Items.Add(CommandMenu(UiText.Get("超链接…"), HyperlinkAsync, true, mutation: false));
            menu.Items.Add(CommandMenu(UiText.Get("线程评论…"), ThreadCommentsAsync, true, mutation: false));
            if (_grid.SelectedDrawingObject is not null) menu.Items.Add(CommandMenu(UiText.Get("复制对象为图片"), CopyDrawingPictureAsync, true, mutation: false));
            return;
        }
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
        var ranges = _grid.SelectionRanges.ToArray(); var drawing = _grid.SelectedDrawingObject;
        MenuItem Item(string title, Func<Task> run) => CommandMenu(title, async () =>
        {
            if (session != Document.Session.SessionId || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId ||
                snapshot.Revision != Document.Snapshot.Revision || !ranges.SequenceEqual(_grid.SelectionRanges) || drawing != _grid.SelectedDrawingObject)
                throw new InvalidOperationException(UiText.Get("文档或选区已改变，请重新打开右键菜单。"));
            await run();
        });
        void Group(params MenuItem[] items)
        { if (menu.Items.Count > 0) menu.Items.Add(new Separator()); foreach (var item in items) menu.Items.Add(item); }
        MenuItem Submenu(string title, params MenuItem[] items) => new() { Header = title, ItemsSource = items };
        Group(Item(UiText.Get("剪切"), () => CopyAsync(true)), Item(UiText.Get("复制"), () => CopyAsync(false)), Item(UiText.Get("粘贴"), () => PasteAsync("all")),
            Submenu(UiText.Get("选择性粘贴"), Item(UiText.Get("只贴值"), () => PasteAsync("values")), Item(UiText.Get("粘贴公式"), () => PasteAsync("formulas")),
                Item(UiText.Get("只贴格式"), () => PasteAsync("formats")), Item(UiText.Get("转置粘贴"), () => PasteAsync("transpose"))));
        if (drawing is not null)
        {
            Group(Item(UiText.Get("复制对象为图片"), CopyDrawingPictureAsync), Item(UiText.Get("粘贴为可编辑图表（附带数据快照）"), () => PasteAsync("chart")), Item(UiText.Get("粘贴为图片"), () => PasteAsync("image")));
            Group(Item(UiText.Get("删除对象"), ClearAsync)); return;
        }
        if (axis is { } a)
        {
            var name = a == SheetAxis.Row ? UiText.Get("行") : UiText.Get("列");
            var copied = Item(UiText.Format($"插入复制的{name}"), () => InsertCopiedAxisAsync(a));
            copied.IsEnabled = _clipboardAxes?.Axis == a && _clipboardToken is not null && _clipboardSession == session && _clipboardAxes.Cells.Revision == snapshot.Revision;
            ToolTip.SetTip(copied, UiText.Format($"先选择完整{name}并复制，再在目标{name}前插入。可一次撤销。"));
            Group(Item(UiText.Format($"插入{name}"), () => ShiftAxisAsync(a, false)), copied, Item(UiText.Format($"删除{name}"), () => ShiftAxisAsync(a, true)));
            Group(Item(UiText.Get("清除内容"), ClearAsync), Item(UiText.Get("清除格式"), ClearFormatsAsync));
            Group(Item(a == SheetAxis.Row ? UiText.Get("调整行高…") : UiText.Get("调整列宽…"), () => AxisSizeAsync(a)),
                Item(a == SheetAxis.Row ? UiText.Get("自动调整行高") : UiText.Get("自动调整列宽"), () => AutoFitSelectionAsync(a)),
                Item(UiText.Format($"隐藏{name}"), () => SetAxisVisibilityAsync(a, true)), Item(UiText.Format($"取消隐藏{name}"), () => SetAxisVisibilityAsync(a, false)),
                Item(UiText.Format($"显示所有{name}"), () => SetAxisVisibilityAsync(a, false, true)));
        }
        else
        {
            Group(Submenu(UiText.Get("插入"), Item(UiText.Get("插入行"), () => ShiftAxisAsync(SheetAxis.Row, false)), Item(UiText.Get("插入列"), () => ShiftAxisAsync(SheetAxis.Column, false))),
                Submenu(UiText.Get("删除"), Item(UiText.Get("删除行"), () => ShiftAxisAsync(SheetAxis.Row, true)), Item(UiText.Get("删除列"), () => ShiftAxisAsync(SheetAxis.Column, true))),
                Item(UiText.Get("清除内容"), ClearAsync), Item(UiText.Get("清除格式"), ClearFormatsAsync));
            Group(Item(UiText.Get("设置单元格格式…"), CellFormatAsync), Item(UiText.Get("数字格式…"), NumberFormatAsync), Item(UiText.Get("富文本编辑…"), RichTextAsync),
                Item(UiText.Get("合并单元格"), () => MergeAsync(false)), Item(UiText.Get("取消合并单元格"), () => MergeAsync(true)));
            Group(Submenu(UiText.Get("排序与筛选"), Item(UiText.Get("升序排序"), () => SortAsync(false)), Item(UiText.Get("降序排序"), () => SortAsync(true)), Item(UiText.Get("建立筛选"), EstablishFilterAsync), Item(UiText.Get("清除筛选条件"), ClearFiltersAsync)),
                Submenu(UiText.Get("填充"), Item(UiText.Get("向下填充"), () => FillDirectionAsync(false)), Item(UiText.Get("向右填充"), () => FillDirectionAsync(true)), Item(UiText.Get("填充序列…"), SeriesAsync)),
                Item(UiText.Get("批注…"), NoteAsync), Item(UiText.Get("超链接…"), HyperlinkAsync), Item(UiText.Get("线程评论…"), ThreadCommentsAsync));
            Group(Item(UiText.Get("粘贴为可编辑图表（附带数据快照）"), () => PasteAsync("chart")), Item(UiText.Get("粘贴为图片"), () => PasteAsync("image")));
        }
        InstallPageBreakMenu(menu);
    }
    private async Task InsertCopiedAxisAsync(SheetAxis axis)
    {
        if (!await CommitEditorsAsync()) return;
        var source = _clipboardAxes;
        var revision = Document.Snapshot.Revision; var session = Document.Session.SessionId; var sheet = Document.Snapshot.ActiveSheetId; var range = SelectedRange();
        var clipboard = await _clipboard.ReadAsync();
        if (source is null || source.Axis != axis || _clipboardToken is null || clipboard.Token != _clipboardToken ||
            _clipboardSession != session || session != Document.Session.SessionId || sheet != Document.Snapshot.ActiveSheetId ||
            revision != Document.Snapshot.Revision || range != SelectedRange())
            throw new InvalidOperationException(UiText.Get("请先在当前工作簿复制完整行列，再重新打开目标位置的右键菜单。"));
        var first = axis == SheetAxis.Row ? range.Top : range.Left;
        if (!await ExecuteAsync(new InsertCopiedAxisCommand(source, sheet, first), revision)) return;
        _clipboardAxes = null; _clipboardCells = null; _clipboardToken = null;
        _grid.SelectRange(axis == SheetAxis.Row ? new(first, 0, first + source.Count - 1, WorkbookLimits.MaxColumns - 1)
            : new(0, first, WorkbookLimits.MaxRows - 1, first + source.Count - 1), axis == SheetAxis.Row ? new(first, 0) : new(0, first));
        await CalculatePendingAsync(); _status.Text = UiText.Get("已插入复制的行列，可一次撤销。");
    }
    private async Task SetAxisVisibilityAsync(SheetAxis axis, bool hidden, bool all = false)
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var row = axis == SheetAxis.Row;
        await ExecuteAsync(new SetAxisMetadataCommand(Document.Snapshot.ActiveSheetId, axis, all ? 0 : row ? range.Top : range.Left,
            all ? row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns : row ? range.Rows : range.Columns, hidden: hidden));
    }
    private async Task AutoFitSelectionAsync(SheetAxis axis)
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var row = axis == SheetAxis.Row;
        var count = row ? range.Rows : range.Columns;
        if (count > 1000) throw new InvalidOperationException(UiText.Get("一次最多自动调整 1,000 行或列，请缩小选区。"));
        var sizes = Enumerable.Range(row ? range.Top : range.Left, count).Select(index =>
        { var pixels = _grid.MeasureAutoFit(axis, index); return new KeyValuePair<int, double>(index, row ? pixels : Math.Max(.01, (pixels - 5) / 7)); }).ToArray();
        await ExecuteAsync(new SetAxisMetadataCommand(Document.Snapshot.ActiveSheetId, axis, row ? range.Top : range.Left, count) { Sizes = sizes.ToImmutableDictionary() });
    }
}
