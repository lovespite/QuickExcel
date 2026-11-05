using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Clipboard;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private IClipboardExchange _clipboard = null!;
    private WorkbookClipboard? _clipboardCells;
    private AxisClipboard? _clipboardAxes;
    private string? _clipboardToken;
    private Guid _clipboardSession;
    private long _clipboardRevision;
    private SheetViewState CurrentView
    {
        get
        {
            var view = new SheetViewState(); view.Synchronize(Document.Snapshot, Document.Snapshot.ActiveSheetId, _grid.Current);
            view.RefreshFilter(Document.Snapshot); return view;
        }
    }
    private GridRange SelectedRange(bool allowObject = false)
    {
        if (!allowObject && _grid.SelectedDrawingObject is not null) throw new InvalidOperationException("此操作需要单元格选区；请先点击单元格或按 Esc 取消对象选择。");
        var ranges = _grid.SelectionRanges;
        if (ranges.Count == 0) return new(_grid.Current.Row, _grid.Current.Column, _grid.Current.Row, _grid.Current.Column);
        if (ranges.Count != 1) throw new ArgumentException("此操作需要一个连续矩形选区。");
        return ranges[0];
    }
    private static CellRectangle Rectangle(GridRange range) => new(range.Top, range.Left, range.Bottom, range.Right);
    private CellAddress[] SelectedAddresses(int limit = 20_000)
    {
        if (_grid.SelectedDrawingObject is not null) throw new InvalidOperationException("此操作需要单元格选区；请先取消对象选择。");
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        var ranges = _grid.SelectionRanges.Select(r => SelectionPlanning.BoundToUsedCells(sheet, Rectangle(r))).ToArray();
        if (ranges.Sum(r => r.Count) > limit) throw new ArgumentException($"实际操作区域超过 {limit:N0} 格，请缩小选区。");
        return ranges.SelectMany(r => Enumerable.Range(r.Top, r.Rows).SelectMany(row =>
            Enumerable.Range(r.Left, r.Columns).Select(column => new CellAddress(sheet.Id, row, column)))).Distinct().ToArray();
    }
    private async Task<bool> ExecuteAsync(DocumentCommand command, long? revision = null)
    {
        if (!await CommitEditorsAsync()) return false;
        var expected = revision ?? Document.Snapshot.Revision;
        try { await _workspace.Editor.ExecuteAsync(command, ConfirmValidationAsync, expected); }
        catch (DataValidationException ex) when (ex.CanOverride) { Error(ex); return false; }
        Synchronize(); return true;
    }
    private Task<bool> ConfirmValidationAsync(DataValidationException error) =>
#if QUICKEXCEL_DIAGNOSTICS
        _validationChoice?.Invoke(error) ??
#endif
        Dialogs.Confirm(this, "数据验证", error.Message + "\n仍要提交整批输入吗？");
#if QUICKEXCEL_DIAGNOSTICS
    private Func<DataValidationException, Task<bool>>? _validationChoice;
#endif
    private void RequireVisiblePaste(GridRange target)
    {
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        CurrentView.RequireVisibleTargets(snapshot, SheetIndex, Enumerable.Range(target.Top, target.Rows));
        if (sheet.Rows.Any(p => p.Value.Hidden && p.Key >= target.Top && p.Key <= target.Bottom) ||
            sheet.Columns.Any(p => p.Value.Hidden && p.Key >= target.Left && p.Key <= target.Right))
            throw new InvalidOperationException("目标含隐藏行列，请先取消隐藏；本次没有写入。");
    }
    private async Task CopyAsync(bool cut, bool externalFormulas = false, bool valuesOnly = false)
    {
        if (!await CommitEditorsAsync()) return;
        if (_grid.SelectedDrawingObject is { } selected)
        {
            if (cut) throw new NotSupportedException("对象暂支持复制；粘贴确认后可删除原对象。");
            if (externalFormulas || valuesOnly) throw new NotSupportedException("对象请使用复制或复制为图片。");
            await CopyDrawingAsync(selected); return;
        }
        var view = CurrentView; view.RequireCurrentFilter(Document.Snapshot);
        var range = SelectedRange(); var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        AxisClipboard? axes = null;
        var wholeRow = SelectionPlanning.IsWholeRow(Rectangle(range)); var wholeColumn = SelectionPlanning.IsWholeColumn(Rectangle(range));
        if (!cut && wholeRow != wholeColumn && !externalFormulas && !valuesOnly)
        {
            if (sheet.Data.AllFilters.Any(f => f.Columns.Count > 0)) throw new InvalidOperationException("请先清除筛选条件，再复制完整行列。");
            axes = AxisClipboard.Capture(snapshot, sheet.Id, wholeRow ? SheetAxis.Row : SheetAxis.Column, wholeRow ? range.Top : range.Left, wholeRow ? range.Rows : range.Columns);
            var bounded = axes.Cells.Range; range = new(bounded.Top, bounded.Left, bounded.Bottom, bounded.Right);
        }
        bool Excluded(int row) => sheet.Rows.GetValueOrDefault(row)?.Hidden == true || view.IsFilteredOut(snapshot, SheetIndex, row);
        string? text = GridClipboard.Write(range, (r, c) => Document.GetCellValue(SheetIndex, r, c), Excluded);
        var plain = range.Count > WorkbookClipboard.MaxCells || Enumerable.Range(range.Top, range.Rows).Any(Excluded);
        if (cut && plain) throw new InvalidOperationException("剪切需要不含隐藏行且不超过 20,000 格的连续矩形。");
        if (cut && snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var capture = axes?.Cells ?? (plain ? null : WorkbookClipboard.Capture(snapshot, sheet.Id, Rectangle(range), cut));
        var token = plain || valuesOnly || externalFormulas ? null : Guid.NewGuid().ToString("N");
        byte[]? xml = null, html = null; string? limitation = null;
        if (capture is not null && OperatingSystem.IsWindows())
        {
            try
            {
                var region = SpreadsheetClipboardCodec.Capture(snapshot, capture, externalFormulas, valuesOnly);
                var encodedXml = SpreadsheetClipboardCodec.Write(region, externalFormulas);
                var encodedHtml = HtmlClipboardCodec.Write(region, externalFormulas);
                // Publish the external representations together only after every encoder succeeds.
                xml = encodedXml; html = encodedHtml;
            }
            catch (Exception ex) when (!externalFormulas && !valuesOnly && (ex is NotSupportedException || ex is InvalidOperationException && capture.HasOldResults))
            { limitation = ex.Message; text = null; }
        }
        if (externalFormulas && (plain || !OperatingSystem.IsWindows())) throw new NotSupportedException("公式对外复制需要 Windows 上不超过 20,000 格的完整可见矩形。");
        await _clipboard.WriteAsync(new(text, token, xml, html));
        _clipboardCells = capture; _clipboardToken = token; _clipboardAxes = axes;
        _clipboardSession = Document.Session.SessionId; _clipboardRevision = snapshot.Revision;
        _status.Text = limitation is not null ? (cut ? "已准备内部剪切，Esc 取消。" : "已内部复制。") + "未发布外部数据。" + limitation : cut ? "已准备内部剪切；跨应用粘贴仅复制，来源不会删除。Esc 取消。" :
            plain ? "已复制可见行显示文本。" : externalFormulas ? "已对外复制公式（相对引用）。" : valuesOnly ? "已复制类型化结果值。" : "已复制；对外粘贴结果值、格式和完整合并区域。";
    }
    private async Task PasteAsync(string mode)
    {
        if (!await CommitEditorsAsync()) return;
        var session = Document.Session.SessionId; var revision = Document.Snapshot.Revision; var target = SelectedRange(allowObject: true); var sheetId = Document.Snapshot.ActiveSheetId;
        var selectedObject = _grid.SelectedDrawingObject;
        var data = await _clipboard.ReadAsync();
        if (session != Document.Session.SessionId || revision != Document.Snapshot.Revision || sheetId != Document.Snapshot.ActiveSheetId || target != SelectedRange(allowObject: true) || selectedObject != _grid.SelectedDrawingObject)
            throw new InvalidOperationException("读取剪贴板期间文档、工作表或目标选区已改变，请重试。");
        if (SelectionPlanning.IsWholeRow(Rectangle(target)) || SelectionPlanning.IsWholeColumn(Rectangle(target)))
            target = new(target.Top, target.Left, target.Top, target.Left);
        if (data.Token != _clipboardToken || data.Token is null || mode is "image" or "chart")
        { await PasteAvailableClipboardAsync(data, mode, target, sheetId, revision); return; }
        if (data.Token is not null && data.Token == _clipboardToken)
        {
            if (_clipboardCells is not { } cells || _clipboardSession != session || _clipboardRevision != revision)
                throw new InvalidOperationException("复制/剪切已取消或文档已改变，请重新复制。");
            var paste = mode switch { "values" => CellPasteMode.Values, "formulas" => CellPasteMode.Formulas, "formats" => CellPasteMode.Formats, _ => CellPasteMode.All };
            if (paste != CellPasteMode.Formats && !cells.IsCut && cells.Rows == 1 && cells.Columns == 1 &&
                Document.Snapshot.Sheets[SheetIndex].MergedCells.Select(CellRectangle.Parse).Any(m => m == Rectangle(target)))
                target = new(target.Top, target.Left, target.Top, target.Left);
            var transpose = mode == "transpose"; var fill = cells.Rows == 1 && cells.Columns == 1 && !cells.IsCut;
            var rows = fill ? target.Rows : transpose ? cells.Columns : cells.Rows;
            var columns = fill ? target.Columns : transpose ? cells.Rows : cells.Columns;
            RequireVisiblePaste(new(target.Top, target.Left, target.Top + rows - 1, target.Left + columns - 1));
            try
            {
                if (!await ExecuteAsync(new PasteCellsCommand(cells, Document.Snapshot.ActiveSheetId, Rectangle(target), revision, paste, transpose), revision)) return;
            }
            catch (NotSupportedException ex) when (!cells.IsCut)
            {
                // Keep the internal capture's typed results; do not reparse formatted display strings.
                var partial = new ExternalClipboardRegion(cells.Rows, cells.Columns, cells.Cells.Select((c, i) =>
                    new ExternalClipboardCell(c.Value, Document.Snapshot.Styles[c.Content.StyleId],
                        Formula: c.Content.Kind == CellContentKind.Formula ? c.Content.FormulaText : null,
                        FormulaOrigin: new(cells.Range.Top + i / cells.Columns, cells.Range.Left + i % cells.Columns),
                        MissingResult: c.Content.Kind == CellContentKind.Formula && cells.HasOldResults)),
                    merges: cells.MergedCells.Select(m => new CellRectangle(m.Top - cells.Range.Top, m.Left - cells.Range.Left, m.Bottom - cells.Range.Top, m.Right - cells.Range.Left)),
                    limitations: ["部分内部结构未复制：" + ex.Message]);
                await PasteAvailableCellsAsync(partial, mode, target, sheetId, revision, []);
                _clipboardRevision = Document.Snapshot.Revision;
                return;
            }
            _clipboardRevision = Document.Snapshot.Revision; if (cells.IsCut) _clipboardCells = null;
            _grid.SelectRange(new(target.Top, target.Left, target.Top + rows - 1, target.Left + columns - 1), new(target.Top, target.Left));
        }
        await CalculatePendingAsync(); _status.Text = "已粘贴，可一次撤销。";
    }
    private async Task FillAsync(GridRange source, GridRange extent, long? revision = null)
    {
        CurrentView.RequireVisibleTargets(Document.Snapshot, SheetIndex, Enumerable.Range(extent.Top, extent.Rows));
        if (await ExecuteAsync(new FillCellsCommand(Document.Snapshot.ActiveSheetId, Rectangle(source), Rectangle(extent)), revision))
        { _grid.SelectRange(extent, new(extent.Top, extent.Left)); await CalculatePendingAsync(); }
    }
    private Task FillDirectionAsync(bool right)
    {
        var extent = SelectedRange();
        var source = right ? extent with { Right = extent.Left } : extent with { Bottom = extent.Top };
        return FillAsync(source, extent);
    }
    private async Task SeriesAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var range = SelectedRange();
        var dialog = new FieldsDialog("填充序列").Choice("kind", "序列类型", ["数字", "日期：天", "日期：月", "日期：年", "文本尾号"])
            .Text("step", "步长", "1");
        dialog.Validate = () => dialog.Number("step", -1e12, 1e12);
        if (!await dialog.ShowAsync(this)) return;
        RequireVisiblePaste(range);
        if (await ExecuteAsync(new FillSeriesCommand(Document.Snapshot.ActiveSheetId, Rectangle(range), (CellSeriesKind)dialog.Index("kind"), dialog.Number("step", -1e12, 1e12)), revision)) await CalculatePendingAsync();
    }
    private void InstallEditingCommands()
    {
        _clipboard = new AvaloniaClipboard(this);
        _grid.DimensionChanged += (axis, index, pixels, revision) => _ = PerformAsync(async () =>
            await ExecuteAsync(new SetAxisMetadataCommand(Document.Snapshot.ActiveSheetId, axis, index, 1, axis == SheetAxis.Row ? pixels : Math.Max(.01, (pixels - 5) / 7)), revision));
        _grid.FillRequested += (source, extent, revision) => _ = PerformAsync(() => FillAsync(source, extent, revision));
        var edit = (MenuItem)_menu.Items[1]!;
        edit.Items.Add(new Separator());
        edit.Items.Add(CommandMenu("复制  Ctrl+C", () => CopyAsync(false))); edit.Items.Add(CommandMenu("剪切  Ctrl+X", () => CopyAsync(true)));
        edit.Items.Add(CommandMenu("复制结果值（对外）", () => CopyAsync(false, valuesOnly: true)));
        edit.Items.Add(CommandMenu("复制公式（对外，相对引用）", () => CopyAsync(false, externalFormulas: true)));
        edit.Items.Add(CommandMenu("复制对象为图片", CopyDrawingPictureAsync));
        foreach (var (title, mode) in new[] { ("粘贴  Ctrl+V", "all"), ("只贴值", "values"), ("粘贴公式", "formulas"), ("只贴格式", "formats"), ("转置粘贴", "transpose"), ("粘贴为可编辑图表（附带数据快照）", "chart"), ("粘贴为图片", "image") })
            edit.Items.Add(CommandMenu(title, () => PasteAsync(mode)));
        edit.Items.Add(new Separator()); edit.Items.Add(CommandMenu("清除内容", ClearAsync)); edit.Items.Add(CommandMenu("清除格式", ClearFormatsAsync));
        edit.Items.Add(new Separator()); edit.Items.Add(CommandMenu("向下填充  Ctrl+D", () => FillDirectionAsync(false)));
        edit.Items.Add(CommandMenu("向右填充  Ctrl+R", () => FillDirectionAsync(true))); edit.Items.Add(CommandMenu("填充序列", SeriesAsync));
        InstallGridContextMenu();
    }
    private MenuItem CommandMenu(string title, Func<Task> action)
    { var item = new MenuItem { Header = title }; item.Click += async (_, _) => await PerformAsync(action); return item; }
}
