using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private const string InternalClipboardFormat = "QuickExcel.WorkbookClipboard.Token.v1";
    private WorkbookClipboard? _workbookClipboard;
    private string? _clipboardMarker;
    private Guid _clipboardSession;
    private long _clipboardRevision;

    private void CaptureWorkbookClipboard(GridRange range, string text, bool cut)
    {
        if (cut && _document.Snapshot.FileState?.FormulaReadOnly == true)
            throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var payload = WorkbookClipboard.Capture(_document.Snapshot, _document.Snapshot.Sheets[_sheetIndex].Id,
            new(range.Top, range.Left, range.Bottom, range.Right), cut);
        var marker = Guid.NewGuid().ToString("N");
        var data = new DataObject(); data.SetData(DataFormats.UnicodeText, text); data.SetData(InternalClipboardFormat, marker);
        Clipboard.SetDataObject(data, true);
        _workbookClipboard = payload; _clipboardMarker = marker;
        _clipboardSession = _document.Session.SessionId; _clipboardRevision = _document.Snapshot.Revision;
        StatusText.Text = cut ? "已准备剪切；粘贴后才移动，可按 Esc 取消。" :
            "已复制值、公式和格式；粘贴到本工作簿时保留类型，公式引用随位置调整。";
    }

    private async Task PasteClipboardAsync(string mode)
    {
        if (_busy || !CanCommitComposition()) return;
        // The OS carries only a random marker and TSV. Never deserialize objects or Excel HTML.
        var marker = Clipboard.GetData(InternalClipboardFormat) as string;
        if (marker is not null && marker == _clipboardMarker)
        {
            var payload = _workbookClipboard;
            if (payload is null || _clipboardSession != _document.Session.SessionId || _clipboardRevision != _document.Snapshot.Revision)
                throw new InvalidOperationException("复制/剪切已取消或文档已改变，请重新复制；本次没有写入。");
            var pasteMode = mode switch { "all" or "transpose" => CellPasteMode.All, "values" => CellPasteMode.Values,
                "formulas" => CellPasteMode.Formulas, "formats" => CellPasteMode.Formats, _ => throw new ArgumentException("未知粘贴方式。") };
            var transpose = mode == "transpose";
            var selection = GridSurface.Selection;
            var range = selection.Count == 0 ? new GridRange(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn) : GridRange.Rectangle(selection);
            var fill = payload.Rows == 1 && payload.Columns == 1 && !payload.IsCut;
            var rows = fill ? range.Rows : transpose ? payload.Columns : payload.Rows;
            var columns = fill ? range.Columns : transpose ? payload.Rows : payload.Columns;
            if ((long)range.Top + rows > WorkbookLimits.MaxRows || (long)range.Left + columns > WorkbookLimits.MaxColumns)
                throw new ArgumentException($"粘贴超出界面 {WorkbookLimits.MaxRows:N0} 行 × {WorkbookLimits.MaxColumns:N0} 列范围。");
            if (!CommitEditors()) return;
            CurrentView.RequireVisibleTargets(_document.Snapshot, _sheetIndex, Enumerable.Range(range.Top, rows));
            var sourceIndex = _document.Snapshot.Sheets.IndexOf(_document.Snapshot.GetSheet(payload.SheetId));
            if (_sheetViews.TryGetValue(payload.SheetId, out var sourceView))
                sourceView.RequireVisibleTargets(_document.Snapshot, sourceIndex, Enumerable.Range(payload.Range.Top, payload.Rows));
            var result = _editor.Execute(new PasteCellsCommand(payload, _document.Snapshot.Sheets[_sheetIndex].Id,
                new(range.Top, range.Left, range.Bottom, range.Right), _clipboardRevision, pasteMode, transpose));
            _clipboardRevision = result.Revision;
            if (payload.IsCut) _workbookClipboard = null;
            _calculationPending |= result.Changed && !result.CalculationUnchanged;
            RefreshSheet(true); UpdateChrome();
            GridSurface.SelectRange(new(range.Top, range.Left, range.Top + rows - 1, range.Left + columns - 1), new(range.Top, range.Left));
            if (_calculationPending) await RecalculateAsync();
            if (!result.Changed) { StatusText.Text = "目标内容与粘贴结果相同，未增加撤销记录。"; return; }
            StatusText.Text = (payload.IsCut ? "已移动选区" : mode switch { "values" => "已只贴值", "formulas" => "已粘贴公式和常量（保留目标格式）",
                "formats" => "已只贴格式", "transpose" => "已转置粘贴", _ => "已粘贴值、公式和格式" }) + "，可一次撤销。" +
                (pasteMode == CellPasteMode.Values && payload.HasOldResults ? "使用的是来源缓存，未宣称它已重算。" : "");
            return;
        }
        if (mode is "formats" or "transpose") throw new InvalidOperationException("只贴格式和转置需要本工作簿的内部复制；外部剪贴板目前仅支持文本和值/显式公式解析。");
        if (Clipboard.ContainsText()) await PasteSelectionAsync(Clipboard.GetText(), mode == "formulas");
    }

    private void RibbonCut_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition()) return;
        try { CopySelection(cut: true); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
