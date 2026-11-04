using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    // Catch only while decoding a source representation, never around destination validation or commits.
    private static bool IsClipboardSourceFailure(Exception error) => error is
        NotSupportedException or InvalidDataException or FormatException or ArgumentException or
        System.Xml.XmlException or System.Text.Json.JsonException or OverflowException;

    private async Task PasteAvailableClipboardAsync(ClipboardExchange data, string mode, GridRange target, SheetId sheetId, long revision)
    {
        var notices = new List<string>();
        if (data.ReadNotice is not null) notices.Add(data.ReadNotice);
        var hasObjects = data.Png is not null || data.OfficeDrawing is not null || data.ChartPackage is not null;
        var objectsFirst = mode is "image" or "chart" || mode == "all" && data.SpreadsheetXml is null && data.Html is null && hasObjects;
        if (objectsFirst && await TryPasteDrawingAsync(data, mode, new(target.Top, target.Left), revision, notices)) return;
        ExternalClipboardRegion? source = null;
        if (data.SpreadsheetXml is { } xml)
            try { source = SpreadsheetClipboardCodec.Read(xml, data.Biff8); }
            catch (Exception ex) when (IsClipboardSourceFailure(ex)) { notices.Add("原始表格格式未采用：" + ex.Message); }
        if (source is null && data.Html is { } html)
            try { source = HtmlClipboardCodec.Read(html); }
            catch (Exception ex) when (IsClipboardSourceFailure(ex)) { notices.Add("HTML 表格格式未采用：" + ex.Message); }
        if (source is not null)
        {
            if (hasObjects) notices.Add("已保留单元格区域，区域附带的浮动对象未复制。");
            if (mode is "image" or "chart") notices.Add("所选对象形式不可用，已改贴单元格内容。");
            await PasteAvailableCellsAsync(source, mode, target, sheetId, revision, notices); return;
        }
        if (!objectsFirst && hasObjects && mode != "formats" && await TryPasteDrawingAsync(data, mode, new(target.Top, target.Left), revision, notices)) return;
        // Image-only selection may still salvage an editable chart if no image was supplied.
        if (objectsFirst && mode == "image" && (data.ChartPackage is not null || data.OfficeDrawing is not null))
        {
            notices.Add("图片形式不可用，已尝试可编辑图表。");
            if (await TryPasteDrawingAsync(data with { Png = null }, "chart", new(target.Top, target.Left), revision, notices)) return;
        }
        if (mode == "formats") { _status.Text = "剪贴板没有可用的单元格格式，目标内容保持不变。 " + string.Join(" ", notices); return; }
        var text = data.Text;
        if (text is null && data.Html is { } textHtml)
            try { text = HtmlClipboardCodec.ReadDisplayText(textHtml); }
            catch (Exception ex) when (IsClipboardSourceFailure(ex)) { notices.Add("HTML 文本未采用：" + ex.Message); }
        if (text is null || text.Length == 0 && (hasObjects || data.Html is not null))
        { _status.Text = "剪贴板没有可读取的单元格、图片或文本内容，目标保持不变。 " + string.Join(" ", notices); return; }
        string[][]? values = null;
        try { values = GridClipboard.Read(text); }
        catch (ArgumentException ex)
        {
            notices.Add(ex.Message);
            // An irregular small text selection still has a useful single-cell representation.
            if (text.Length <= WorkbookLimits.MaxCellTextLength) { values = [[text]]; notices.Add("非矩形文本已粘贴到单个单元格。"); }
        }
        if (values is null) { _status.Text = "没有可在当前大小限制内粘贴的内容，目标保持不变。 " + string.Join(" ", notices); return; }
        if (values.SelectMany(row => row).Any(value => value.Length > WorkbookLimits.MaxCellTextLength))
        { _status.Text = "文本超过单元格大小限制，目标保持不变。 " + string.Join(" ", notices); return; }
        if (notices.Count > 0 || hasObjects || mode is "image" or "chart" or "transpose")
        {
            notices.Add("已降级为文本值，保留目标格式。");
            var region = new ExternalClipboardRegion(values.Length, values[0].Length,
                values.SelectMany(row => row.Select(value => new ExternalClipboardCell(value.Length == 0 ? CellValue.Blank : CellValue.FromText(value)))));
            await PasteAvailableCellsAsync(region, mode == "transpose" ? "transpose" : "values", target, sheetId, revision, notices); return;
        }
        // Retain ordinary TSV's established numeric parsing and explicit formula-paste semantics.
        var sheet = Document.Snapshot.GetSheet(sheetId);
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(m => m == Rectangle(target)) && values is [{ Length: 1 }])
            target = new(target.Top, target.Left, target.Top, target.Left);
        var edits = GridClipboard.Prepare(text, target, mode == "formulas");
        RequireVisiblePaste(new(target.Top, target.Left, edits.Max(v => v.Row), edits.Max(v => v.Column)));
        if (!await ExecuteAsync(new SetCellsCommand(edits.Select(v => new CellEdit(new(sheetId, v.Row, v.Column),
            CellInput.ParseForCell(v.Value, Document.Snapshot, sheet, v.Row, v.Column)))), revision)) return;
        _grid.SelectRange(new(target.Top, target.Left, edits.Max(v => v.Row), edits.Max(v => v.Column)), new(target.Top, target.Left));
        await CalculateAsync(); _status.Text = "已粘贴，可一次撤销。";
    }

    private async Task PasteAvailableCellsAsync(ExternalClipboardRegion source, string mode, GridRange target, SheetId sheetId, long revision, List<string> notices)
    {
        if (mode == "transpose")
            source = new(source.Columns, source.Rows,
                Enumerable.Range(0, source.Columns).SelectMany(c => Enumerable.Range(0, source.Rows).Select(r => source.Cells[r * source.Columns + c])),
                source.Merges.Select(m => new CellRectangle(m.Left, m.Top, m.Right, m.Bottom)), source.Limitations, source.ValuesAreDisplayText, source.HasFilterMetadata);
        var requested = mode switch { "values" => CellPasteMode.Values, "formulas" => CellPasteMode.Formulas, "formats" => CellPasteMode.Formats, _ => CellPasteMode.All };
        var plan = ExternalClipboardPastePolicy.Prepare(source, requested);
        if (plan.Notice is not null) notices.Add(plan.Notice);
        if (plan.Mode != CellPasteMode.Formats && source.Rows == 1 && source.Columns == 1 &&
            Document.Snapshot.GetSheet(sheetId).MergedCells.Select(CellRectangle.Parse).Any(m => m == Rectangle(target)))
            target = new(target.Top, target.Left, target.Top, target.Left);
        var rows = source.Rows == 1 && source.Columns == 1 ? target.Rows : source.Rows;
        var columns = source.Rows == 1 && source.Columns == 1 ? target.Columns : source.Columns;
        RequireVisiblePaste(new(target.Top, target.Left, checked(target.Top + rows - 1), checked(target.Left + columns - 1)));
        try
        {
            if (!await ExecuteAsync(new PasteExternalCellsCommand(source, sheetId, Rectangle(target), revision, plan.Mode, allowPartial: true), revision)) return;
        }
        catch (NotSupportedException) when (plan.Mode == CellPasteMode.All)
        {
            // Preserve destination attachments/rules if copying formats cannot safely coexist with them.
            if (!await ExecuteAsync(new PasteExternalCellsCommand(source, sheetId, Rectangle(target), revision, CellPasteMode.Values, allowPartial: true), revision)) return;
            notices.Add("已自动改为只贴值，保留目标格式、附件和规则。");
        }
        catch (NotSupportedException ex) when (plan.Mode == CellPasteMode.Formats)
        { _status.Text = "没有可应用的格式，目标内容保持不变。 " + ex.Message; return; }
        _grid.SelectRange(new(target.Top, target.Left, target.Top + rows - 1, target.Left + columns - 1), new(target.Top, target.Left));
        await CalculateAsync();
        _status.Text = (Document.Snapshot.Revision == revision ? "没有可应用的更改，目标保持不变。 " : "已粘贴，可一次撤销。 ") + string.Join(" ", notices);
    }
}
