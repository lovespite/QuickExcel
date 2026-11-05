using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallArithmeticPaste()
    {
        var item = new MenuItem { Header = UiText.Get("运算粘贴…") };
        item.Bind(IsEnabledProperty, _book.GetObservable(IsVisibleProperty));
        item.Click += async (_, _) => { if (_book.IsVisible) await PerformAsync(ArithmeticPasteAsync); };
        ((MenuItem)_menu.Items[1]!).Items.Add(item);
    }
    private async Task ArithmeticPasteAsync()
    {
        Document.Session.RequireWritable(); if (!await FlushEditsAsync()) return;
        DismissPasteOptions();
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var selection = SelectedRange(); var id = snapshot.ActiveSheetId;
        void RequireCurrent()
        {
            if (session != Document.Session.SessionId || snapshot.Revision != Document.Snapshot.Revision || id != Document.Snapshot.ActiveSheetId ||
                selection != SelectedRange() || !ReferenceEquals(snapshot.GetSheet(id), Document.Snapshot.GetSheet(id)))
                throw new InvalidOperationException(UiText.Get("文档或目标选区已改变，请重新粘贴。"));
        }
        var data = await _clipboard.ReadAsync(); RequireCurrent(); ExternalClipboardRegion source;
        if (data.Token is not null && data.Token == _clipboardToken)
        {
            if (_clipboardCells is not { } cells || _clipboardSession != session || _clipboardRevision != snapshot.Revision || cells.IsCut || cells.HasOldResults)
                throw new InvalidOperationException(UiText.Get("运算粘贴需要有效的复制结果，不能使用剪切或过期内容；请重算并重新复制。"));
            source = new(cells.Rows, cells.Columns, cells.Cells.Select(c => new ExternalClipboardCell(c.Value, Formula: c.Content.Kind == CellContentKind.Formula ? c.Content.FormulaText : null)),
                merges: cells.MergedCells.Select(m => new CellRectangle(m.Top - cells.Range.Top, m.Left - cells.Range.Left, m.Bottom - cells.Range.Top, m.Right - cells.Range.Left)));
        }
        else if (_visibleCapture is { } visible && data.Token == visible.Token)
        {
            if (visible.Session != session || visible.Revision != snapshot.Revision)
                throw new InvalidOperationException(UiText.Get("复制后文档或筛选已改变，请重新复制可见单元格。"));
            source = visible.Data;
        }
        else if (data.SpreadsheetXml is { } xml) source = SpreadsheetClipboardCodec.Read(xml, data.Biff8);
        else if (data.Html is not null) throw new NotSupportedException(UiText.Get("运算粘贴需要原始数值，请从 Excel 复制表格或复制使用小数点的数字文本。"));
        else
        {
            var rows = GridClipboard.Read(data.Text ?? throw new InvalidOperationException(UiText.Get("剪贴板没有可用的数字。")));
            CellValue Parse(string text) => text.Length == 0 ? CellValue.Blank :
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number) ? CellValue.FromNumber(number) :
                throw new InvalidOperationException(UiText.Get("数字文本必须使用小数点，不含千位分隔符；整批未修改。"));
            source = new(rows.Length, rows[0].Length, rows.SelectMany(r => r).Select(s => new ExternalClipboardCell(Parse(s))));
        }
        var output = ArithmeticPasteCommand.OutputRange(source, Rectangle(selection));
        var dialog = new FieldsDialog(UiText.Get("运算粘贴"))
            .Note(UiText.Format($"将修改 {output}，共 {output.Count:N0} 格，保留目标格式和附件。源公式使用结果值；目标公式保留并追加运算。日期按数值序列号运算。"))
            .Note(UiText.Get("空白按 0 运算；勾选跳过空白时不修改对应目标。文本、布尔、错误、除零或溢出使整批取消。隐藏、筛除及数组目标不支持。整批可一次撤销。"))
            .Choice("operation", UiText.Get("运算"), [UiText.Get("加"), UiText.Get("减"), UiText.Get("乘"), UiText.Get("除")])
            .Check("skip", UiText.Get("跳过源空白"), true);
        if (!await dialog.ShowAsync(this)) return; RequireCurrent();
        if (await ExecuteAsync(new ArithmeticPasteCommand(source, id, Rectangle(selection), snapshot.Revision, (PasteArithmetic)dialog.Index("operation"), dialog.Boolean("skip")), snapshot.Revision))
        { await CalculatePendingAsync(); _grid.SelectRange(new(output.Top, output.Left, output.Bottom, output.Right), new(output.Top, output.Left)); }
    }
}
