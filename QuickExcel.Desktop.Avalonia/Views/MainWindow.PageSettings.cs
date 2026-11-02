using System.Collections.Immutable;
using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task PageSettingsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var printing = sheet.Printing; var setup = printing.Setup ?? new();
        uint[] codes = [9, 8, 11, 1, 5]; var margins = printing.Margins ?? new(.7, .7, .75, .75, .3, .3);
        var oldAreas = string.Join(';', printing.Areas.Select(a => new CellRectangle(a.Top, a.Left, a.Bottom, a.Right).ToString()));
        var rows = printing.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Rows); var columns = printing.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Columns);
        var dialog = new FieldsDialog("页面设置").Choice("paper", "纸张", ["A4", "A3", "A5", "Letter", "Legal"], Math.Max(0, Array.IndexOf(codes, setup.PaperSize ?? 9)))
            .Choice("orientation", "方向", ["纵向", "横向"], setup.Orientation == "landscape" ? 1 : 0).Text("scale", "缩放比例（10–400%）", (setup.Scale ?? 100).ToString())
            .Check("fit", "适应页数（0 表示该方向不限）", printing.Properties?.FitToPage == true).Text("fitWidth", "适应宽度（页）", (setup.FitToWidth ?? 1).ToString()).Text("fitHeight", "适应高度（页）", (setup.FitToHeight ?? 0).ToString())
            .Text("margins", "页边距：左,右,上,下（厘米，使用小数点）", string.Join(',', new[] { margins.Left, margins.Right, margins.Top, margins.Bottom }.Select(v => (v * 2.54).ToString("0.###", CultureInfo.InvariantCulture))))
            .Text("areas", "打印区域（分号分隔；留空使用已用区域）", oldAreas)
            .Text("rows", "重复行（例如 1:2）", rows is null ? "" : $"{rows.Top + 1}:{rows.Bottom + 1}").Text("columns", "重复列序号（例如 1:2）", columns is null ? "" : $"{columns.Left + 1}:{columns.Right + 1}")
            .Text("header", "页眉（&L 左 / &C 中 / &R 右；&P 页码 / &N 总页数）", printing.HeaderFooter?.OddHeader ?? "")
            .Text("footer", "页脚", printing.HeaderFooter?.OddFooter ?? "").Check("grid", "打印网格线", printing.Options?.GridLines == true).Check("headings", "打印行列标题", printing.Options?.Headings == true);
        WorksheetPrinting Read()
        {
            var values = dialog.String("margins").Split(',').Select(v => double.Parse(v, CultureInfo.InvariantCulture) / 2.54).ToArray();
            if (values.Length != 4 || values.Any(v => !double.IsFinite(v) || v < 0)) throw new ArgumentException("请输入四个非负页边距。");
            var titles = new List<PrintRange>();
            void Titles(string text, bool row)
            {
                if (string.IsNullOrWhiteSpace(text)) return;
                var pair = text.Split(':').Select(v => int.Parse(v.Trim()) - 1).ToArray(); if (pair.Length != 2) throw new ArgumentException("重复标题格式应为 起始:结束。");
                titles.Add(row ? new(PrintRangeKind.Rows, pair[0], 0, pair[1], PrintRange.MaxColumns - 1) : new(PrintRangeKind.Columns, 0, pair[0], PrintRange.MaxRows - 1, pair[1]));
            }
            Titles(dialog.String("rows"), true); Titles(dialog.String("columns"), false);
            return printing with
            {
                Setup = setup with { PaperSize = codes[dialog.Index("paper")], PaperHeight = null, PaperWidth = null, Orientation = dialog.Index("orientation") == 1 ? "landscape" : "portrait", Scale = (uint)dialog.Integer("scale", 10, 400), FitToWidth = (uint)dialog.Integer("fitWidth", 0, 32767), FitToHeight = (uint)dialog.Integer("fitHeight", 0, 32767) },
                Properties = (printing.Properties ?? new()) with { FitToPage = dialog.Boolean("fit") }, Margins = new(values[0], values[1], values[2], values[3], margins.Header, margins.Footer),
                Areas = dialog.String("areas") == oldAreas ? printing.Areas : dialog.String("areas").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(v => { var r = CellRectangle.Parse(v); return new PrintRange(PrintRangeKind.Cells, r.Top, r.Left, r.Bottom, r.Right); }).ToImmutableArray(),
                Titles = titles.ToImmutableArray(), HeaderFooter = (printing.HeaderFooter ?? new()) with { OddHeader = dialog.String("header"), OddFooter = dialog.String("footer") },
                Options = (printing.Options ?? new()) with { GridLines = dialog.Boolean("grid"), Headings = dialog.Boolean("headings") }
            };
        }
        dialog.Validate = () => sheet.WithPrinting(Read());
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetPrintingCommand(sheet.Id, Read()), snapshot.Revision);
    }
    private async Task PrintAreaAsync(bool clear)
    {
        if (!await CommitEditorsAsync()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex]; var range = SelectedRange();
        await ExecuteAsync(new SetPrintingCommand(sheet.Id, sheet.Printing with { Areas = clear ? [] : [new(PrintRangeKind.Cells, range.Top, range.Left, range.Bottom, range.Right)] }));
    }
    private async Task PageBreaksAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var cell = _grid.Current;
        var dialog = new FieldsDialog("分页").Choice("axis", "方向", ["当前行前", "当前列前"]).Choice("action", "操作", ["插入手动分页", "删除当前位置分页", "清除全部手动分页"])
            .Note($"当前有 {sheet.Printing.RowBreaks.Length} 个行分页、{sheet.Printing.ColumnBreaks.Length} 个列分页。");
        if (!await dialog.ShowAsync(this)) return;
        var printing = sheet.Printing;
        if (dialog.Index("action") == 2) printing = printing with { RowBreaks = printing.RowBreaks.Where(b => !b.Manual).ToImmutableArray(), ColumnBreaks = printing.ColumnBreaks.Where(b => !b.Manual).ToImmutableArray() };
        else
        {
            var row = dialog.Index("axis") == 0; var index = row ? cell.Row : cell.Column;
            if (index == 0) throw new ArgumentException("首行或首列前无需分页。");
            var breaks = (row ? printing.RowBreaks : printing.ColumnBreaks).Where(b => b.Id != index).ToList();
            if (dialog.Index("action") == 0) breaks.Add(new(index, 0, (row ? PrintRange.MaxColumns : PrintRange.MaxRows) - 1, Manual: true));
            printing = row ? printing with { RowBreaks = breaks.OrderBy(b => b.Id).ToImmutableArray() } : printing with { ColumnBreaks = breaks.OrderBy(b => b.Id).ToImmutableArray() };
        }
        await ExecuteAsync(new SetPrintingCommand(sheet.Id, printing), snapshot.Revision);
    }
}
