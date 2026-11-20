using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task ImportTextAsync(bool useTransferDestination = false)
    {
        if (!await CommitEditorsAsync()) return;
        var file = await _files.OpenAsync(new(UiText.Get("导入 CSV / TSV…"), [new(UiText.Get("分隔文本文件"), ["*.csv", "*.tsv", "*.txt"])]));
        if (file is null) return;
        var document = Document; var revision = document.Snapshot.Revision;
        var settings = NumericFields(new FieldsDialog(UiText.Get("文本导入设置"))
            .Choice("delimiter", UiText.Get("分隔符"), [UiText.Get("逗号（CSV）"), UiText.Get("制表符（TSV）"), UiText.Get("分号"), UiText.Get("竖线")], Path.GetExtension(file.Path).Equals(".tsv", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
            .Choice("encoding", UiText.Get("编码"), ["UTF-8", "UTF-16 LE", "UTF-16 BE", "GB18030"])
            .Choice("dates", UiText.Get("日期系统"), ["1900", "1904"])
            .Check("header", UiText.Get("首行保留为文本标题"), false));
        settings.Validate = () => ReadNumeric(settings);
        if (!await settings.ShowAsync(this)) return;
        var numeric = ReadNumeric(settings); var delimiter = new[] { ',', '\t', ';', '|' }[settings.Index("delimiter")];
        var encoding = (TextFileEncoding)settings.Index("encoding");
        var header = settings.Boolean("header");
        string[][] rows;
        using (var cancellation = new CancellationTokenSource())
        using (BeginExecution(UiText.Get("正在读取文本…"), cancellation))
            rows = await Task.Run(() => TextImportReader.Read(TextImportReader.ReadFile(file.Path, cancellation.Token), delimiter, encoding, cancellation.Token));
        var columns = rows.Length == 0 ? 0 : rows.Max(row => row.Length);
        if (columns == 0) throw new ArgumentException(UiText.Get("文本文件中没有可导入的数据。"));
        if (columns > WorkbookLimits.MaxColumns) throw new ArgumentException(UiText.Get("CSV 超出工作表行列上限。"));
        var types = new[] { CleaningColumnType.Text, CleaningColumnType.Number, CleaningColumnType.DateYmd, CleaningColumnType.DateMdy, CleaningColumnType.DateDmy, CleaningColumnType.Skip };
        var labels = new[] { UiText.Get("文本"), UiText.Get("数字"), UiText.Get("日期（年月日）"), UiText.Get("日期（月日年）"), UiText.Get("日期（日月年）"), UiText.Get("跳过此列") };
        var choices = new ComboBox[columns]; var list = new StackPanel { Spacing = 6 };
        for (var column = 0; column < columns; column++)
        {
            choices[column] = new ComboBox { ItemsSource = labels, SelectedIndex = 0, Width = 170 };
            list.Children.Add(new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, Children =
            { new TextBlock { Text = QuickExcel.Formulas.FormulaSyntax.ColumnName(column), Width = 40 }, choices[column],
              new TextBlock { Text = rows[0].ElementAtOrDefault(column), MaxWidth = 200, TextTrimming = TextTrimming.CharacterEllipsis } } });
        }
        var preview = new FieldsDialog(UiText.Get("文本导入预览"))
            .Note(useTransferDestination ? UiText.Format($"共 {rows.Length} 行、{columns} 列。默认文本保留前导零与等号原文；下一步选择导入目标。") : UiText.Format($"共 {rows.Length} 行、{columns} 列。默认文本保留前导零与等号原文；导入为新工作簿，原文件不修改。"))
            .Preview(new TextBox { Text = string.Join("\n", rows.Take(10).Select(row => string.Join(" | ", row.Select(v => v.Length > 80 ? v[..80] + "…" : v)))), IsReadOnly = true, AcceptsReturn = true, Height = 140 })
            .Preview(new ScrollViewer { Content = list, MaxHeight = 220 });
        if (!await preview.ShowAsync(this)) return;
        TextImportResult result;
        using (var cancellation = new CancellationTokenSource())
        using (BeginExecution(UiText.Get("正在转换文本…"), cancellation))
        {
            var selected = choices.Select(c => types[c.SelectedIndex]).ToArray();
            var dates = settings.Index("dates") == 0 ? ExcelDateSystem.Windows1900 : ExcelDateSystem.Mac1904;
            result = await Task.Run(() => TextImport.Convert(rows, selected, numeric, dates, header, cancellation.Token));
        }
        if (useTransferDestination)
        {
            await ApplyDataImportAsync(result.Workbook, result.Notices, document, revision);
            return;
        }
        var converted = new FieldsDialog(UiText.Get("确认导入"))
            .Note(string.Join("\n", result.Notices.Select(UiText.Get)))
            .Preview(new TextBox { Text = string.Join("\n", Enumerable.Range(0, Math.Min(10, rows.Length)).Select(r => string.Join(" | ", Enumerable.Range(0, Math.Min(columns, 20)).Select(c =>
                CellDisplayFormatter.Format(result.Workbook.Sheets[0].GetCell(r, c)?.Value ?? CellValue.Blank, result.Workbook.Styles[result.Workbook.Sheets[0].GetStyleId(r, c)], result.Workbook.DateSystem))))), IsReadOnly = true, AcceptsReturn = true, Height = 200 });
        if (!await converted.ShowAsync(this)) return;
        if (!ReferenceEquals(document, Document) || Document.Snapshot.Revision != revision) throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
        if (!await ConfirmDiscardAsync()) return;
        _workspace.Import(result.Workbook); ShowBook(); _status.Text = UiText.Get("文本已导入新工作簿，请另存。脚本未授权。");
    }
}
