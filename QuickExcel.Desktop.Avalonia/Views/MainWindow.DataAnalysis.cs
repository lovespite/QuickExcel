using System.Globalization;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private DropDownButton DataToolDropdown(string title, params Control[] items)
    {
        var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        content.Children.Add(RibbonIcons.Create(RibbonIcon.Default)); content.Children.Add(new TextBlock { Text = title, VerticalAlignment = VerticalAlignment.Center });
        var button = new DropDownButton { Content = content, Height = 28, MinHeight = 28, Padding = new(6, 3), Margin = new(2, 1),
            Background = Brushes.Transparent, BorderThickness = new(0), HorizontalAlignment = HorizontalAlignment.Left,
            Flyout = new MenuFlyout { ItemsSource = items } };
        button.Classes.Add("ribbon-command"); AutomationProperties.SetName(button, title); ToolTip.SetTip(button, title); _mutationControls.Add(button); return button;
    }

    private static string DataColumnLabel(WorksheetSnapshot sheet, int row, int column)
    {
        var label = CellInput.DisplayText(sheet.GetCell(row, column)?.Value ?? CellValue.Blank);
        if (label.Length > 60) label = label[..60] + "…";
        return string.IsNullOrEmpty(label) ? QuickExcel.Formulas.FormulaSyntax.ColumnName(column) : $"{label} ({QuickExcel.Formulas.FormulaSyntax.ColumnName(column)})";
    }
    private CellRectangle DataToolRange(bool singleColumn = false)
    {
        if (_grid.SelectionRanges.Count != 1) throw new ArgumentException("请选择一个连续的数据区域。");
        var selected = Rectangle(SelectedRange()); var sheet = Document.Snapshot.Sheets[SheetIndex];
        if (singleColumn && selected.Columns != 1) throw new ArgumentException("分列来源须为一列，请只选择一个来源列。");
        var range = selected.Count == 1 ? SortPlanning.CurrentRegion(sheet, selected.Top, selected.Left) :
            SelectionPlanning.IsWholeColumn(selected) || SelectionPlanning.IsWholeRow(selected) ? SortPlanning.WholeAxisRegion(sheet, selected, _grid.Current) : selected;
        if (singleColumn) range = range with { Left = selected.Left, Right = selected.Left };
        if (range.Count > DataCleaning.MaxCells) throw new ArgumentException("数据工具一次最多处理 20,000 格，请缩小区域。");
        return range;
    }

    private async Task FlashFillAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var target = _grid.Current.Column;
        var selected = Rectangle(SelectedRange()); var range = DataToolRange();
        // A blank target beside a filled source still belongs to the source's record block.
        if (range.Count == 1)
        {
            var adjacent = target > 0 ? SortPlanning.CurrentRegion(sheet, _grid.Current.Row, target - 1) : range;
            if (adjacent.Count == 1 && target + 1 < WorkbookLimits.MaxColumns) adjacent = SortPlanning.CurrentRegion(sheet, _grid.Current.Row, target + 1);
            range = adjacent with { Left = Math.Min(adjacent.Left, target), Right = Math.Max(adjacent.Right, target) };
        }
        else if (range.Columns == 1) range = SortRangeCommand.ExpandColumns(sheet, range, false);
        // A selected example/blank cell infers a header only when the target header itself is empty.
        if (selected.Count == 1 && range.Rows > 2 && range.Top < _grid.Current.Row && sheet.GetCell(range.Top, target) is null &&
            Enumerable.Range(range.Left, range.Columns).Where(c => c != target).All(c => sheet.GetCell(range.Top, c)?.Value.Kind == CellValueKind.Text)) range = range with { Top = range.Top + 1 };
        var dialog = new FieldsDialog("快速填充").Note("在目标列先输入示例。支持文本提取、拼接、大小写转换；仅填充空白，先预览。")
            .Text("range", "来源与示例区域（不含标题）", range.ToString()).Text("target", "目标列", QuickExcel.Formulas.FormulaSyntax.ColumnName(target));
        dialog.Validate = () => { _ = CellRectangle.Parse(dialog.String("range")); if (ColumnNumbers(dialog.String("target")).Length != 1) throw new ArgumentException("请选择一个目标列。"); };
        if (!await dialog.ShowAsync(this)) return;
        range = CellRectangle.Parse(dialog.String("range")); target = ColumnNumbers(dialog.String("target"))[0];
        var preview = await Task.Run(() => FlashFill.Preview(snapshot, sheet.Id, range, target)); await ApplyCleaningAsync(preview);
    }

    private static CellPosition DataPosition(string text)
    { var r = CellRectangle.Parse(text.Trim()); if (r.Count != 1) throw new ArgumentException("请输入当前工作表的单个单元格地址，例如 B2。"); return new(r.Top, r.Left); }
    private static CellAddress DataAddress(SheetId sheet, string text) { var p = DataPosition(text); return new(sheet, p.Row, p.Column); }
    private string CurrentAddress => new CellRectangle(_grid.Current.Row, _grid.Current.Column, _grid.Current.Row, _grid.Current.Column).ToString();
    private static double[] AnalysisNumbers(string text)
    {
        var parts = text.Split([',', '，', ';', '；', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length is < 1 or > 200) throw new ArgumentException("请输入 1–200 个数值，以逗号或空格分隔。");
        return parts.Select(s => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? n : throw new ArgumentException("无效数值：" + s)).ToArray();
    }
    private async Task<T> AnalyzeAsync<T>(string title, Func<CancellationToken, T> action)
    {
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        using var feedback = BeginExecution(title, cancellation);
        return await Task.Run(() => action(cancellation.Token), cancellation.Token);
    }

    private async Task GoalSeekAsync()
    {
        if (!await FlushEditsAsync()) return; var snapshot = Document.Snapshot; var sheet = snapshot.ActiveSheetId;
        var dialog = new FieldsDialog("单变量求解").Note("通过改变一个数值单元格，使公式达到目标。源数据在确认前保持不变。")
            .Text("formula", "目标公式单元格", CurrentAddress).Text("target", "目标值", "0").Text("input", "可变单元格", "A1").Text("tolerance", "允许误差", "0.000001");
        dialog.Validate = () => { _ = DataPosition(dialog.String("formula")); _ = DataPosition(dialog.String("input")); dialog.Number("target", -1e100, 1e100); dialog.Number("tolerance", 1e-12, 1); };
        if (!await dialog.ShowAsync(this)) return;
        var formula = DataAddress(sheet, dialog.String("formula")); var input = DataAddress(sheet, dialog.String("input")); var target = dialog.Number("target", -1e100, 1e100); var tolerance = dialog.Number("tolerance", 1e-12, 1);
        var result = await AnalyzeAsync("正在单变量求解…", c => DataAnalysis.GoalSeek(snapshot, _services.Calculator, formula, input, target, tolerance, c));
        await ApplyCleaningAsync(result.Preview);
    }
    private async Task DataTableAsync()
    {
        if (!await FlushEditsAsync()) return; var snapshot = Document.Snapshot; var sheet = snapshot.ActiveSheetId;
        var dialog = new FieldsDialog("模拟运算表").Note("用内置公式试算一组或两组输入。本次打开期间可刷新，文件保存结果值。")
            .Text("formula", "结果公式单元格", CurrentAddress).Text("column", "列输入单元格", "A1").Text("columns", "列输入的试算值", "1,2,3,4,5")
            .Text("row", "行输入单元格（二变量；一变量留空）").Text("rows", "行输入的试算值（二变量）").Text("output", "输出左上角", "D1");
        dialog.Validate = () => { _ = DataPosition(dialog.String("formula")); _ = DataPosition(dialog.String("column")); _ = DataPosition(dialog.String("output")); _ = AnalysisNumbers(dialog.String("columns"));
            if (!string.IsNullOrWhiteSpace(dialog.String("row"))) { _ = DataPosition(dialog.String("row")); _ = AnalysisNumbers(dialog.String("rows")); } };
        if (!await dialog.ShowAsync(this)) return;
        var formula = DataAddress(sheet, dialog.String("formula")); var column = DataAddress(sheet, dialog.String("column")); var columns = AnalysisNumbers(dialog.String("columns"));
        CellAddress? row = string.IsNullOrWhiteSpace(dialog.String("row")) ? null : DataAddress(sheet, dialog.String("row")); var rows = row is null ? [] : AnalysisNumbers(dialog.String("rows")); var output = DataPosition(dialog.String("output"));
        var table = new SessionDataTable(formula, column, columns, row, rows, output);
        var preview = await AnalyzeAsync("正在计算模拟运算表…", c => table.Preview(snapshot, _services.Calculator, c));
        if (await ApplyCleaningAsync(preview))
        {
            _dataTables.RemoveAll(table.Overlaps);
            _dataTables.Add(table);
        }
    }
    private async Task ScenariosAsync()
    {
        if (!await FlushEditsAsync()) return; var snapshot = Document.Snapshot; var sheet = snapshot.ActiveSheetId;
        var dialog = new FieldsDialog("方案分析").Note("每行一个命名方案：名称;数值1;数值2…；输出各方案的输入和计算结果，可保存为普通单元格。")
            .Text("formula", "结果公式单元格", CurrentAddress).Text("inputs", "可变单元格（逗号分隔）", "A1,B1")
            .Text("scenarios", "方案（分号分隔，数值用小数点）", "基准;10;20\n增长;15;25", true).Text("output", "摘要输出左上角", "D1");
        CellAddress[] Inputs() => dialog.String("inputs").Split([',', '，'], StringSplitOptions.RemoveEmptyEntries).Select(s => DataAddress(sheet, s)).ToArray();
        AnalysisScenario[] Scenarios() => dialog.String("scenarios").Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries).Select(line =>
        { var parts = line.Split([';', '；']); if (parts.Length < 2) throw new ArgumentException("方案须包含名称与至少一个数值。"); return new AnalysisScenario(parts[0].Trim(), parts.Skip(1).SelectMany(AnalysisNumbers).ToArray()); }).ToArray();
        dialog.Validate = () => { _ = DataPosition(dialog.String("formula")); _ = DataPosition(dialog.String("output")); var count = Inputs().Length; if (Scenarios().Any(s => s.Values.Count != count)) throw new ArgumentException("各方案的数值个数必须等于可变单元格个数。"); };
        if (!await dialog.ShowAsync(this)) return;
        var formula = DataAddress(sheet, dialog.String("formula")); var inputs = Inputs(); var scenarios = Scenarios(); var output = DataPosition(dialog.String("output"));
        var preview = await AnalyzeAsync("正在计算方案摘要…", c => DataAnalysis.Scenarios(snapshot, _services.Calculator, formula, inputs, scenarios, sheet, output, c)); await ApplyCleaningAsync(preview);
    }
    private async Task ForecastAsync()
    {
        if (!await FlushEditsAsync()) return; var snapshot = Document.Snapshot; var range = DataToolRange();
        var dialog = new FieldsDialog("预测").Note("选择两列：递增、等间隔的日期/时间或序号，以及数值。生成历史与未来预测表。")
            .Text("range", "历史数据区域", range.ToString()).Check("header", "首行为标题", true).Text("periods", "预测期数", "5")
            .Choice("method", "预测方法", ["线性趋势", "Holt 双指数平滑（无季节性）"]).Text("alpha", "水平平滑参数 α", "0.5").Text("beta", "趋势平滑参数 β", "0.3")
            .Text("output", "输出左上角", QuickExcel.Formulas.FormulaSyntax.ColumnName(Math.Min(range.Right + 2, WorkbookLimits.MaxColumns - 3)) + (range.Top + 1));
        dialog.Validate = () => { _ = CellRectangle.Parse(dialog.String("range")); _ = DataPosition(dialog.String("output")); dialog.Integer("periods", 1, 1000); dialog.Number("alpha", 0.000001, 1); dialog.Number("beta", 0.000001, 1); };
        if (!await dialog.ShowAsync(this)) return;
        range = CellRectangle.Parse(dialog.String("range")); var output = DataPosition(dialog.String("output")); var periods = dialog.Integer("periods", 1, 1000); var header = dialog.Boolean("header");
        var method = (ForecastMethod)dialog.Index("method"); var alpha = dialog.Number("alpha", 0.000001, 1); var beta = dialog.Number("beta", 0.000001, 1);
        var preview = await Task.Run(() => DataAnalysis.Forecast(snapshot, snapshot.ActiveSheetId, range, header, periods, output, method, alpha, beta)); await ApplyCleaningAsync(preview);
    }
}
