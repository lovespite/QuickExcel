using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Printing;
using Microsoft.Win32;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async void PageSettings_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var current = sheet.Printing; var setup = current.Setup ?? new();
        var dialog = new ReportFieldsDialog(this, "页面设置");
        var paperCodes = new uint[] { 9, 8, 11, 1, 5 };
        var paper = dialog.Choice("纸张", ["A4", "A3", "A5", "Letter", "Legal"], Math.Max(0, Array.IndexOf(paperCodes, setup.PaperSize ?? 9)));
        var orientation = dialog.Choice("方向", ["纵向", "横向"], setup.Orientation == "landscape" ? 1 : 0);
        var scale = dialog.Text("缩放百分比（10–400）", (setup.Scale ?? 100).ToString());
        var fit = dialog.Check("适应页数（0 表示该方向不限）", current.Properties?.FitToPage == true);
        var fitWidth = dialog.Text("适应宽度（页）", (setup.FitToWidth ?? 1).ToString()); var fitHeight = dialog.Text("适应高度（页）", (setup.FitToHeight ?? 0).ToString());
        var m = current.Margins ?? new(.7, .7, .75, .75, .3, .3);
        var margins = dialog.Text("页边距：左,右,上,下（厘米）", string.Join(",", new[] { m.Left, m.Right, m.Top, m.Bottom }.Select(v => (v * 2.54).ToString("0.###", CultureInfo.InvariantCulture))));
        var areas = dialog.Text("打印区域（如 A1:F30；留空为已用区域）", string.Join(";", current.Areas.Select(a => new CellRectangle(a.Top, a.Left, a.Bottom, a.Right).ToString())));
        var rows = current.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Rows); var cols = current.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Columns);
        var titleRows = dialog.Text("重复行（如 1:2）", rows is null ? "" : $"{rows.Top + 1}:{rows.Bottom + 1}");
        var titleCols = dialog.Text("重复列序号（如 1:2）", cols is null ? "" : $"{cols.Left + 1}:{cols.Right + 1}");
        var header = dialog.Text("页眉（&L 左 / &C 中 / &R 右；&P 页码 / &N 总页数）", current.HeaderFooter?.OddHeader);
        var footer = dialog.Text("页脚", current.HeaderFooter?.OddFooter);
        var grid = dialog.Check("打印网格线", current.Options?.GridLines == true); var headings = dialog.Check("打印行列标题", current.Options?.Headings == true);
        dialog.Apply = () =>
        {
            var values = margins.Text.Split(',').Select(v => double.Parse(v.Trim(), CultureInfo.InvariantCulture) / 2.54).ToArray();
            if (values.Length != 4) throw new ArgumentException("请输入四个页边距。");
            var titles = new List<PrintRange>();
            void TitleRange(string text, bool row)
            {
                if (string.IsNullOrWhiteSpace(text)) return;
                var pair = text.Split(':').Select(s => int.Parse(s.Trim()) - 1).ToArray();
                if (pair.Length != 2) throw new ArgumentException("重复标题格式应为 起始:结束。");
                titles.Add(row ? new(PrintRangeKind.Rows, pair[0], 0, pair[1], PrintRange.MaxColumns - 1) : new(PrintRangeKind.Columns, 0, pair[0], PrintRange.MaxRows - 1, pair[1]));
            }
            TitleRange(titleRows.Text, true); TitleRange(titleCols.Text, false);
            var printing = current with
            {
                Setup = setup with { PaperSize = paperCodes[paper.SelectedIndex], PaperHeight = null, PaperWidth = null, Orientation = orientation.SelectedIndex == 1 ? "landscape" : "portrait", Scale = uint.Parse(scale.Text), FitToWidth = uint.Parse(fitWidth.Text), FitToHeight = uint.Parse(fitHeight.Text) },
                Properties = (current.Properties ?? new()) with { FitToPage = fit.IsChecked == true },
                Margins = new(values[0], values[1], values[2], values[3], m.Header, m.Footer),
                Areas = [.. areas.Text.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(a => { var r = CellRectangle.Parse(a); return new PrintRange(PrintRangeKind.Cells, r.Top, r.Left, r.Bottom, r.Right); })],
                Titles = [.. titles], HeaderFooter = (current.HeaderFooter ?? new()) with { OddHeader = header.Text, OddFooter = footer.Text },
                Options = (current.Options ?? new()) with { GridLines = grid.IsChecked == true, Headings = headings.IsChecked == true }
            };
            ExecuteWorkbookCommand(new SetPrintingCommand(sheet.Id, printing));
        };
        dialog.ShowDialog();
    }

    private async void PrintPreview_Click(object sender, RoutedEventArgs e) =>
        await (_workspaceInput = ShowPrintPreviewAsync());

    private async Task ShowPrintPreviewAsync()
    {
        if (!await RecalculateWorkbookAsync()) return;
        try
        {
            var snapshot = _document.Snapshot;
            if (snapshot.Sheets[_sheetIndex].Cells.Values.Any(c => c.Content.Kind == CellContentKind.Formula && c.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException("重算后当前表仍含无法计算的公式，打印已取消；请查看兼容性报告或脚本诊断。");
            await using var paginator = new WorkbookPaginator(snapshot, snapshot.Sheets[_sheetIndex], _filePath ?? _displayName);
            if (paginator.Warnings.Count > 0)
            {
                AppendLog("打印内容提示：\n" + string.Join("\n", paginator.Warnings));
                new PrintWarningsDialog(this, paginator.Warnings).ShowDialog();
            }
            var preview = new Window { Owner = this, Title = $"打印预览 · {paginator.PageCount} 页", Width = 1000, Height = 760, WindowStartupLocation = WindowStartupLocation.CenterOwner };
            var root = new DockPanel(); var print = new Button { Content = "系统打印…", Padding = new(16, 8, 16, 8), Margin = new(8), HorizontalAlignment = HorizontalAlignment.Left };
            var pdf = new Button { Content = "导出 PDF…", Padding = new(16, 8, 16, 8), Margin = new(8) };
            var cancelPdf = new Button { Content = "取消导出", Padding = new(16, 8, 16, 8), Margin = new(8), Visibility = Visibility.Collapsed };
            CancellationTokenSource? exporting = null; cancelPdf.Click += (_, _) => exporting?.Cancel();
            var actions = new StackPanel { Orientation = Orientation.Horizontal }; actions.Children.Add(print); actions.Children.Add(pdf); actions.Children.Add(cancelPdf);
            DockPanel.SetDock(actions, Dock.Top); root.Children.Add(actions);
            var viewer = new DocumentViewer { Document = paginator.CreatePreviewDocument() };
            var thumbnails = paginator.CreateThumbnails(viewer); DockPanel.SetDock(thumbnails, Dock.Left); root.Children.Add(thumbnails); root.Children.Add(viewer); preview.Content = root;
            pdf.Click += async (_, _) =>
            {
                var save = new SaveFileDialog { Filter = "PDF 文档|*.pdf", FileName = Path.GetFileNameWithoutExtension(_displayName) + ".pdf" };
                if (save.ShowDialog(preview) != true) return;
                pdf.IsEnabled = false; cancelPdf.Visibility = Visibility.Visible;
                using var cancellation = new CancellationTokenSource(); exporting = cancellation;
                try { await PdfExport.WriteAsync(paginator, save.FileName, cancellation.Token); StatusText.Text = "PDF 已导出：" + save.FileName; }
                catch (OperationCanceledException) { StatusText.Text = "PDF 导出已取消，原文件未替换。"; }
                catch (Exception ex) { if (preview.IsVisible) MessageBox.Show(preview, ex.Message, "PDF 导出未完成", MessageBoxButton.OK, MessageBoxImage.Error); }
                finally { exporting = null; pdf.IsEnabled = true; cancelPdf.Visibility = Visibility.Collapsed; }
            };
            print.Click += (_, _) =>
            {
                try
                {
                    var dialog = new PrintDialog();
                    if (dialog.ShowDialog() != true) return;
                    dialog.PrintTicket.PageMediaSize = new PageMediaSize(paginator.PageSize.Width, paginator.PageSize.Height);
                    dialog.PrintTicket.PageOrientation = PageOrientation.Portrait;
                    if (snapshot.Sheets[_sheetIndex].Printing.Setup?.BlackAndWhite == true) dialog.PrintTicket.OutputColor = OutputColor.Monochrome;
                    dialog.PrintDocument(paginator.CreatePrinterPaginator(), _displayName);
                }
                catch (Exception ex) { MessageBox.Show(preview, ex.Message, "打印未完成", MessageBoxButton.OK, MessageBoxImage.Error); }
            };
            preview.ShowDialog();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void ImageManage_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            var choose = new ReportFieldsDialog(this, "图片：插入、移动、缩放和删除");
            var list = choose.Choice("图片", new[] { "插入新图片…" }.Concat(sheet.Images.Select((_, i) => $"图片 {i + 1}")).ToArray(), 0);
            var remove = choose.Check("删除所选图片", false);
            choose.Apply = () =>
            {
                var index = list.SelectedIndex - 1;
                if (remove.IsChecked == true)
                {
                    if (index < 0) throw new ArgumentException("请先选择已有图片。");
                    ExecuteWorkbookCommand(new SetImageCommand(sheet.Id, index, null)); return;
                }
                SheetImage image;
                if (index < 0)
                {
                    var file = new OpenFileDialog { Filter = "图片|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff" };
                    if (file.ShowDialog(this) != true) return;
                    if (new FileInfo(file.FileName).Length > 16 * 1024 * 1024) throw new ArgumentException("单张图片最多 16 MiB。");
                    image = new SheetImage(File.ReadAllBytes(file.FileName), Path.GetExtension(file.FileName), _selectedRow, _selectedColumn);
                    index = sheet.Images.Length;
                }
                else image = sheet.Images[index];
                var bounds = image.Anchor.Bounds(sheet); var edit = new ReportFieldsDialog(this, "图片位置和尺寸（96 DPI 像素）");
                var left = edit.Text("距工作表左边", bounds.X.ToString("0.###", CultureInfo.InvariantCulture)); var top = edit.Text("距工作表顶部", bounds.Y.ToString("0.###", CultureInfo.InvariantCulture));
                var width = edit.Text("宽度", bounds.Width.ToString("0.###", CultureInfo.InvariantCulture)); var height = edit.Text("高度", bounds.Height.ToString("0.###", CultureInfo.InvariantCulture));
                var behavior = edit.Choice("随单元格变化", ["移动并缩放", "移动但不缩放", "固定位置和尺寸"], (int)image.Anchor.Kind);
                edit.Apply = () => ExecuteWorkbookCommand(new SetImageCommand(sheet.Id, index, image.WithAnchor(DrawingGeometry.FromBounds(sheet,
                    new(double.Parse(left.Text, CultureInfo.InvariantCulture), double.Parse(top.Text, CultureInfo.InvariantCulture), double.Parse(width.Text, CultureInfo.InvariantCulture), double.Parse(height.Text, CultureInfo.InvariantCulture)), (DrawingAnchorKind)behavior.SelectedIndex))));
                edit.ShowDialog();
            };
            choose.ShowDialog();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
