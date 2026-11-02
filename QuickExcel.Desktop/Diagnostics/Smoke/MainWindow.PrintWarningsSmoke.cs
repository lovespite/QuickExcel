using System.IO;
using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunPrintWarningsSmokeAsync(string folder, Action<bool, string> check)
    {
        var original = _document.Snapshot; var sheet = original.Sheets[0];
        await using var supported = new WorkbookPaginator(original, sheet, "supported.xlsx");
        check(supported.Warnings.Count == 0,
            "Supported report prints without compatibility warnings");
        var brokenChart = sheet.Data.Charts[0].MapReferences(r => "'Missing'" + r[r.IndexOf('!')..]) with { Title = "缺少数据源" };
        var modified = sheet.WithPrinting(sheet.Printing with
        {
            Setup = sheet.Printing.Setup! with { CellComments = "atEnd" },
            HeaderFooter = sheet.Printing.HeaderFooter! with { OddHeader = "&G", OddFooter = "保留页脚 &P/&N", EvenFooter = "&Z" }
        }).WithImages(sheet.Images.Add(new SheetImage([0, 1, 2, 3], "png", 100, 100)))
            .WithData(sheet.Data with { Charts = sheet.Data.Charts.Add(brokenChart) });
        var diagnostics = new[]
        {
            new CompatibilityDiagnostic("print-test", CompatibilityKind.PreservedOnly, "chart/unsupported", "3D 图表暂未支持", SheetName: sheet.Name),
            new CompatibilityDiagnostic("print-test", CompatibilityKind.PreservedOnly, "drawing/transform", "图片旋转暂未支持", SheetName: sheet.Name),
            new CompatibilityDiagnostic("print-test", CompatibilityKind.PreservedOnly, "drawing/anchor", "其他工作表绘图", SheetName: "Other")
        };
        var state = new WorkbookFileState(Path.Combine(folder, "print-warnings.xlsx"), 2, [], original, [], [], compatibilityDiagnostics: diagnostics);
        var book = new WorkbookSnapshot(original.DocumentId, original.Revision, original.DateSystem, [modified],
            styles: original.Styles, fileState: state);
        await using var paginator = new WorkbookPaginator(book, modified, "print-warnings.xlsx");
        var warnings = string.Join("\n", paginator.Warnings);
        check(paginator.Warnings.Count == 7 && warnings.Contains("3D 图表") && warnings.Contains("图片旋转") &&
            warnings.Contains("图片 2") && warnings.Contains("图表 2") && warnings.Contains("批注") &&
            warnings.Contains("默认/奇数页页眉") && warnings.Contains("偶数页页脚") && !warnings.Contains("其他工作表"),
            "Print warnings identify unsupported drawings, broken images/charts, notes and each omitted header/footer");
        check(ReferenceEquals(book.Sheets[0], modified) && modified.Printing.Setup!.CellComments == "atEnd" &&
            modified.Printing.HeaderFooter!.OddHeader == "&G" && modified.Images.Length == 2 && modified.Data.Charts.Length == 2,
            "Omissions affect the print projection only, retaining all original content and metadata");
        var safe = sheet.WithPrinting(sheet.Printing with { HeaderFooter = sheet.Printing.HeaderFooter! with { OddHeader = null, OddFooter = "保留页脚 &P/&N" } });
        await using var expected = new WorkbookPaginator(original.WithSheets([safe]), safe, "print-warnings.xlsx");
        byte[] Pixels(WorkbookPaginator source)
        {
            var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap((int)Math.Ceiling(source.PageSize.Width),
                (int)Math.Ceiling(source.PageSize.Height), 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            bitmap.Render(source.GetPage(0).Visual);
            var bytes = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4]; bitmap.CopyPixels(bytes, bitmap.PixelWidth * 4, 0); return bytes;
        }
        check(paginator.PageCount == expected.PageCount && Pixels(paginator).SequenceEqual(Pixels(expected)),
            "Omitted objects leave supported cells, images, charts and footer identical to the printable reference");
        await PdfExport.WriteAsync(paginator, Path.Combine(folder, "print-warnings.pdf"));
        check(new FileInfo(Path.Combine(folder, "print-warnings.pdf")).Length > 1000, "PDF export continues despite unsupported content");
        File.WriteAllLines(Path.Combine(folder, "print-warnings.txt"), paginator.Warnings);

        var notices = 0; var previews = 0; Exception? error = null; var observing = true;
        EventManager.RegisterClassHandler(typeof(Window), LoadedEvent, new RoutedEventHandler((sender, _) =>
        {
            if (!observing || sender is not Window window || window.Owner != this ||
                window is not PrintWarningsDialog && !window.Title.StartsWith("打印预览 ·")) return;
            Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    if (window is PrintWarningsDialog)
                    {
                        notices++;
                        var text = string.Join("\n", ProductChildren<TextBlock>(window).Select(t => t.Text));
                        check(previews == 0 && text.Contains("可继续预览、打印或导出 PDF") && text.Contains("3D 图表") && text.Contains("本次跳过"),
                            "Actual print notice explains omissions before opening preview");
                        ProductChildren<Button>(window).Single(b => b.Content as string == "继续预览").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    }
                    else { previews++; check(notices == 1, "Acknowledging the notice continues directly into print preview"); }
                }
                catch (Exception ex) { error = ex; }
                finally { if (window.IsVisible) window.Close(); }
            });
        }));
        try
        {
            InstallDocument(new WorkbookDocument(book, _services), null, "print-warnings.xlsx");
            PrintPreview_Click(this, new()); await _workspaceInput;
            if (error is not null) throw error;
            check(notices == 1 && previews == 1 && !_document.IsDirty && !_document.CanUndo,
                "Unsupported rendering never blocks preview or creates document edits/history");
        }
        finally { observing = false; InstallDocument(new WorkbookDocument(original, _services), null, "E4 UI"); }
    }
}
