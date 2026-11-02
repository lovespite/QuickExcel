using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunReportingSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); _smokeMode = true; var checks = new List<string>();
        void Check(bool value, string message) { if (!value) throw new InvalidOperationException(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); }
        Action<ReportFieldsDialog>? onDialog = null;
        EventManager.RegisterClassHandler(typeof(ReportFieldsDialog), LoadedEvent, new RoutedEventHandler((sender, _) =>
        {
            if (onDialog is { } callback && sender is ReportFieldsDialog dialog) { onDialog = null; Dispatcher.BeginInvoke(() => callback(dialog)); }
        }));
        async Task Dialog(Action open, Action<ReportFieldsDialog> edit)
        {
            var complete = new TaskCompletionSource();
            onDialog = dialog =>
            {
                try { edit(dialog); dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Check(!dialog.IsVisible, "Dialog applies successfully: " + dialog.Title); complete.SetResult(); }
                catch (Exception ex) { dialog.Close(); complete.TrySetException(ex); }
            };
            open(); await complete.Task.WaitAsync(TimeSpan.FromSeconds(15)); await Idle();
        }
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "E4 UI");
            var sheet = _document.Snapshot.Sheets[0];
            string[,] values = { { "Month", "Revenue", "Cost" }, { "Jan", "25", "16" }, { "Feb", "37.5", "24" }, { "Mar", "50", "32" } };
            for (var r = 0; r < 4; r++) for (var c = 0; c < 3; c++) _editor.SetCell(0, r, c, values[r, c]);
            RefreshSheet(true); GridSurface.SelectRange(new(0, 0, 3, 2), new(0, 0)); await Idle();
            await Dialog(() => ChartManage_Click(this, new()), d => ((TextBox)d.Fields["图表标题"]).Text = "E4 revenue");
            Check(_document.Snapshot.Sheets[0].Data.Charts.Length == 1, "Real chart insertion handler commits a chart");
            await Dialog(() => NoteEdit_Click(this, new()), d => ((TextBox)d.Fields["批注内容"]).Text = "E4 traditional note");
            Check(_document.Snapshot.Sheets[0].Annotations.Notes.Count == 1, "Real note editor creates a traditional note");
            GridSurface.Select(new(1, 1)); await Idle();
            await Dialog(() => CellProtection_Click(this, new()), d => ((CheckBox)d.Fields["锁定：工作表保护后禁止修改内容"]).IsChecked = false);
            await Dialog(() => Protection_Click(this, new()), d => ((CheckBox)d.Fields["启用工作表保护（不设置密码）"]).IsChecked = true);
            Check(_document.Snapshot.Sheets[0].Data.Protection?.Sheet == true, "Protection handler commits flags");
            var before = _document.Snapshot;
            try { _editor.SetCell(0, 0, 0, "blocked"); throw new InvalidOperationException("Locked write accepted"); } catch (InvalidOperationException) { Check(ReferenceEquals(before, _document.Snapshot), "Locked write is atomic"); }
            _editor.SetCell(0, 1, 1, "123.5"); Check(_document.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue == 123.5, "Unlocked write succeeds");
            await Dialog(() => Protection_Click(this, new()), d => ((CheckBox)d.Fields["启用工作表保护（不设置密码）"]).IsChecked = false);
            _editor.SetCell(0, 0, 4, "Merged report title"); _editor.Execute(new MergeCellsCommand(sheet.Id, new(0, 4, 1, 6)));
            _editor.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 4)], new(Bold: true, FillColor: "FFDDEBF7")));
            var pixels = Enumerable.Range(0, 16 * 16).SelectMany(i => i % 16 < 8 ? new byte[] { 180, 70, 25, 255 } : new byte[] { 30, 160, 240, 255 }).ToArray();
            var fixtureBitmap = BitmapSource.Create(16, 16, 96, 96, PixelFormats.Bgra32, null, pixels, 64);
            var fixtureEncoder = new PngBitmapEncoder(); fixtureEncoder.Frames.Add(BitmapFrame.Create(fixtureBitmap));
            using var fixtureStream = new MemoryStream(); fixtureEncoder.Save(fixtureStream);
            _editor.Execute(new SetImageCommand(sheet.Id, 0, new SheetImage(fixtureStream.ToArray(), "png", 0, 0).WithAnchor(
                DrawingGeometry.FromBounds(_document.Snapshot.Sheets[0], new(520.5, 70.25, 150, 65), DrawingAnchorKind.OneCell))));
            await Dialog(() => ImageManage_Click(this, new()), d =>
            {
                ((ComboBox)d.Fields["图片"]).SelectedIndex = 1;
                onDialog = edit => { ((TextBox)edit.Fields["宽度"]).Text = "160.5"; edit.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); };
            });
            Check(Math.Abs(_document.Snapshot.Sheets[0].Images[0].Anchor.Bounds(_document.Snapshot.Sheets[0]).Width - 160.5) < .001, "Real image sizing dialog preserves fractional offsets");
            await Dialog(() => PageSettings_Click(this, new()), d =>
            {
                ((TextBox)d.Fields["打印区域（如 A1:F30；留空为已用区域）"]).Text = "A1:G25";
                ((CheckBox)d.Fields["适应页数（0 表示该方向不限）"]).IsChecked = true;
                ((TextBox)d.Fields["适应高度（页）"]).Text = "1";
                ((TextBox)d.Fields["页眉（&L 左 / &C 中 / &R 右；&P 页码 / &N 总页数）"]).Text = "E4 report";
            });
            var path = Path.Combine(folder, "ui-report.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            await OpenPathAsync(path); await RecalculateAsync(); await Idle();
            Check(_document.Snapshot.Sheets[0].Data.Charts.Length == 1 && _document.Snapshot.Sheets[0].Annotations.Notes.Count == 1, "Ordinary save/reopen retains chart and note");
            CaptureGridEvidence(Path.Combine(folder, "report-grid.png"));
            async Task Render(string prefix, int sheetIndex = 0)
            {
                var book = _document.Snapshot; var paginator = new WorkbookPaginator(book, book.Sheets[sheetIndex], prefix + ".xlsx");
                if (prefix == "paged") Check(paginator.PageCount == 2 && paginator.Layout.Pages[1].Rows.Take(2).SequenceEqual([0, 20]), "Manual page starts at row 21 after repeated title, matching Excel");
                Check(paginator.PageCount > 0, "Pagination produces pages: " + prefix);
                var previewPage = paginator.CreatePreviewDocument().Pages[0].Child;
                previewPage.Measure(paginator.PageSize); previewPage.Arrange(new Rect(paginator.PageSize)); previewPage.UpdateLayout();
                foreach (var dpi in new[] { 96, 144, 192 })
                {
                    var bitmap = new RenderTargetBitmap((int)Math.Ceiling(paginator.PageSize.Width * dpi / 96), (int)Math.Ceiling(paginator.PageSize.Height * dpi / 96), dpi, dpi, PixelFormats.Pbgra32);
                    bitmap.Render(paginator.GetPage(0).Visual); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
                    using var output = File.Create(Path.Combine(folder, $"{prefix}-{dpi}.png")); encoder.Save(output);
                    var previewBitmap = new RenderTargetBitmap(bitmap.PixelWidth, bitmap.PixelHeight, dpi, dpi, PixelFormats.Pbgra32);
                    previewBitmap.Render(previewPage);
                    var previewEncoder = new PngBitmapEncoder(); previewEncoder.Frames.Add(BitmapFrame.Create(previewBitmap));
                    using (var previewOutput = File.Create(Path.Combine(folder, $"{prefix}-preview-{dpi}.png"))) previewEncoder.Save(previewOutput);
                    // FixedPage also paints a white background; omit only the fractional paper-edge pixels.
                    var interior = new Int32Rect(0, 0, (int)Math.Floor(paginator.PageSize.Width * dpi / 96), (int)Math.Floor(paginator.PageSize.Height * dpi / 96));
                    var stride = interior.Width * 4; var printed = new byte[stride * interior.Height]; var previewed = new byte[printed.Length];
                    bitmap.CopyPixels(interior, printed, stride, 0); previewBitmap.CopyPixels(interior, previewed, stride, 0);
                    Check(printed.SequenceEqual(previewed), $"Fixed-document preview matches printed page at {dpi} DPI: {prefix}");
                }
                for (var page = 1; page < paginator.PageCount; page++)
                {
                    var bitmap = new RenderTargetBitmap((int)Math.Ceiling(paginator.PageSize.Width), (int)Math.Ceiling(paginator.PageSize.Height), 96, 96, PixelFormats.Pbgra32);
                    bitmap.Render(paginator.GetPage(page).Visual); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
                    using var output = File.Create(Path.Combine(folder, $"{prefix}-page{page + 1}.png")); encoder.Save(output);
                }
                await PdfExport.WriteAsync(paginator, Path.Combine(folder, prefix + ".pdf"));
                Check(new FileInfo(Path.Combine(folder, prefix + ".pdf")).Length > 1000, "Windows PDF export succeeds: " + prefix);
                File.WriteAllText(Path.Combine(folder, prefix + "-layout.json"), JsonSerializer.Serialize(paginator.Layout, new JsonSerializerOptions { WriteIndented = true }));
            }
            await Render("quickexcel");
            await RunPrintWarningsSmokeAsync(folder, Check);
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); await RecalculateAsync(); await Idle();
                Check(_document.Snapshot.Sheets[0].Data.Charts.Length == 3, "Three native Excel chart types import and display");
                CaptureGridEvidence(Path.Combine(folder, "native-grid.png")); await Render("native");
                var paged = Array.FindIndex(_document.Snapshot.Sheets.ToArray(), s => s.Name == "Paged");
                if (paged >= 0) await Render("paged", paged);
            }
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "Sparse boundary");
            _editor.SetCell(0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1, "edge"); RefreshSheet(true);
            NavigateToCell(0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1); await Idle();
            Check(_grid.Current.Row == WorkbookLimits.MaxRows - 1 && _grid.Current.Column == WorkbookLimits.MaxColumns - 1, "New coordinate boundary is reachable in WPF");
            CaptureGridEvidence(Path.Combine(folder, "capacity-boundary.png"));
            File.WriteAllLines(Path.Combine(folder, "reporting-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { onDialog = null; }
    }
}
