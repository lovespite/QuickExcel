using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Presentation;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Window? _printWindow;
    private ContentControl? _printContent;
    private TextBlock? _printStatus;
    private Button? _printRebuild;
    private PrintSession? _printSession;
    private PdfPreviewPage? _pdfPage;
    private CancellationTokenSource? _printCancellation;
    private Task? _printGeneration, _printClose;
    private bool _printClosing;
    private Guid _printDocument;
    private void InstallPrintCommands()
    {
        RibbonGroup("页面布局", "打印输出", RibbonButton("打印预览 / PDF", OpenPrintPreviewAsync, false));
    }
    private async Task OpenPrintPreviewAsync()
    {
        if (_printWindow is not null) { _printWindow.Activate(); return; }
        var window = _printWindow = new Window { Title = "打印预览 · QuickExcel", Width = 1120, Height = 820, MinWidth = 760, MinHeight = 520, WindowStartupLocation = WindowStartupLocation.CenterOwner };
        _printClose = null;
        var root = new DockPanel(); var tools = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, Margin = new(12, 8) };
        _printRebuild = new Button { Content = "从当前工作表重新生成" }; _printRebuild.Click += async (_, _) => await PerformAsync(RebuildPrintPreviewAsync); tools.Children.Add(_printRebuild);
        var settings = new Button { Content = "页面设置" }; settings.Click += async (_, _) => await PerformAsync(async () => { await PageSettingsAsync(); await RebuildPrintPreviewAsync(); }); tools.Children.Add(settings);
        var export = new Button { Content = "导出 PDF" }; export.Click += async (_, _) => await PerformAsync(ExportPreviewPdfAsync); tools.Children.Add(export);
        var print = new Button { Content = "系统打印" }; print.Click += async (_, _) => await PerformAsync(PrintPreviewAsync); tools.Children.Add(print);
        var cancel = new Button { Content = "取消生成 / 导出" }; cancel.Click += (_, _) => { _printCancellation?.Cancel(); _calculationCancellation?.Cancel(); }; tools.Children.Add(cancel);
        DockPanel.SetDock(tools, Dock.Top); root.Children.Add(tools);
        _printStatus = new TextBlock { Margin = new(12, 6), TextWrapping = TextWrapping.Wrap }; DockPanel.SetDock(_printStatus, Dock.Bottom); root.Children.Add(_printStatus);
        _printContent = new ContentControl(); root.Children.Add(_printContent); window.Content = root;
        window.Closing += async (_, e) =>
        {
            if (_printClosing) return; e.Cancel = true; _printCancellation?.Cancel();
            try { await ClosePrintPreviewAsync(); } catch (Exception ex) { Error(ex); }
        };
        window.Show(this);
        await RebuildPrintPreviewAsync();
    }
    private Task RebuildPrintPreviewAsync() => _printGeneration = RebuildPrintCoreAsync();
    private async Task RebuildPrintCoreAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (_printWindow is null || _printClosing || !await CommitEditorsAsync()) return;
        await CalculateAsync();
        if (_printClosing || _calculationPending || _calculationFailure is not null) throw new InvalidOperationException("重算未完成或已取消，本次未生成新的打印任务。" + _calculationFailure);
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        if (sheet.Cells.Values.Any(c => c.Content.Kind == CellContentKind.Formula &&
            (c.Calculation is null || c.Calculation.InputRevision != snapshot.Revision || c.Calculation.State is not (CalculationState.Fresh or CalculationState.Error))))
            throw new InvalidOperationException("当前工作表含受阻或尚未完成重算的公式，请完成计算后再打印。");
        _printCancellation?.Cancel(); using var cancellation = new CancellationTokenSource(); _printCancellation = cancellation;
        var window = _printWindow; var identity = Document.Session.SessionId; var generation = ++_webGeneration;
        _printStatus!.Text = "正在生成并验证完整 PDF完成后显示预览。";
        PrintSession? generated = null;
        try
        {
            generated = await PrintSession.CreateAsync(snapshot, sheet.Id, _workspace.FilePath ?? "未命名.xlsq", Path.Combine(_userDataDirectory, "printing"), ConditionalPresentation.GetStyle, cancellationToken: cancellation.Token, visual: ConditionalPresentation.GetVisual);
            cancellation.Token.ThrowIfCancellationRequested();
            if (_printWindow != window || identity != Document.Session.SessionId || Document.Snapshot.Revision != snapshot.Revision) throw new OperationCanceledException("工作簿在生成期间改变，请重新生成预览。");
            if (_pdfPage is not null) { await _pdfPage.DisposeAsync(); _pdfPage = null; }
            if (_printSession is not null) await _printSession.DisposeAsync();
            _printSession = generated; generated = null; _printDocument = identity;
            var page = _pdfPage = new PdfPreviewPage(WebAssets, identity.ToString("N"), generation, _userDataDirectory);
            page.ExportRequested += () => _ = PerformAsync(ExportPreviewPdfAsync);
            page.PrintRequested += () => _ = PerformAsync(PrintPreviewAsync);
            page.FocusRequested += () => _printRebuild?.Focus();
            page.Failed += ex => { if (_printStatus is not null) _printStatus.Text = "预览页已失败，可重新生成；已生成的 PDF 仍可导出。" + ex.Message; Error(ex); };
            _printContent!.Content = page.Browser.View;
            await page.LoadAsync(_printSession); cancellation.Token.ThrowIfCancellationRequested();
            RefreshPrintStatus();
        }
        catch (Exception ex) { if (_printStatus is not null) _printStatus.Text = ex is OperationCanceledException ? "打印预览生成已取消。" : ex.Message; throw; }
        finally { if (generated is not null) await generated.DisposeAsync(); if (ReferenceEquals(_printCancellation, cancellation)) _printCancellation = null; }
    }
    private void RefreshPrintStatus()
    {
        if (_printStatus is null || _printSession is null) return;
        var job = _printSession.Job; var stale = _printDocument != Document.Session.SessionId || job.Book.Revision != Document.Snapshot.Revision;
        _printStatus.Text = $"{job.Sheet.Name} · {job.Layout.Pages.Length} 页 · 文档修订 {job.Book.Revision}" + (stale ? " · 工作簿已变化，当前显示此前的打印任务；需要最新内容时请重新生成。" : "") +
            (_printSession.Warnings.Count > 0 ? "\n输出提示：" + string.Join("；", _printSession.Warnings.Take(10)) : "");
    }
    private async Task ExportPreviewPdfAsync()
    {
        var session = _printSession ?? throw new InvalidOperationException("请先生成打印预览。");
        var file = await _files.SaveAsync(new("导出当前打印任务", [new("PDF 文档", ["*.pdf"])], Path.GetFileNameWithoutExtension(session.Job.FileName) + ".pdf", ".pdf"));
        if (file is null) return;
        using var cancellation = new CancellationTokenSource(); _printCancellation = cancellation;
        try { await session.ExportAsync(file.Path, cancellation.Token); if (_printStatus is not null) _printStatus.Text = "已导出当前打印任务：" + file.Path; }
        finally { if (ReferenceEquals(_printCancellation, cancellation)) _printCancellation = null; }
    }
    private async Task PrintPreviewAsync()
    {
        var session = _printSession ?? throw new InvalidOperationException("请先生成打印预览。");
        if(!OperatingSystem.IsWindows()) {await PrintCupsAsync(session);return;}
        using var job = WindowsPrintJob.Choose((_printWindow ?? this).TryGetPlatformHandle()?.Handle ?? 0, session);
        if (job is null) return;
        using var cancellation = new CancellationTokenSource(); _printCancellation = cancellation;
        try { await job.PrintAsync(session, cancellation.Token); if (_printStatus is not null) _printStatus.Text = "当前预览任务已提交给系统打印服务。"; }
        finally { if (ReferenceEquals(_printCancellation, cancellation)) _printCancellation = null; }
    }
    private async Task PrintCupsAsync(PrintSession session)
    {
        var printers=await QuickExcel.Printing.CupsPrintService.GetDestinationsAsync();
        if(printers.Count==0)throw new InvalidOperationException("没有可用的系统打印机，请先在系统设置中添加打印机。");
        var dialog=new FieldsDialog("系统打印").Note($"当前预览共 {session.Job.Layout.Pages.Length} 页；按原始尺寸提交 PDF。")
            .Choice("printer","打印机",printers.Select(p=>p.DisplayName).ToArray(),Math.Max(0,printers.ToList().FindIndex(p=>p.IsDefault)))
            .Text("pages","页范围（留空为全部，例如 1-3,5）").Text("copies","份数","1");
        dialog.Validate=()=>{_ = QuickExcel.Printing.CupsPrintService.NormalizePageRanges(dialog.String("pages"),session.Job.Layout.Pages.Length);dialog.Integer("copies",1,999);};
        if(!await dialog.ShowAsync(this))return;
        var printer=printers[dialog.Index("printer")];
        var ranges=QuickExcel.Printing.CupsPrintService.NormalizePageRanges(dialog.String("pages"),session.Job.Layout.Pages.Length);
        await using var lease=session.RetainPdf();
        using var cancellation=new CancellationTokenSource();_printCancellation=cancellation;
        try
        {
            var media=await QuickExcel.Printing.CupsPrintService.ValidateMediaAsync(printer,session.Job.Layout,cancellation.Token);
            var receipt=await QuickExcel.Printing.CupsPrintService.SubmitAsync(printer,media,session.PdfPath,ranges,dialog.Integer("copies",1,999),cancellation.Token);
            if(_printStatus is not null)_printStatus.Text="已提交系统打印队列："+receipt;
        }
        finally{if(ReferenceEquals(_printCancellation,cancellation))_printCancellation=null;}
    }
    private Task ClosePrintPreviewAsync() => _printClose ??= ClosePrintCoreAsync();
    private async Task ClosePrintCoreAsync()
    {
        if (_printWindow is null) return;
        _printClosing = true; _printCancellation?.Cancel(); _calculationCancellation?.Cancel();
        try
        {
            if (_printGeneration is { } pending) try { await pending; } catch { /* The generation caller already reports its failure. */ }
            if (_pdfPage is not null) await _pdfPage.DisposeAsync();
            if (_printSession is not null) await _printSession.DisposeAsync();
            _printWindow.Close();
        }
        finally { _pdfPage = null; _printSession = null; _printWindow = null; _printContent = null; _printStatus = null; _printRebuild = null; _printClosing = false; }
    }
}
