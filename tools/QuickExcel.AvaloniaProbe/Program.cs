using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using QuickExcel.Rendering.Smoke;
using SkiaSharp;
using UglyToad.PdfPig;

namespace QuickExcel.AvaloniaProbe;

internal static class Program
{
    internal static string Output = null!;
    internal static bool Interactive;
    internal static bool VirtualHttps;
    internal static string? WebAssets;
    internal static string? FontDirectory;
    internal static Exception? Failure;
    internal static readonly List<string> Checks = [];

    [STAThread]
    public static int Main(string[] args)
    {
        Output = Path.GetFullPath(args.FirstOrDefault(a => !a.StartsWith("--")) ?? "artifacts/avalonia/probe");
        Directory.CreateDirectory(Output);
        Interactive = args.Contains("--interactive");
        VirtualHttps = args.Contains("--https");
        var webIndex = Array.IndexOf(args, "--web");
        if (webIndex >= 0) WebAssets = Path.GetFullPath(args[webIndex + 1]);
        var fontIndex = Array.IndexOf(args, "--fonts");
        if (fontIndex >= 0) FontDirectory = Path.GetFullPath(args[fontIndex + 1]);
        try
        {
            // WebKitGTK must share the GLib UI loop. A separate GTK thread caused
            // WebKit teardown to abort after managed Main had returned on Ubuntu.
            AppBuilder.Configure<ProbeApp>().UsePlatformDetect()
                .With(new X11PlatformOptions { UseGLibMainLoop = true })
                .LogToTrace().StartWithClassicDesktopLifetime(args);
            if (Failure is null && Checks.Count != (WebAssets is null ? 2 : 3)) Failure = new InvalidOperationException("Probe closed before completing its checks.");
            WriteResult(Failure);
            return Failure is null ? 0 : 1;
        }
        catch (Exception ex)
        {
            WriteResult(ex);
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    internal static void WriteResult(Exception? error)
    {
        var modules = Process.GetCurrentProcess().Modules.Cast<ProcessModule>()
            .Where(m => m.ModuleName.Contains("skia", StringComparison.OrdinalIgnoreCase) || m.ModuleName.Contains("harfbuzz", StringComparison.OrdinalIgnoreCase))
            .Select(m => new { m.ModuleName, m.FileName, m.FileVersionInfo.FileVersion,
                sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(m.FileName))) }).ToArray();
        File.WriteAllText(Path.Combine(Output, "result.json"), JsonSerializer.Serialize(new
        {
            passed = error is null, checks = Checks, error = error?.ToString(),
            runtime = RuntimeInformation.FrameworkDescription, rid = RuntimeInformation.RuntimeIdentifier,
            os = RuntimeInformation.OSDescription, modules,
            assemblies = AppDomain.CurrentDomain.GetAssemblies().Where(a => a.GetName().Name is { } n &&
                (n.StartsWith("Avalonia") || n.Contains("Skia") || n.Contains("HarfBuzz")))
                .Select(a => new { name = a.GetName().Name, version = a.GetName().Version?.ToString(), a.Location })
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}

public sealed class ProbeApp : Application
{
    private Bitmap? _gridImage;
    public override void Initialize() => Styles.Add(new FluentTheme());

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var closeRequested = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var allowClose = false;
            desktop.MainWindow = new Window
            {
                Title = "QuickExcel A0 — dependency probe", Width = 1100, Height = 760,
                Content = new StackPanel
                {
                    Margin = new Thickness(16), Spacing = 12,
                    Children =
                    {
                        new TextBlock { Text = "Avalonia / Skia / HarfBuzz 中文 ffi العربية", FontSize = 24 },
                        new TextBox { Text = "中文组合输入验证（需真实输入法）" },
                        new Button { Content = "Fluent control", ContextMenu = new ContextMenu
                        {
                            Items = { new MenuItem { Header = "A0 弹窗覆盖验证" }, new MenuItem { Header = "仅验证宿主，无工作簿写入" } }
                        } }
                    }
                }
            };
            var native = (StackPanel)desktop.MainWindow.Content!;
            var input = (TextBox)native.Children[1];
            var panes = new Grid { ColumnDefinitions = new ColumnDefinitions("400,5,*") };
            var divider = new GridSplitter();
            Grid.SetColumn(divider, 1);
            desktop.MainWindow.Content = null;
            panes.Children.Add(native); panes.Children.Add(divider);
            desktop.MainWindow.Content = panes;
            desktop.MainWindow.Closing += (_, e) =>
            {
                if (allowClose) return;
                e.Cancel = true;
                closeRequested.TrySetResult();
            };
            desktop.MainWindow.Opened += async (_, _) =>
            {
                try
                {
                    await Dispatcher.UIThread.InvokeAsync(() => CheckRendering(desktop.MainWindow), DispatcherPriority.Background);
                    if (Program.WebAssets is { } assets)
                    {
                        await using var web = new WebProbe(assets);
                        Grid.SetColumn(web.View, 2);
                        panes.Children.Add(web.View);
                        await web.VerifyAsync();
                        if (Program.Interactive)
                        {
                            await closeRequested.Task;
                        }
                        await web.CloseAsync(() => panes.Children.Remove(web.View));
                    }
                    else if (Program.Interactive) await closeRequested.Task;
                    File.WriteAllText(Path.Combine(Program.Output, "desktop.json"), JsonSerializer.Serialize(new
                    {
                        nativeText = input.Text, scaling = desktop.MainWindow.RenderScaling,
                        display = Environment.GetEnvironmentVariable("DISPLAY"),
                        session = Environment.GetEnvironmentVariable("XDG_SESSION_TYPE"),
                        interactive = Program.Interactive
                    }));
                    allowClose = true;
                    desktop.Shutdown();
                }
                catch (Exception ex)
                {
                    Program.Failure = ex;
                    Program.WriteResult(ex);
                    Console.Error.WriteLine(ex);
                    allowClose = true;
                    desktop.Shutdown(1);
                }
            };
            desktop.Exit += (_, _) => _gridImage?.Dispose();
        }
        base.OnFrameworkInitializationCompleted();
    }

    private void CheckRendering(Window window)
    {
        // Forces Avalonia's own text renderer and shared RichTextKit into the same process.
        using var capture = new RenderTargetBitmap(new PixelSize(960, 640), new Vector(96, 96));
        capture.Render(window);
        capture.Save(Path.Combine(Program.Output, "avalonia.png"), PngBitmapEncoderOptions.Default);
        Program.Checks.Add("Avalonia Fluent controls and text rendered");
        var fonts = Program.FontDirectory is null ? [] : Directory.GetFiles(Program.FontDirectory).Where(f => f.EndsWith(".ttf") || f.EndsWith(".otf")).ToArray();
        using var renderer = new SkiaRenderer(fonts, systemFonts: Program.FontDirectory is null);
        using var surface = SKSurface.Create(new SKImageInfo(900, 300));
        surface.Canvas.Clear(SKColors.White);
        var cell = new QuickExcel.Rendering.CellRenderInput(new(10, 10, 880, 100),
            new("QuickExcel 中文 ffi e\u0301 العربية", new StyleDefinition { FontName = "Noto Sans", EastAsianFontName = "Noto Sans CJK SC", FontSize = 20 }), true);
        renderer.DrawCell(surface.Canvas, cell);
        for (var row = 0; row < 5; row++)
        for (var column = 0; column < 5; column++)
            renderer.DrawCell(surface.Canvas, new(new(10 + column * 176, 110 + row * 36, 176, 36),
                new($"{(char)('A' + column)}{row + 1}: 中文 {row * column}", cell.Text.Style with { FontSize = 14 }), true));
        using (var png = surface.Snapshot().Encode(SKEncodedImageFormat.Png, 100))
        using (var file = File.Create(Path.Combine(Program.Output, "shared.png"))) png.SaveTo(file);
        using (var stream = File.Create(Path.Combine(Program.Output, "shared.pdf")))
        using (var pdf = SKDocument.CreatePdf(stream))
        {
            renderer.DrawCell(pdf.BeginPage(900, 120), cell);
            pdf.EndPage(); pdf.Close();
        }
        using var parsed = PdfDocument.Open(Path.Combine(Program.Output, "shared.pdf"));
        if (!parsed.GetPage(1).Text.Contains("QuickExcel")) throw new InvalidDataException("Shared PDF has no expected text.");
        var fixture = ReportFixture.Create();
        using var report = new SkiaReport(fixture, fixture.Sheets[0], "a0-fixture.xlsx", fontFiles: fonts, systemFonts: Program.FontDirectory is null);
        report.WritePdf(Path.Combine(Program.Output, "report.pdf"));
        File.WriteAllText(Path.Combine(Program.Output, "report-layout.json"), JsonSerializer.Serialize(report.Job.Layout));
        File.WriteAllText(Path.Combine(Program.Output, "report-diagnostics.json"), JsonSerializer.Serialize(new { report.Job.Warnings, report.Renderer.Diagnostics }));
        using (var reportPdf = PdfDocument.Open(Path.Combine(Program.Output, "report.pdf")))
            if (reportPdf.NumberOfPages != 3) throw new InvalidDataException("Report page count changed.");
        for (var page = 0; page < report.Job.Layout.Pages.Length; page++)
        {
            using var pageBitmap = report.RenderPage(page);
            using var encoded = pageBitmap.Encode(SKEncodedImageFormat.Png, 100);
            using var target = File.Create(Path.Combine(Program.Output, $"report-{page + 1}.png"));
            encoded.SaveTo(target);
        }
        foreach (var extension in new[] { "png", "jpg", "gif", "bmp", "tiff" })
        {
            using var decoded = renderer.DecodeImage(new SheetImage(ImageFixtures.Bytes(extension), extension, 0, 0));
            if (decoded.Width != 2 || decoded.Height != 1 || decoded.GetPixel(0, 0).Red < 180) throw new InvalidDataException($"Image decode failed: {extension}");
        }
        File.WriteAllText(Path.Combine(Program.Output, "render-diagnostics.json"), JsonSerializer.Serialize(new { renderer.Diagnostics, reportDiagnostics = report.Renderer.Diagnostics }));
        _gridImage = new Bitmap(Path.Combine(Program.Output, "shared.png"));
        ((StackPanel)((Grid)window.Content!).Children[0]).Children.Add(new Image { Source = _gridImage, Stretch = Stretch.Uniform, Height = 240 });
        Program.Checks.Add("Shared Skia/RichTextKit cell and PDF rendered and parsed");
    }
}
