using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;

namespace QuickExcel.Desktop.Avalonia;

internal static class Program
{
    internal static string? SmokeDirectory;
    internal static string SmokeMode = "workspace";
    internal static string? InitialFile;
    internal static bool WebTestMode => SmokeMode is "web" or "web-benchmark";
    internal static string WebBenchmarkKind = "all";
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length >= 2 && args[0] is "--workspace-smoke" or "--grid-smoke" or "--grid-benchmark" or "--web-smoke" or "--web-benchmark" or "--manual-check")
        { SmokeDirectory = Path.GetFullPath(args[1]); SmokeMode = args[0] switch { "--grid-smoke" => "grid", "--grid-benchmark" => "benchmark", "--web-smoke" => "web", "--web-benchmark" => "web-benchmark", "--manual-check" => "manual", _ => "workspace" }; if (SmokeMode == "manual") InitialFile = args.ElementAtOrDefault(2); }
        else InitialFile = args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal));
        if (SmokeMode == "web-benchmark") WebBenchmarkKind = args.ElementAtOrDefault(2) ?? "all";
        return AppBuilder.Configure<DesktopApp>().UsePlatformDetect().With(new X11PlatformOptions { UseGLibMainLoop = true })
            .LogToTrace().StartWithClassicDesktopLifetime(args);
    }
}

public sealed class DesktopApp : Application
{
    public override void Initialize() { RequestedThemeVariant = global::Avalonia.Styling.ThemeVariant.Light; Styles.Add(new FluentTheme()); }
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var data = Program.SmokeDirectory is null ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel", "Avalonia") : Path.Combine(Program.SmokeDirectory, "settings");
            desktop.MainWindow = new MainWindow(data);
        }
        base.OnFrameworkInitializationCompleted();
    }
}
