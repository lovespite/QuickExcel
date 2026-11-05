using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;

namespace QuickExcel.Desktop.Avalonia;

internal static class Program
{
    internal static bool SoftwareRendering => Environment.GetEnvironmentVariable("QUICKEXCEL_SOFTWARE_RENDERING") == "1";
    internal static string? SmokeDirectory = null;
    internal static string SmokeMode = "workspace";
    internal static string? InitialFile;
    internal static bool WebTestMode =>
#if QUICKEXCEL_DIAGNOSTICS
        SmokeMode is "web" or "web-benchmark";
#else
        false;
#endif
    internal static string? SelfCheckDirectory, UserDataDirectory;
    internal static string? SettingsMigrationError;
    internal static string WebBenchmarkKind = "all";
    internal static bool WebScriptsOnly;
    internal static bool WebModulesOnly;
    [STAThread]
    public static int Main(string[] args)
    {
#if !QUICKEXCEL_DIAGNOSTICS
        if (args.FirstOrDefault()?.StartsWith("--", StringComparison.Ordinal) == true &&
            (args.Length < 2 || args[0] is not ("--self-check" or "--user-data"))) return 2;
#endif
        if (args.Length >= 2 && args[0] == "--self-check" && Directory.Exists(args[1])) return 2;
#if QUICKEXCEL_DIAGNOSTICS
        if (args.Length >= 2 && args[0] is "--localization-smoke" or "--sheets-smoke" or "--data-tools-smoke" or "--borders-smoke" or "--number-format-smoke" or "--menus-smoke" or "--conditional-format-smoke" or "--colors-smoke" or "--clipboard-exchange" or "--drawing-smoke" or "--workspace-smoke" or "--execution-smoke" or "--automation-smoke" or "--ribbon-filter-smoke" or "--page-break-smoke" or "--compatibility-smoke" or "--grid-smoke" or "--grid-benchmark" or "--web-smoke" or "--web-benchmark" or "--manual-check")
        { SmokeDirectory = Path.GetFullPath(args[1]); SmokeMode = args[0] switch { "--localization-smoke" => "localization", "--sheets-smoke" => "sheets", "--data-tools-smoke" => "data-tools", "--borders-smoke" => "borders", "--number-format-smoke" => "number-format", "--menus-smoke" => "menus", "--conditional-format-smoke" => "conditional-format", "--colors-smoke" => "colors", "--clipboard-exchange" => "clipboard-exchange", "--drawing-smoke" => "drawing", "--execution-smoke" => "execution", "--automation-smoke" => "automation", "--ribbon-filter-smoke" => "ribbon-filter", "--page-break-smoke" => "page-break", "--compatibility-smoke" => "compatibility", "--grid-smoke" => "grid", "--grid-benchmark" => "benchmark", "--web-smoke" => "web", "--web-benchmark" => "web-benchmark", "--manual-check" => "manual", _ => "workspace" }; if (SmokeMode == "manual") InitialFile = args.ElementAtOrDefault(2); }
        else InitialFile = args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal));
        if (SmokeMode == "web-benchmark") WebBenchmarkKind = args.ElementAtOrDefault(2) ?? "all";
        if (SmokeMode == "web")
        {
            WebScriptsOnly = args.ElementAtOrDefault(2) == "--scripts-only";
            WebModulesOnly = args.ElementAtOrDefault(2) == "--modules-only";
        }
#else
        InitialFile = args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal));
#endif
        if (args.Length >= 2 && args[0] == "--self-check")
        { SelfCheckDirectory = Path.GetFullPath(args[1]); UserDataDirectory = Path.Combine(SelfCheckDirectory, "settings"); InitialFile = null; }
        if (args.Length >= 2 && args[0] == "--user-data")
        { UserDataDirectory = Path.GetFullPath(args[1]); InitialFile = args.ElementAtOrDefault(2); }
        var builder = AppBuilder.Configure<DesktopApp>().UsePlatformDetect().With(new X11PlatformOptions { UseGLibMainLoop = true });
        if (OperatingSystem.IsWindows() && SoftwareRendering) builder = builder.With(new Win32PlatformOptions { RenderingMode = [Win32RenderingMode.Software] });
        return builder.LogToTrace().StartWithClassicDesktopLifetime(args);
    }
}

public sealed class DesktopApp : Application
{
    internal const string ProductIconKey = "QuickExcel.ProductIcon";

    public override void Initialize()
    {
        RequestedThemeVariant = ThemeVariant.Light;
        Styles.Add(new FluentTheme());
        using var stream = AssetLoader.Open(new Uri("avares://QuickExcel.Avalonia/Assets/QuickExcel.png"));
        var icon = new Bitmap(stream);
        Resources[ProductIconKey] = icon;
        Styles.Add(new Style(selector => selector.Is<Window>())
        {
            Setters = { new Setter(Window.IconProperty, new WindowIcon(icon)) }
        });
    }
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var legacy = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel");
            var data = Program.UserDataDirectory ?? (Program.SmokeDirectory is null ? Path.Combine(legacy, "Avalonia") : Path.Combine(Program.SmokeDirectory, "settings"));
            if (Program.UserDataDirectory is null && Program.SmokeDirectory is null)
                Program.SettingsMigrationError = QuickExcel.Presentation.DesktopSettingsMigration.ImportOnce(legacy, data);
            var settings = new QuickExcel.Presentation.DesktopSettings(Path.Combine(data, "settings.json"));
            settings.Load();
            UiText.Initialize(settings.Value.Language);
            desktop.MainWindow = new MainWindow(data, settings);
        }
        base.OnFrameworkInitializationCompleted();
    }
}
