using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;

namespace QuickExcel.Desktop.Avalonia;

internal static class Program
{
    internal static string? SmokeDirectory;
    internal static string? InitialFile;
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length >= 2 && args[0] == "--workspace-smoke") SmokeDirectory = Path.GetFullPath(args[1]);
        else InitialFile = args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal));
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
            var data = Program.SmokeDirectory is null ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel") : Path.Combine(Program.SmokeDirectory, "settings");
            desktop.MainWindow = new MainWindow(data);
        }
        base.OnFrameworkInitializationCompleted();
    }
}
