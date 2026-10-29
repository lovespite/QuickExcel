using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace QuickExcel.Desktop;

public partial class App : Application
{
    private string? _smokeEvidenceDirectory;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        if (e.Args.Length > 0 && e.Args[0] is "--smoke-test" or "--interaction-smoke" or "--workspace-smoke" or "--product-smoke" or "--grid-benchmark" or "--resize-benchmark")
            _smokeEvidenceDirectory = e.Args.Length > 1 ? e.Args[1] : Path.Combine(AppContext.BaseDirectory, "smoke");
        DispatcherUnhandledException += OnUnhandled;
        var window = new MainWindow();
        MainWindow = window;
        window.Loaded += async (_, _) => await window.StartAsync(e.Args);
        window.Show();
    }

    private static void OnUnhandled(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        if (Current is App { _smokeEvidenceDirectory: { } folder } app)
        {
            Directory.CreateDirectory(folder);
            File.AppendAllText(Path.Combine(folder, "smoke-error.txt"), e.Exception + Environment.NewLine);
            e.Handled = true;
            app.Shutdown(1);
            return;
        }
        MessageBox.Show(e.Exception.Message, "QuickExcel", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }
}
