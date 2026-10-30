using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace QuickExcel.Desktop;

public partial class App : Application
{
    private string? _smokeEvidenceDirectory;

    protected override void OnStartup(StartupEventArgs e)
    {
        // Keep WPF's automatic rendering as the default. SoftwareOnly avoids the
        // reported desktop stall but is too slow for normal use on the affected
        // machine, so retain it only as an explicit process-local diagnostic.
        var explicitSoftware = e.Args.Contains("--software-rendering", StringComparer.OrdinalIgnoreCase);
        var explicitDefault = e.Args.Contains("--default-rendering", StringComparer.OrdinalIgnoreCase);
        var softwareRendering = explicitSoftware;
        var args = e.Args.Where(arg => !arg.Equals("--software-rendering", StringComparison.OrdinalIgnoreCase) &&
            !arg.Equals("--default-rendering", StringComparison.OrdinalIgnoreCase)).ToArray();
        RenderOptions.ProcessRenderMode = softwareRendering ? RenderMode.SoftwareOnly : RenderMode.Default;
        base.OnStartup(e);
        if (args.Length > 0 && args[0] is "--smoke-test" or "--interaction-smoke" or "--workspace-smoke" or "--product-smoke" or "--ribbon-smoke" or "--grid-benchmark" or "--resize-benchmark" or "--data-operations-smoke" or "--signature-smoke" or "--horizontal-scroll-benchmark" or "--fast-grid-benchmark" or "--fast-grid-drag-benchmark" or "--grid-integration-smoke" or "--xlsx-rendering-smoke" or "--fill-smoke" or "--formula-entry-smoke")
            _smokeEvidenceDirectory = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "smoke");
        if (_smokeEvidenceDirectory is { } evidence)
        {
            Directory.CreateDirectory(evidence);
            File.WriteAllText(Path.Combine(evidence, "rendering.json"), JsonSerializer.Serialize(new
            {
                requestedSoftwareRendering = softwareRendering,
                renderModeSelection = explicitSoftware ? "explicit-software" : explicitDefault ? "explicit-default" : "automatic-default",
                processRenderMode = RenderOptions.ProcessRenderMode.ToString(),
                renderCapabilityTier = RenderCapability.Tier >> 16,
                runtime = Environment.Version.ToString(), os = Environment.OSVersion.ToString()
            }, new JsonSerializerOptions { WriteIndented = true }));
        }
        DispatcherUnhandledException += OnUnhandled;
        var window = new MainWindow();
        MainWindow = window;
        window.Loaded += async (_, _) => await window.StartAsync(args);
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
