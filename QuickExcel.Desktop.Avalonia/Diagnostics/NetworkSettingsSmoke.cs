using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckNetworkSettingsAsync(Action<bool, string> check, string directory)
    {
        static async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var original = _settings.Value;
        var network = Document.Session.Network;
        try
        {
            _dialogDriver = async window =>
            {
                await Layout();
                NumericUpDown Input(string name) => window.GetVisualDescendants().OfType<NumericUpDown>().Single(input => input.Name == name);
                var body = Input("NetworkBodyLimitMiB"); var timeout = Input("NetworkActionTimeoutSeconds");
                var entries = Input("NetworkPanelEntryLimit"); var memory = Input("NetworkPanelMemoryLimitMiB");
                var allow = window.GetVisualDescendants().OfType<CheckBox>().Single(box => box.Name == "AllowScriptNetwork");
                var apply = window.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "ApplyNetworkLimits");
                check(body.Value == 8 && timeout.Value == 10 && entries.Value == 200 && memory.Value == 50 && allow.IsChecked == true,
                    "Network settings show the requested defaults: 8 MiB, 10 seconds, 200 entries and 50 MiB");
                allow.IsChecked = false;
                check(!network.Options.AllowNetwork && !_settings.Value.ScriptNetwork.AllowNetwork, "Allow network switch immediately updates the current document and saved preferences");
                var denied = false;
                try { await Task.Run(() => network.EnsureAuthorized(["api.example.com"], default)).WaitAsync(TimeSpan.FromSeconds(2)); }
                catch (InvalidOperationException) { denied = true; }
                check(denied && window.OwnedWindows.Count == 0 && _scriptWindow!.OwnedWindows.Count == 0,
                    "Disabled network access rejects permission requests without opening a consent dialog");
                body.Value = 4; timeout.Value = 20; entries.Value = 2; memory.Value = 1;
                apply.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                check(network.Options == new ScriptNetworkOptions(4, 20, 1, 2, false), "Applying numeric limits preserves the disabled network switch");
                allow.IsChecked = true;
                var persisted = new DesktopSettings(Path.Combine(directory, "settings", "settings.json")); persisted.Load();
                check(persisted.Value.ScriptNetwork == new ScriptNetworkOptions(4, 20, 1, 2), "All network settings survive reload through generated JSON metadata");
                var first = new ScriptNetworkEntry { Url = "https://api.example.com/report", ResponseBody = new byte[600_000] };
                network.Record(first);
                network.Record(new() { Url = first.Url, ResponseBody = new byte[600_000] });
                network.Record(new() { Url = first.Url, ResponseBody = new byte[600_000] });
                await Layout();
                check(network.Entries.Length == 2 && network.RetainedMemoryBytes <= 1024 * 1024 && first.ResponseBody is null,
                    "New settings immediately enforce both panel count and memory limits");
                check(_networkSummary!.Text!.Contains("/ 1 MiB") && _networkSummary.Text.Contains("最多 2 条"), "Network footer reflects configured limits");
                apply.BringIntoView(); await Layout();
                using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96)))
                { bitmap.Render(window); bitmap.Save(Path.Combine(directory, "script-network-settings.png"), PngBitmapEncoderOptions.Default); }
                window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, UiText.Get("关闭"))).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            await OpenSettingsAsync(SettingsSection.Scripts);
        }
        finally
        {
            _settings.Save(original); UpdateScriptExecutionPolicy(); network.Clear();
        }
    }
}
