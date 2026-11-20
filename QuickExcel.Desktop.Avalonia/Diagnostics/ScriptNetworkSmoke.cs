using System.Text;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptNetworkPanelAsync(Action<bool, string> check, string directory)
    {
        static async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var network = Document.Session.Network;
        network.Revoke();
        _dialogDriver = async window =>
        {
            await Layout();
            var addresses = window.GetVisualDescendants().OfType<ListBox>().Single();
            check(addresses.Items.Cast<string>().SequenceEqual(new[] { "127.0.0.1", "::1", "api.example.com" }),
                "Network consent shows the complete sorted domain and IP list");
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96)))
            { bitmap.Render(window); bitmap.Save(Path.Combine(directory, "script-network-consent.png"), PngBitmapEncoderOptions.Default); }
            window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == "拒绝").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        };
        var denied = false;
        try { await Task.Run(() => network.EnsureAuthorized(["api.example.com", "127.0.0.1", "::1"], default)); }
        catch (InvalidOperationException error) when (error.Message.Contains("未授权")) { denied = true; }
        check(denied, "Rejecting the actual network consent dialog prevents execution authorization");
        _dialogDriver = window =>
        {
            window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == "允许本次会话").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            return Task.CompletedTask;
        };
        await Task.Run(() => network.EnsureAuthorized(["api.example.com"], default));
        await Task.Run(() => network.EnsureAuthorized(["api.example.com"], default));
        check(_scriptWindow!.OwnedWindows.Count == 0, "Explicit network consent is reused within the same document session");
        network.Revoke();
        check(!network.RecordingEnabled, "Network capture is disabled while the console tab is selected");
        SetScriptOutputVisible(true, tab: 2); await Layout();
        check(network.RecordingEnabled, "Opening the Network tab enables capture for this document");
        var entry = new ScriptNetworkEntry { Url = "https://api.example.com/report", ResponseUrl = "https://api.example.com/report", Method = "POST",
            Status = 200, StatusText = "OK", Duration = TimeSpan.FromMilliseconds(125),
            RequestHeaders = [new("content-type", "application/json")], ResponseHeaders = [new("content-type", "application/json")],
            RequestBody = Encoding.UTF8.GetBytes("{\"query\":\"example\"}"), ResponseBody = Encoding.UTF8.GetBytes("{\"message\":\"成功\",\"rows\":[1,2,3]}") };
        network.Record(entry); await Layout();
        _networkRequests!.SelectedItem = entry; _networkDetailTabs!.SelectedIndex = 2; await Layout();
        check(_networkRequests.ItemCount == 1 && _networkDetails[2].Text!.Contains("rows") && _networkDetails[2].Text!.Contains('\n'),
            "Network rows expose selectable formatted JSON previews");
        _networkDetailTabs.SelectedIndex = 0;
        check(_networkDetails[2].Text == "" && _networkDetails[0].Text!.Contains("content-type"), "Switching detail tabs releases the hidden response text");
        _networkFilter!.Text = "missing"; await Layout();
        check(_networkRequests.ItemCount == 0 && _networkDetails.All(d => d.Text == ""), "Network filtering clears stale selections and details");
        _networkFilter.Text = "report"; await Layout(); _networkRequests.SelectedItem = entry; _networkDetailTabs.SelectedIndex = 2; await Layout();
        using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(_scriptOutputTabs!.Bounds.Width), (int)Math.Ceiling(_scriptOutputTabs.Bounds.Height)), new Vector(96, 96)))
        { bitmap.Render(_scriptOutputTabs); bitmap.Save(Path.Combine(directory, "script-network.png"), PngBitmapEncoderOptions.Default); }
        SetScriptOutputVisible(false); check(!network.RecordingEnabled, "Collapsing output immediately stops network capture");
        SetScriptOutputVisible(true, tab: 2);
        _scriptWindow!.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "清空网络记录并释放响应内容").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        await Layout(); check(network.Entries.IsEmpty && entry.ResponseBody is null && _networkDetails.All(d => d.Text == ""), "Clear releases retained bodies and visible previews");
        await CheckNetworkSettingsAsync(check, directory);
        SetScriptOutputVisible(true, tab: 0, focusInput: true);
    }
}
