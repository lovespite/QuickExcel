using System.Text;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptNetworkPanelAsync(Action<bool, string> check, string directory)
    {
        static async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var network = Document.Session.Network;
        network.Revoke();
        async Task CheckConsent(Window window)
        {
            await Layout();
            var english = UiText.Language == UiText.English;
            var addresses = window.GetVisualDescendants().OfType<ListBox>().Single();
            check(addresses.Items.Cast<string>().SequenceEqual(new[] { "127.0.0.1", "::1", "api.example.com" }),
                "Network consent shows the complete sorted domain and IP list");
            var captions = ((Panel)window.Content!).Children.OfType<TextBlock>().Select(block => block.Text);
            check(window.Title == (english ? "Script network access" : "脚本网络访问授权") &&
                AutomationProperties.GetName(addresses) == (english ? "Requested network addresses" : "请求授权的网络地址") &&
                captions.SequenceEqual(english
                    ? new[] { "Actions in this workbook may access the following domains or IP addresses.", "Permission applies only to this session. Scripts can send workbook data to the listed addresses." }
                    : new[] { "此工作簿的 Actions 脚本可能访问以下域名或 IP 地址：", "授权仅限本次会话。脚本可向上述地址发送工作簿数据。" }),
                $"Network consent title, address list label and explanation use {UiText.Language}");
            var buttons = window.GetVisualDescendants().OfType<Button>().ToArray();
            var deny = buttons.Single(button => button.IsCancel);
            check(Equals(deny.Content, english ? "Deny" : "拒绝访问") &&
                buttons.Any(button => Equals(button.Content, english ? "Allow for this session" : "允许本次会话")),
                $"Network consent buttons use {UiText.Language}");
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96)))
            { bitmap.Render(window); bitmap.Save(Path.Combine(directory, english ? "script-network-consent-en.png" : "script-network-consent.png"), PngBitmapEncoderOptions.Default); }
            deny.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }
        var priorLanguage = UiText.Language;
        try
        {
            foreach (var language in new[] { UiText.Chinese, UiText.English })
            {
                UiText.Initialize(language);
                _dialogDriver = CheckConsent;
                var denied = false;
                try { await Task.Run(() => network.EnsureAuthorized(["api.example.com", "127.0.0.1", "::1"], default)); }
                catch (InvalidOperationException error) when (error.Message.Contains("未授权")) { denied = true; }
                check(denied, $"Rejecting the {language} network consent dialog prevents execution authorization");
            }
        }
        finally { UiText.Initialize(priorLanguage); }
        await CloseScriptWindowAsync();
        const string initialization = "host.permissions.network.request('api.example.com'); export function main(w) { const cell = w.getActiveWorksheet().getRange('D1'); cell.setValue((Number(cell.getValue()) || 0) + 1); }";
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", initialization),
            ["action"] = new("scripts/action.js", "export function main() {}")
        }, ["main", "action"], scriptApiVersion: 4, bindings: [new("main", "initialization"), new("action", "actions")])));
        await OpenScriptsAsync();
        var prompts = new List<string[]>();
        string ConsentState() => $"Enabled={Document.ScriptsEnabled}; prompts={prompts.Count} [{string.Join("; ", prompts.Select(hosts => string.Join(", ", hosts)))}]; status={_scriptStatus?.Text}; documentStatus={_status.Text}; source={Document.Snapshot.ScriptModules.Modules["main"].Source}; errors={string.Join("; ", Document.ScriptErrors)}";
        Task AllowNetwork(Window window)
        {
            prompts.Add(window.GetVisualDescendants().OfType<ListBox>().Single().Items.Cast<string>().ToArray());
            // ShowOwnedAsync consumes its driver once; keep counting every prompt,
            // including an unexpected prompt on a repeated load.
            _dialogDriver = AllowNetwork;
            window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == "允许本次会话").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            return Task.CompletedTask;
        }
        _dialogDriver = AllowNetwork;
        await ApplyScriptsAsync();
        check(Document.ScriptsEnabled && prompts.Count == 1 && prompts[0].SequenceEqual(new[] { "api.example.com" }),
            "Loading scripts requests the declared network list once before initialization");
        await ApplyScriptsAsync();
        check(Document.ScriptsEnabled && prompts.Count == 1 && Document.GetCellValue(0, 0, 3) == "2" && _scriptWindow!.OwnedWindows.Count == 0,
            "Repeated loading reruns initialization without repeating unchanged session network consent");
        await _scriptPage!.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?>
            { ["action"] = "edit", ["module"] = "main", ["text"] = initialization + "\n// edited without changing the network list\n" }, TimeSpan.FromSeconds(30));
        check(!Document.ScriptsEnabled, "Editing the actual Monaco draft invalidates script execution");
        await ApplyScriptsAsync();
        check(Document.ScriptsEnabled && prompts.Count == 1 && Document.GetCellValue(0, 0, 3) == "3",
            "Loading an edited draft reuses network consent when its declared list is unchanged");
        var changedDeclaration = "host.permissions.network.request('127.0.0.1');\n" + initialization;
        await _scriptPage.Browser.RequestAsync("test", "testResult", new Dictionary<string, object?>
            { ["action"] = "edit", ["module"] = "main", ["text"] = changedDeclaration }, TimeSpan.FromSeconds(30));
        var changedCommitted = await CommitScriptsAsync();
        check(changedCommitted && Document.Snapshot.ScriptModules.Modules["main"].Source.ReplaceLineEndings("\n") == changedDeclaration,
            "Changed network declaration commits from the actual Monaco draft; " + ConsentState());
        await ApplyScriptsAsync();
        await File.WriteAllTextAsync(Path.Combine(directory, "script-network-reload-state.json"), System.Text.Json.JsonSerializer.Serialize(new
        {
            prompts, enabled = Document.ScriptsEnabled, source = Document.Snapshot.ScriptModules.Modules["main"].Source,
            scriptStatus = _scriptStatus?.Text, documentStatus = _status.Text, errors = Document.ScriptErrors,
            console = Document.Session.ConsoleEntries.TakeLast(10).Select(entry => entry.Text).ToArray()
        }));
        check(Document.ScriptsEnabled && prompts.Count == 2 && prompts[1].SequenceEqual(new[] { "127.0.0.1", "api.example.com" }),
            "Loading a changed declaration asks for the complete updated network list; " + ConsentState());
        DisableScripts(); await ApplyScriptsAsync();
        check(Document.ScriptsEnabled && prompts.Count == 3,
            "Explicitly disabling scripts revokes network consent before the next load");
        _dialogDriver = null;
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
