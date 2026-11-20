using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using QuickExcel.Core;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private ScriptNetworkSession? _networkSession;
    private ListBox? _networkRequests;
    private TextBox? _networkFilter;
    private TextBlock? _networkSummary;
    private TabControl? _networkDetailTabs;
    private readonly List<TextBox> _networkDetails = [];
    private int _networkRefreshPending;
    private const int NetworkPreviewCharacters = 128 * 1024;
    private const string NetworkColumns = "*,40,55,60";

    private void BindScriptNetwork()
    {
        var session = Document.Session.Network;
        session.Configure(_settings.Value.ScriptNetwork);
        if (ReferenceEquals(session, _networkSession)) return;
        if (_networkSession is { } previous)
        { previous.RecordingEnabled = false; previous.Changed -= NetworkChanged; previous.Authorize = null; }
        _networkSession = session;
        session.Authorize = (hosts, cancellation) => AuthorizeScriptNetwork(session, hosts, cancellation);
        session.Changed += NetworkChanged;
        UpdateNetworkRecording(); RefreshScriptNetwork();
    }

    private bool AuthorizeScriptNetwork(ScriptNetworkSession session, IReadOnlyList<string> hosts, CancellationToken cancellation)
    {
        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        Dispatcher.UIThread.Post(async () =>
        {
            try
            {
                if (cancellation.IsCancellationRequested || !ReferenceEquals(session, Document.Session.Network))
                { completion.TrySetCanceled(cancellation); return; }
                var window = new Window { Title = UiText.Get("脚本网络访问授权"), Width = 540, Height = 430, ShowInTaskbar = false,
                    MinWidth = 420, MinHeight = 330, WindowStartupLocation = WindowStartupLocation.CenterOwner };
                var root = new LayoutGrid { RowDefinitions = new("Auto,*,Auto,Auto"), Margin = new(20) };
                root.Children.Add(new TextBlock { Text = UiText.Get("此工作簿的 Actions 脚本可能访问以下域名或 IP 地址："),
                    TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 14) });
                var addresses = new ListBox { ItemsSource = hosts, FontFamily = new("Consolas"), FontSize = 14 };
                AutomationProperties.SetName(addresses, UiText.Get("请求授权的网络地址"));
                LayoutGrid.SetRow(addresses, 1); root.Children.Add(addresses);
                var note = new TextBlock { Text = UiText.Get("授权仅用于当前工作簿会话。地址清单不变时，重复加载无需再次授权；清单变化、停用脚本或关闭网络访问后需要重新授权。脚本可以向这些地址发送工作簿数据；只允许列出的地址，重定向也会检查。"),
                    TextWrapping = TextWrapping.Wrap, Margin = new(0, 12) };
                LayoutGrid.SetRow(note, 2); root.Children.Add(note);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 10 };
                var deny = new Button { Content = UiText.Get("拒绝"), IsCancel = true };
                var allow = new Button { Content = UiText.Get("允许本次会话"), IsDefault = false };
                deny.Click += (_, _) => window.Close(false); allow.Click += (_, _) => window.Close(true);
                buttons.Children.Add(deny); buttons.Children.Add(allow); LayoutGrid.SetRow(buttons, 3); root.Children.Add(buttons);
                window.Content = root;
                using var registration = cancellation.Register(() => Dispatcher.UIThread.Post(() => window.Close(false)));
                var accepted = await ShowOwnedAsync<bool>(window, _scriptWindow ?? this);
                completion.TrySetResult(accepted && !cancellation.IsCancellationRequested && ReferenceEquals(session, Document.Session.Network));
            }
            catch (Exception error) { completion.TrySetException(error); }
        });
        return completion.Task.WaitAsync(cancellation).GetAwaiter().GetResult();
    }

    private Control CreateScriptNetworkPanel()
    {
        BindScriptNetwork();
        var panel = new LayoutGrid { RowDefinitions = new("Auto,*,Auto"), Background = ConsoleBackground };
        var toolbar = new LayoutGrid { ColumnDefinitions = new("Auto,*,Auto"), Margin = new(5, 3) };
        var clear = ConsoleButton("⊘", UiText.Get("清空网络记录并释放响应内容"));
        clear.Click += (_, _) => _networkSession?.Clear(); toolbar.Children.Add(clear);
        _networkFilter = new TextBox { Height = 26, MinHeight = 0, Padding = new(7, 3), Margin = new(5, 0), FontSize = 12,
            PlaceholderText = UiText.Get("过滤 URL / 方法 / 状态"), Background = ConsoleBackground, Foreground = ConsoleInk, BorderBrush = ConsoleLine };
        AutomationProperties.SetName(_networkFilter, UiText.Get("过滤网络请求"));
        _networkFilter.TextChanged += (_, _) => RefreshScriptNetwork(); LayoutGrid.SetColumn(_networkFilter, 1); toolbar.Children.Add(_networkFilter);
        var capture = new TextBlock { Text = UiText.Get("● 面板打开时记录"), Foreground = ConsoleBlue, FontSize = 11, VerticalAlignment = VerticalAlignment.Center, Margin = new(8, 0) };
        LayoutGrid.SetColumn(capture, 2); toolbar.Children.Add(capture); panel.Children.Add(toolbar);
        var body = new LayoutGrid { ColumnDefinitions = new("45*,5,55*") };
        var requests = new DockPanel();
        var heading = new LayoutGrid { ColumnDefinitions = new(NetworkColumns), Margin = new(10, 4) };
        var labels = new[] { "名称 / 方法", "状态", "大小", "耗时" };
        for (var column = 0; column < labels.Length; column++)
        {
            var label = new TextBlock { Text = UiText.Get(labels[column]), FontSize = 11, Foreground = ConsoleMuted };
            LayoutGrid.SetColumn(label, column); heading.Children.Add(label);
        }
        DockPanel.SetDock(heading, Dock.Top); requests.Children.Add(heading);
        _networkRequests = new ListBox { Background = ConsoleBackground, BorderThickness = new(0), Padding = new(0),
            ItemTemplate = new FuncDataTemplate<ScriptNetworkEntry>((entry, _) => entry is null ? null : CreateNetworkRow(entry)) };
        _networkRequests.Styles.Add(new Style(x => x.OfType<ListBoxItem>()) { Setters = {
            new Setter(TemplatedControl.PaddingProperty, new Thickness(0)), new Setter(Layoutable.MarginProperty, new Thickness(0)),
            new Setter(ContentControl.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch) } });
        AutomationProperties.SetName(_networkRequests, UiText.Get("网络请求列表"));
        _networkRequests.SelectionChanged += (_, _) => RefreshNetworkDetails(); requests.Children.Add(_networkRequests); body.Children.Add(requests);
        var divider = new GridSplitter { ResizeDirection = GridResizeDirection.Columns, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, Background = ConsoleLine };
        LayoutGrid.SetColumn(divider, 1); body.Children.Add(divider);
        _networkDetails.Clear(); var tabs = new List<TabItem>();
        foreach (var name in new[] { "标头", "载荷", "预览", "响应", "计时" })
        {
            var text = new TextBox { IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap,
                FontFamily = new("Consolas"), FontSize = 12, BorderThickness = new(0), Background = ConsoleBackground, Foreground = ConsoleInk,
                Padding = new(10, 6) };
            AutomationProperties.SetName(text, UiText.Get("网络请求") + " · " + UiText.Get(name));
            ScrollViewer.SetVerticalScrollBarVisibility(text, ScrollBarVisibility.Auto);
            _networkDetails.Add(text); tabs.Add(new TabItem { Header = UiText.Get(name), FontSize = 12, Padding = new(8, 5), Content = text });
        }
        _networkDetailTabs = new TabControl { ItemsSource = tabs, Background = ConsoleBackground, Foreground = ConsoleInk };
        _networkDetailTabs.SelectionChanged += (_, _) => RefreshNetworkDetails();
        LayoutGrid.SetColumn(_networkDetailTabs, 2); body.Children.Add(_networkDetailTabs); LayoutGrid.SetRow(body, 1); panel.Children.Add(body);
        _networkSummary = new TextBlock { Foreground = ConsoleMuted, FontSize = 11, Margin = new(8, 4) };
        LayoutGrid.SetRow(_networkSummary, 2); panel.Children.Add(_networkSummary);
        RefreshScriptNetwork(); return panel;
    }

    private static Control CreateNetworkRow(ScriptNetworkEntry entry)
    {
        var row = new LayoutGrid { ColumnDefinitions = new(NetworkColumns), Margin = new(10, 2) };
        var name = Uri.TryCreate(entry.Url, UriKind.Absolute, out var url) ? url.AbsolutePath.Split('/').LastOrDefault() : entry.Url;
        if (string.IsNullOrEmpty(name)) name = url?.Host ?? entry.Url;
        var title = new TextBlock { Text = entry.Method + "  " + name, TextTrimming = TextTrimming.CharacterEllipsis, FontSize = 12 };
        ToolTip.SetTip(title, entry.Url); row.Children.Add(title);
        var status = new TextBlock { Text = entry.Status?.ToString() ?? "×", FontSize = 12,
            Foreground = entry.Error is not null || entry.Status >= 400 ? ConsoleErrorInk : ConsoleInk };
        LayoutGrid.SetColumn(status, 1); row.Children.Add(status);
        var size = new TextBlock { Text = NetworkSize(entry.ResponseBytes), FontSize = 11, Foreground = ConsoleMuted };
        LayoutGrid.SetColumn(size, 2); row.Children.Add(size);
        var duration = new TextBlock { Text = entry.Duration.TotalMilliseconds.ToString("0") + " ms", FontSize = 11, Foreground = ConsoleBlue };
        LayoutGrid.SetColumn(duration, 3); row.Children.Add(duration);
        return row;
    }

    private void UpdateNetworkRecording()
    {
        if (_networkSession is not null)
            _networkSession.RecordingEnabled = _scriptWindow is not null && _scriptOutputTabs?.IsVisible == true && _scriptOutputTabs.SelectedIndex == 2;
    }

    private void NetworkChanged()
    {
        // Coalesce notifications without capturing request bodies in dispatcher closures.
        if (Interlocked.Exchange(ref _networkRefreshPending, 1) != 0) return;
        Dispatcher.UIThread.Post(() => { Interlocked.Exchange(ref _networkRefreshPending, 0); RefreshScriptNetwork(); });
    }

    private void RefreshScriptNetwork()
    {
        if (_networkRequests is null || _networkSession is null) return;
        var selected = (_networkRequests.SelectedItem as ScriptNetworkEntry)?.Id;
        var entries = _networkSession.Entries;
        var filter = _networkFilter?.Text?.Trim() ?? "";
        var visible = entries.Where(e => filter.Length == 0 || e.Url.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
            e.Method.Contains(filter, StringComparison.OrdinalIgnoreCase) || (e.Status?.ToString() ?? "").Contains(filter, StringComparison.OrdinalIgnoreCase)).ToArray();
        _networkRequests.ItemsSource = visible;
        _networkRequests.SelectedItem = visible.FirstOrDefault(e => e.Id == selected);
        var options = _networkSession.Options;
        if (_networkSummary is not null) _networkSummary.Text = UiText.Format($"{visible.Length} / {entries.Length} 个请求 · 保留内存 {NetworkSize(_networkSession.RetainedMemoryBytes)} / {options.PanelMemoryLimitMiB} MiB · 最多 {options.PanelEntryLimit} 条，超限释放旧响应");
        RefreshNetworkDetails();
    }

    private void RefreshNetworkDetails()
    {
        // Hidden tabs and evicted selections must not retain copies of old response text.
        foreach (var detail in _networkDetails) detail.Text = "";
        if (_networkRequests?.SelectedItem is not ScriptNetworkEntry entry || _networkDetailTabs is null || _networkDetails.Count != 5) return;
        var index = _networkDetailTabs.SelectedIndex;
        if (index is < 0 or > 4) return;
        var text = index switch
        {
            0 => "Request URL: " + entry.Url + "\nRequest Method: " + entry.Method + "\nStatus: " + entry.Status + " " + entry.StatusText +
                "\nResponse URL: " + entry.ResponseUrl + (entry.Error is null ? "" : "\nError: " + entry.Error) +
                "\n\n" + UiText.Get("请求标头") + "\n" + NetworkHeaders(entry.RequestHeaders) + "\n\n" + UiText.Get("响应标头") + "\n" + NetworkHeaders(entry.ResponseHeaders),
            1 => entry.BodiesEvicted ? UiText.Get("内容已超过保留上限，已释放。") : NetworkBodyText(entry.RequestBody, false),
            2 => entry.BodiesEvicted ? UiText.Get("响应已超过保留上限，已释放。") : NetworkBodyText(entry.ResponseBody, true),
            3 => entry.BodiesEvicted ? UiText.Get("响应已超过保留上限，已释放。") : NetworkBodyText(entry.ResponseBody, false),
            _ => UiText.Get("开始时间") + ": " + entry.Started.ToLocalTime().ToString("HH:mm:ss.fff") + "\n" + UiText.Get("总耗时") + ": " + entry.Duration.TotalMilliseconds.ToString("0.0") + " ms\n" +
                UiText.Get("请求内容") + ": " + NetworkSize(entry.RequestBytes) + "\n" + UiText.Get("响应内容") + ": " + NetworkSize(entry.ResponseBytes)
        };
        _networkDetails[index].Text = text.Length > NetworkPreviewCharacters ? text[..NetworkPreviewCharacters] + "\n…" + UiText.Get("预览已截断") : text;
    }

    private static string NetworkHeaders(IReadOnlyList<KeyValuePair<string, string>> headers) => string.Join("\n", headers.Select(h => h.Key + ": " + h.Value));
    private static string NetworkSize(long bytes) => bytes >= 1024 * 1024 ? (bytes / 1048576d).ToString("0.0") + " MB" : bytes >= 1024 ? (bytes / 1024d).ToString("0.0") + " kB" : bytes + " B";
    private static string NetworkBodyText(byte[]? bytes, bool preview)
    {
        if (bytes is null || bytes.Length == 0) return UiText.Get("无内容");
        var text = Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, NetworkPreviewCharacters));
        if (preview && bytes.Length <= NetworkPreviewCharacters)
        {
            try
            {
                using var parsed = JsonDocument.Parse(bytes, new() { MaxDepth = 32 });
                using var stream = new MemoryStream();
                using (var writer = new Utf8JsonWriter(stream, new() { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping })) parsed.RootElement.WriteTo(writer);
                if (stream.Length <= NetworkPreviewCharacters) text = Encoding.UTF8.GetString(stream.GetBuffer(), 0, (int)stream.Length);
            }
            catch (JsonException) { }
        }
        return text + (bytes.Length > NetworkPreviewCharacters ? "\n…" + UiText.Get("预览已截断") : "");
    }

    private void CloseScriptNetworkPanel()
    {
        if (_networkSession is not null) _networkSession.RecordingEnabled = false;
        foreach (var detail in _networkDetails) detail.Text = "";
        _networkDetails.Clear();
        if (_networkRequests is not null) _networkRequests.ItemsSource = null;
        _networkRequests = null; _networkFilter = null; _networkSummary = null; _networkDetailTabs = null;
    }
}
