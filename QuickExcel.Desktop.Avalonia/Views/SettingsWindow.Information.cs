using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Markdown.Avalonia;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;
using QuickExcel.Presentation;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

internal sealed partial class SettingsWindow
{
    private void ScriptInformation()
    {
        var disableScripts = new CheckBox { Name = "DisableScripts", Content = UiText.Get("禁用脚本"), IsChecked = _settings.Value.DisableScripts };
        AutomationProperties.SetName(disableScripts, UiText.Get("禁用脚本"));
        disableScripts.IsCheckedChanged += (_, _) =>
        {
            Save(_settings.Value with { DisableScripts = disableScripts.IsChecked == true });
            if (_settings.Error is not null) disableScripts.IsChecked = _settings.Value.DisableScripts;
        };
        _page.Children.Add(disableScripts);
        _page.Children.Add(Note(UiText.Get("立即停用所有脚本及编译 worker，包括已签名且受信任的脚本。取消勾选后需重新应用脚本。")));
        _page.Children.Add(Note(UiText.Get("命令行 --no-scripts 或环境变量 QUICKEXCEL_NO_SCRIPTS=1 也会禁用脚本；任一来源生效时均不可执行。命令行及环境变量在启动时读取。")));
        _page.Children.Add(new Separator());
        _page.Children.Add(Heading(UiText.Get("Source Map（源码映射）"), 18));
        var sourceMap = new ComboBox { Name = "ScriptSourceMapMode", HorizontalAlignment = HorizontalAlignment.Stretch,
            ItemsSource = new[] { UiText.Get("无"), UiText.Get("精简"), UiText.Get("完整") }, SelectedIndex = (int)_settings.Value.ScriptSourceMap };
        AutomationProperties.SetName(sourceMap, UiText.Get("Source Map（源码映射）"));
        sourceMap.SelectionChanged += (_, _) =>
        {
            if (sourceMap.SelectedIndex < 0) return;
            Save(_settings.Value with { ScriptSourceMap = (DesktopSourceMapMode)sourceMap.SelectedIndex });
            if (_settings.Error is null) _feedback.Text = UiText.Get("Source Map 设置已保存，关闭设置后生效；请重新应用已启用的脚本。");
        };
        _page.Children.Add(sourceMap);
        _page.Children.Add(Note(UiText.Get("无：不生成映射，运行错误显示生成 JS 的位置。\n精简（默认）：保留原始行列映射，不重复保存源码。\n完整：保留映射及嵌入的原始源码。")));
        _page.Children.Add(Note(UiText.Get("仅保存在本机。更改后当前脚本会停用；下次应用或签名时使用新设置，已有签名保持有效。")));
        _page.Children.Add(Heading(UiText.Format($"脚本 API v{ScriptModuleSet.CurrentApiVersion}"), 18));
        _page.Children.Add(Note(DesktopServices.SupportsScripts
            ? UiText.Format($"JSC 内核    {WorkerProtocol.Engine}\nWorker 协议    {WorkerProtocol.Version}")
            : UiText.Get("当前平台的脚本执行已停用。")));
        _page.Children.Add(Heading(UiText.Get("API 支持矩阵"), 16));
        _page.Children.Add(Note(UiText.Get("下表列出运行时支持的对象成员和枚举值。未列出的 Office Scripts 成员尚未实现。")));
        using var manifest = JsonDocument.Parse(ScriptCapabilities.Current.Json);
        var rows = new List<string>();
        foreach (var type in manifest.RootElement.GetProperty("types").EnumerateObject())
            foreach (var member in type.Value.GetProperty("members").EnumerateObject())
                rows.Add($"{type.Name}.{member.Name}    ·    " + (member.Value.GetProperty("kind").GetString() == "host" ? UiText.Get("宿主调用") : UiText.Get("本地成员")));
        foreach (var type in manifest.RootElement.GetProperty("enums").EnumerateObject())
            foreach (var member in type.Value.EnumerateObject())
                rows.Add(UiText.Format($"{type.Name}.{member.Name}    ·    枚举值 {member.Value.GetInt32()}"));
        rows.Sort(StringComparer.OrdinalIgnoreCase);
        var filter = new TextBox { Name = "SettingsApiSearch", PlaceholderText = UiText.Get("搜索对象或成员，如 Range.getValues") };
        var list = new ListBox { Name = "SettingsApiMatrix", Height = 230 };
        var summary = Note("");
        AutomationProperties.SetName(filter, UiText.Get("搜索 API 支持矩阵")); AutomationProperties.SetName(list, UiText.Get("支持的对象成员和枚举值"));
        void Filter()
        {
            var visible = rows.Where(row => row.Contains(filter.Text?.Trim() ?? "", StringComparison.OrdinalIgnoreCase)).ToArray();
            list.ItemsSource = visible;
            summary.Text = visible.Length == 0 ? UiText.Get("没有匹配的已支持成员。") : UiText.Format($"显示 {visible.Length} / {rows.Count} 项");
        }
        filter.TextChanged += (_, _) => Filter();
        _page.Children.Add(filter); _page.Children.Add(summary); _page.Children.Add(list); Filter();
        _page.Children.Add(Note(UiText.Get("脚本按模块角色自动发现导出的自定义函数、操作和事件处理器；主入口用于初始化。")));
    }

    private void About()
    {
        var assembly = typeof(SettingsWindow).Assembly;
        var version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0];
        var icon = new Image { Name = "AboutProductIcon", Width = 32, Height = 32 };
        icon.Bind(Image.SourceProperty, icon.GetResourceObservable(DesktopApp.ProductIconKey));
        RenderOptions.SetBitmapInterpolationMode(icon, BitmapInterpolationMode.HighQuality);
        var title = Heading("QuickExcel " + version, 20);
        title.VerticalAlignment = VerticalAlignment.Center;
        _page.Children.Add(new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, Children = { icon, title } });
        _page.Children.Add(Note($"{RuntimeInformation.FrameworkDescription}    ·    {RuntimeInformation.ProcessArchitecture}"));
        _page.Children.Add(Note("Copyright in the original code belongs to its respective authors. QuickExcel has not declared a project-wide open-source license. Third-party components are distributed under their respective licenses."));
        var files = LicenseFiles(AppContext.BaseDirectory);
        var selector = new ComboBox { Name = "SettingsLicenseSelector", HorizontalAlignment = HorizontalAlignment.Stretch, ItemsSource = files };
        AutomationProperties.SetName(selector, UiText.Get("选择第三方许可证或来源声明"));
        var viewer = new TextBox { Name = "SettingsLicenseText", IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetName(viewer, UiText.Get("许可证正文"));
        var markdown = new MarkdownScrollViewer { Name = "SettingsLicenseMarkdown", SelectionEnabled = true, Padding = new(12), MarkdownStyle = MarkdownStyle.GithubLike };
        AutomationProperties.SetName(markdown, UiText.Get("许可证正文"));
        var readingArea = new LayoutGrid { Name = "SettingsLicenseViewer", Children = { viewer, markdown } };
        var previous = new Button { Content = UiText.Get("上一页") }; var next = new Button { Content = UiText.Get("下一页") }; var position = Note("");
        position.VerticalAlignment = VerticalAlignment.Center;
        var navigation = new StackPanel { Name = "SettingsLicensePagination", Orientation = Orientation.Horizontal, Spacing = 10, Children = { previous, next, position } };
        const int pageSize = 4000;
        var contents = ""; var page = 0;
        void Display()
        {
            var start = page * pageSize;
            viewer.Text = contents.Substring(start, Math.Min(pageSize, contents.Length - start)); viewer.CaretIndex = 0;
            var pages = Math.Max(1, (contents.Length + pageSize - 1) / pageSize);
            previous.IsEnabled = page > 0; next.IsEnabled = page + 1 < pages;
            position.Text = $"{page + 1} / {pages}";
        }
        selector.SelectionChanged += (_, _) =>
        {
            if (selector.SelectedItem is not string file) return;
            var isMarkdown = Path.GetExtension(file).Equals(".md", StringComparison.OrdinalIgnoreCase);
            try { contents = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, file)); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { contents = ex.Message; isMarkdown = false; }
            markdown.IsVisible = isMarkdown;
            viewer.IsVisible = navigation.IsVisible = !isMarkdown;
            // Parse the whole Markdown document so tables, lists and links are never split across pages.
            markdown.Markdown = isMarkdown ? contents : "";
            markdown.ScrollValue = default;
            page = 0;
            if (isMarkdown) viewer.Text = "";
            else Display();
        };
        previous.Click += (_, _) => { page--; Display(); }; next.Click += (_, _) => { page++; Display(); };
        // Bound the header separately, leaving a reading viewport even at the minimum window size.
        // The header uses its natural height up to the limit set by ResizeAboutHeader.
        _page.Margin = default;
        _pageLayout.Margin = new(24, 20); _pageLayout.RowSpacing = 14;
        _pageLayout.RowDefinitions = new("Auto,Auto,*,Auto");
        LayoutGrid.SetRow(selector, 1); _pageLayout.Children.Add(selector);
        LayoutGrid.SetRow(readingArea, 2); _pageLayout.Children.Add(readingArea);
        LayoutGrid.SetRow(navigation, 3); _pageLayout.Children.Add(navigation);
        selector.SelectedIndex = 0;
    }

    private void ResizeAboutHeader()
    {
        if (_navigation.SelectedIndex == (int)SettingsSection.About && _pageLayout.RowDefinitions.Count == 4)
            _pageLayout.RowDefinitions[0].MaxHeight = Math.Max(80, _pageLayout.Bounds.Height - 200);
    }

    private static string[] LicenseFiles(string directory)
    {
        var files = new List<string>();
        foreach (var folder in new[] { "", "Licenses", "licenses", "JSC-Licenses", "Web/licenses", "Web/cmaps", "Web/standard_fonts", "Web/wasm" })
        {
            var path = Path.Combine(directory, folder);
            if (!Directory.Exists(path)) continue;
            var recursive = folder is "Licenses" or "licenses";
            foreach (var file in Directory.EnumerateFiles(path, "*", recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileName(file);
                if (name.StartsWith("LICENSE", StringComparison.OrdinalIgnoreCase) ||
                    (Path.GetExtension(name).ToLowerInvariant() is ".md" or ".txt") &&
                    (folder.Length > 0 || name.Contains("LICENSE", StringComparison.OrdinalIgnoreCase) || name.Contains("NOTICE", StringComparison.OrdinalIgnoreCase)))
                    files.Add(Path.GetRelativePath(directory, file));
            }
        }
        return files.Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(f => Path.GetFileName(f) == "THIRD-PARTY-LICENSES.md").ThenBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
    }
}
