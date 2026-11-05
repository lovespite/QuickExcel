using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;

namespace QuickExcel.Desktop.Avalonia;

internal sealed partial class SettingsWindow
{
    private void ScriptInformation()
    {
        _page.Children.Add(Heading(UiText.Format($"脚本 API v{ScriptCapabilities.ApiVersion}"), 18));
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
        var filter = new TextBox { PlaceholderText = UiText.Get("搜索对象或成员，如 Range.getValues") };
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
        _page.Children.Add(Note(UiText.Get("宿主扩展入口包括 host.functions.register（自定义函数）和 host.events.on（事件）。")));
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
        _page.Children.Add(Heading("Third-Party Components, Copyright and Licenses", 16));
        _page.Children.Add(Note("Ribbon icons: Fluent UI System Icons — Copyright (c) 2020 Microsoft Corporation (MIT). Selected SVG paths are converted to native geometry; source details are in THIRD-PARTY-LICENSES.md."));
        var files = LicenseFiles(AppContext.BaseDirectory);
        var selector = new ComboBox { HorizontalAlignment = HorizontalAlignment.Stretch, ItemsSource = files };
        AutomationProperties.SetName(selector, UiText.Get("选择第三方许可证或来源声明"));
        var viewer = new TextBox { Name = "SettingsLicenseText", IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, Height = 210 };
        AutomationProperties.SetName(viewer, UiText.Get("许可证正文"));
        var previous = new Button { Content = UiText.Get("上一页") }; var next = new Button { Content = UiText.Get("下一页") }; var position = Note("");
        position.VerticalAlignment = VerticalAlignment.Center;
        var navigation = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, Children = { previous, next, position } };
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
            try { contents = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, file)); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { contents = ex.Message; }
            page = 0; Display();
        };
        previous.Click += (_, _) => { page--; Display(); }; next.Click += (_, _) => { page++; Display(); };
        _page.Children.Add(selector); _page.Children.Add(viewer); _page.Children.Add(navigation);
        selector.SelectedIndex = 0;
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
