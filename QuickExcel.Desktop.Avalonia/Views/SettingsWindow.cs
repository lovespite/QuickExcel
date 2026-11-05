using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Microsoft.Win32;
using QuickExcel.Presentation;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Task OpenSettingsAsync() => OpenSettingsAsync(SettingsSection.Associations);

    private async Task OpenSettingsAsync(SettingsSection section)
    {
        if (!await CommitEditorsAsync()) return;
        await ShowOwnedAsync<bool>(new SettingsWindow(_settings, _recentFiles, section));
        ApplySheetPlacement();
        ScheduleCalculation();
        RefreshRecentFiles();
    }
}

internal enum SettingsSection { Associations, History, Calculation, Appearance, Scripts, About }

internal sealed partial class SettingsWindow : Window
{
    private readonly DesktopSettings _settings;
    private readonly RecentFileStore _history;
    private readonly ListBox _navigation = new() { Name = "SettingsNavigation", Background = Brushes.Transparent };
    private readonly StackPanel _page = new() { Spacing = 14, Margin = new(24, 20) };
    private readonly TextBlock _feedback = new() { TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };

    public SettingsWindow(DesktopSettings settings, RecentFileStore history, SettingsSection initialSection = SettingsSection.Associations)
    {
        _settings = settings; _history = history;
        FontFamily = new FontFamily("Segoe UI, Microsoft YaHei UI, Noto Sans CJK SC, sans-serif");
        Title = UiText.Get("设置"); Width = 820; Height = 600; MinWidth = 580; MinHeight = 380;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var layout = new LayoutGrid { ColumnDefinitions = new("156,*"), RowDefinitions = new("*,Auto") };
        var nav = new DockPanel { Background = Brush.Parse("#F1F6F3") };
        var title = Heading(UiText.Get("设置"), 24); title.Margin = new(16, 20, 12, 16); DockPanel.SetDock(title, Dock.Top);
        nav.Children.Add(title); nav.Children.Add(_navigation); layout.Children.Add(nav);
        var scroll = new ScrollViewer { Content = _page, HorizontalScrollBarVisibility = global::Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
        LayoutGrid.SetColumn(scroll, 1); layout.Children.Add(scroll);
        var footer = new DockPanel { Margin = new(16, 8) };
        var close = new Button { Content = UiText.Get("关闭"), MinWidth = 76, IsCancel = true, HorizontalContentAlignment = HorizontalAlignment.Center };
        close.Click += (_, _) => Close(true); DockPanel.SetDock(close, Dock.Right);
        footer.Children.Add(close); footer.Children.Add(_feedback); LayoutGrid.SetRow(footer, 1); LayoutGrid.SetColumnSpan(footer, 2);
        layout.Children.Add(footer); Content = layout;
        (SettingsSection Id, string Title, Action Build)[] sections = [
            (SettingsSection.Associations, UiText.Get("文件关联"), Associations), (SettingsSection.History, UiText.Get("历史管理"), History),
            (SettingsSection.Calculation, UiText.Get("保存与计算"), Calculation), (SettingsSection.Appearance, UiText.Get("界面"), Appearance),
            (SettingsSection.Scripts, UiText.Get("脚本与 API"), ScriptInformation), (SettingsSection.About, UiText.Get("关于与许可"), About)];
        _navigation.ItemsSource = sections.Select(s => s.Title).ToArray();
        AutomationProperties.SetName(_navigation, UiText.Get("设置分区"));
        _navigation.SelectionChanged += (_, _) =>
        {
            if (_navigation.SelectedIndex < 0) return;
            var section = sections[_navigation.SelectedIndex];
            _page.Children.Clear(); _page.Children.Add(Heading(section.Title, 22)); _feedback.Text = "";
            scroll.Offset = default;
            try { section.Build(); }
            catch (Exception ex) { _feedback.Text = ex.Message; }
        };
        _navigation.SelectedIndex = Array.FindIndex(sections, section => section.Id == initialSection);
    }

    private static TextBlock Heading(string text, double size) => new() { Text = text, FontSize = size, FontWeight = FontWeight.SemiBold };
    private static TextBlock Note(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap, Foreground = Brush.Parse("#46554E") };
    private void Save(DesktopPreferences preferences)
    {
        _settings.Save(preferences);
        _feedback.Text = _settings.Error is { } error ? UiText.Get("保存设置失败：") + error : UiText.Get("设置已保存");
    }

    private void Associations()
    {
        if (OperatingSystem.IsWindows()) WindowsAssociations();
        else _page.Children.Add(Note(UiText.Get("请在系统的文件管理器中选择默认打开应用。")));
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private void WindowsAssociations()
    {
        _page.Children.Add(Note(UiText.Get("选择要用 QuickExcel 打开的文件类型，再到 Windows「默认应用」确认关联。")));
        var selected = WindowsFileAssociations.ReadSelection(Registry.CurrentUser);
        var boxes = new List<CheckBox>();
        foreach (var ext in WindowsFileAssociations.Extensions)
        {
            var description = ext switch { ".xlsq" => UiText.Get("QuickExcel 工作簿"), ".xlsx" => UiText.Get("Excel 工作簿"), _ => UiText.Get("逗号分隔文本（UTF-8）") };
            var box = new CheckBox { Content = ext + "    " + description, Tag = ext, IsChecked = selected.Contains(ext) };
            boxes.Add(box); _page.Children.Add(box);
        }
        _page.Children.Add(Note(UiText.Get("CSV 按文本导入，保存时另存为工作簿。")));
        var apply = new Button { Content = UiText.Get("注册所选类型并打开默认应用") }; _page.Children.Add(apply);
        apply.Click += (_, _) =>
        {
            try
            {
                WindowsFileAssociations.ApplyAndOpenDefaults(boxes.Where(b => b.IsChecked == true).Select(b => (string)b.Tag!).ToArray());
                _feedback.Text = UiText.Get("已更新文件类型，请在 Windows 设置中确认默认应用。");
            }
            catch (Exception ex) { _feedback.Text = ex.Message; }
        };
    }

    private void History()
    {
        _history.Load();
        var list = new ListBox { Name = "SettingsRecentFiles", MaxHeight = 300 };
        AutomationProperties.SetName(list, UiText.Get("最近使用的文件路径"));
        var empty = Note(UiText.Get("暂无历史文件。"));
        var remove = new Button { Content = UiText.Get("移除选中记录") };
        var clear = new Button { Content = UiText.Get("清空历史记录") };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Children = { remove, clear } };
        void Refresh()
        {
            list.ItemsSource = _history.Paths;
            empty.IsVisible = _history.Paths.Count == 0;
            list.IsVisible = actions.IsVisible = !empty.IsVisible;
            remove.IsEnabled = list.SelectedItem is string;
            _feedback.Text = _history.Error;
        }
        list.SelectionChanged += (_, _) => remove.IsEnabled = list.SelectedItem is string;
        remove.Click += (_, _) => { if (list.SelectedItem is string path) _history.Remove(path); Refresh(); };
        clear.Click += async (_, _) => { if (await Dialogs.Confirm(this, UiText.Get("清空历史文件记录"), UiText.Get("移除列表中的全部记录？磁盘上的文件会保留。"))) { _history.Clear(); Refresh(); } };
        _page.Children.Add(Note(UiText.Get("管理开始页上显示的最近文件。移除记录会保留磁盘上的文件。")));
        _page.Children.Add(list); _page.Children.Add(empty); _page.Children.Add(actions); Refresh();
    }

    private void Calculation()
    {
        _page.Children.Add(Heading(UiText.Get("工作簿计算"), 16));
        (DesktopCalculationMode Mode, string Label)[] modes = [
            (DesktopCalculationMode.Automatic, UiText.Get("自动 · 编辑后重算公式和模拟运算表")),
            (DesktopCalculationMode.Partial, UiText.Get("部分 · 普通公式自动重算，模拟运算表单独刷新")),
            (DesktopCalculationMode.Manual, UiText.Get("手动 · 按 F9 或保存前重算"))];
        foreach (var (mode, label) in modes)
        {
            var choice = new RadioButton { Content = label, GroupName = "Calculation", IsChecked = _settings.Value.Calculation == mode };
            choice.IsCheckedChanged += (_, _) => { if (choice.IsChecked == true) Save(_settings.Value with { Calculation = mode }); };
            _page.Children.Add(choice);
        }
        _page.Children.Add(Note(UiText.Get("通过「数据 → 模拟分析 → 刷新模拟运算表」更新本次打开期间创建的表。文件中保存试算结果。重新打开、改写输出区或调整表结构后，需重新创建模拟运算表。")));
        _page.Children.Add(new Separator());
        _page.Children.Add(Heading(UiText.Get("自动保存"), 16));
        var autosave = new ToggleSwitch { Name = "AutoSavePlaceholder", Content = UiText.Get("自动保存工作簿"), OffContent = UiText.Get("关"), IsEnabled = false, IsChecked = false };
        _page.Children.Add(autosave); _page.Children.Add(Note(UiText.Get("此功能尚未开放。请使用 Ctrl+S 保存。")));
    }

    private void Appearance()
    {
        var label = UiText.Get("界面语言");
        _page.Children.Add(Heading(label, 16));
        var language = new ComboBox { Name = "DisplayLanguage", ItemsSource = UiText.Languages,
            SelectedItem = UiText.Resolve(_settings.Value.Language), MinWidth = 180 };
        AutomationProperties.SetName(language, label);
        language.SelectionChanged += (_, _) =>
        {
            if (language.SelectedItem is not DisplayLanguage selected || selected.Tag == _settings.Value.Language) return;
            Save(_settings.Value with { Language = selected.Tag });
            if (_settings.Error is not null) language.SelectedItem = UiText.Resolve(_settings.Value.Language);
        };
        _page.Children.Add(language);
        _page.Children.Add(Note(UiText.Get("重启 QuickExcel 后应用所选语言。")));
        _page.Children.Add(new Separator());
        _page.Children.Add(Heading(UiText.Get("工作表标签位置"), 16));
        var placement = new ComboBox { ItemsSource = new[] { UiText.Get("底部"), UiText.Get("左侧"), UiText.Get("右侧") }, SelectedIndex = (int)_settings.Value.Sheets, MinWidth = 180 };
        AutomationProperties.SetName(placement, UiText.Get("工作表标签位置"));
        placement.SelectionChanged += (_, _) => { if (placement.SelectedIndex >= 0) Save(_settings.Value with { Sheets = (SheetNavigationPlacement)placement.SelectedIndex }); };
        _page.Children.Add(placement);
    }
}
