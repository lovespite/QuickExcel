using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private QuickExcel.Core.WorkbookDocument CreateProductWorkbook()
    {
        using var blank = QuickExcel.Core.WorkbookDocument.CreateBlank(_services);
        var modules = blank.Snapshot.ScriptModules;
        return new(blank.Snapshot.WithScriptModules(new QuickExcel.Domain.ScriptModuleSet(modules.Modules, modules.EntryModules, scriptApiVersion: 2)), _services);
    }
    private DesktopSettings _settings = new(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel", "settings.json"));
    private bool _hasWorkbook;
    private sealed record RecentEntry(string Path) { public string Name => System.IO.Path.GetFileName(Path); }

    private void ShowStartPage()
    {
        StartPage.Visibility = Visibility.Visible;
        CommandBar.Visibility = Workspace.Visibility = Visibility.Collapsed;
        ResumeButton.Visibility = _hasWorkbook ? Visibility.Visible : Visibility.Collapsed;
        RecentList.ItemsSource = _recentFiles.Paths.Select(path => new RecentEntry(path)).ToArray();
        NoRecentText.Visibility = _recentFiles.Paths.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        FileTitle.Text = ""; Title = "QuickExcel"; SelectionText.Text = ""; StatusText.Text = "就绪";
        StartPage.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
    }

    private void ShowWorkbookPage(bool script = false)
    {
        StartPage.Visibility = Visibility.Collapsed;
        CommandBar.Visibility = Workspace.Visibility = Visibility.Visible;
        WorkbookPage.Visibility = script ? Visibility.Collapsed : Visibility.Visible;
        ScriptPanel.Visibility = script ? Visibility.Visible : Visibility.Collapsed;
        RibbonTabs.IsEnabled = !script;
        WorkbookPageButton.FontWeight = script ? FontWeights.Normal : FontWeights.SemiBold;
        ScriptPageButton.FontWeight = script ? FontWeights.SemiBold : FontWeights.Normal;
        if (IsLoaded) UpdateLayout();
        if (_document is not null) UpdateChrome();
    }

    private Task _homeNavigation = Task.CompletedTask;
    private async void Home_Click(object sender, RoutedEventArgs e) => await (_homeNavigation = NavigateHomeAsync());
    private async Task NavigateHomeAsync()
    {
        // Leaving for the start page must not bypass the save prompt: the workbook
        // stays in memory, but an unprompted exit hides unsaved edits behind
        // "返回工作簿" and invites a later discard.
        if (_busy || !await ConfirmDiscardAsync()) return;
        ShowStartPage();
    }
    private void WorkbookPage_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        ShowWorkbookPage();
        GridSurface.Select(new(_selectedRow, _selectedColumn));
    }
    private void ScriptPage_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        ShowWorkbookPage(true);
        ScriptEditor.Focus();
    }
    private async void Recent_Activate(object sender, MouseButtonEventArgs e) => await OpenRecentAsync();
    private async void Recent_KeyDown(object sender, KeyEventArgs e)
    { if (e.Key == Key.Enter) { e.Handled = true; await OpenRecentAsync(); } }
    private async Task OpenRecentAsync()
    {
        if (!_busy && RecentList.SelectedItem is RecentEntry file && await ConfirmDiscardAsync()) await OpenPathAsync(file.Path);
    }

    private void ApplyPreferences(DesktopPreferences preferences)
    {
        var vertical = preferences.Sheets != SheetNavigationPlacement.Bottom;
        DockPanel.SetDock(SheetNavigation, preferences.Sheets switch { SheetNavigationPlacement.Left => Dock.Left, SheetNavigationPlacement.Right => Dock.Right, _ => Dock.Bottom });
        SheetNavigation.Width = vertical ? 190 : double.NaN;
        DockPanel.SetDock(AddSheetButton, vertical ? Dock.Bottom : Dock.Right);
        ScrollViewer.SetHorizontalScrollBarVisibility(SheetPicker, vertical ? ScrollBarVisibility.Disabled : ScrollBarVisibility.Auto);
        ScrollViewer.SetVerticalScrollBarVisibility(SheetPicker, vertical ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled);
        var panel = new FrameworkElementFactory(typeof(StackPanel));
        panel.SetValue(StackPanel.OrientationProperty, vertical ? Orientation.Vertical : Orientation.Horizontal);
        SheetPicker.ItemsPanel = new ItemsPanelTemplate(panel);
        ApplyWorksheetZoom();
        if (SheetPicker.SelectedItem is { } selected) SheetPicker.ScrollIntoView(selected);
    }

    private Window ToolDialog(string title, UIElement content, double width = 440, double height = 320) => new()
    {
        Owner = this, Title = title, Content = content, Width = width, Height = height,
        MinWidth = 350, MinHeight = 240, WindowStartupLocation = WindowStartupLocation.CenterOwner,
        Background = Brushes.White, FontFamily = FontFamily, FontSize = 13, ShowInTaskbar = false
    };

    private void Settings_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var panel = new StackPanel { Margin = new(24) };
        panel.Children.Add(new TextBlock { Text = "工作表标签位置", Margin = new(0, 0, 0, 8) });
        var placement = new ComboBox { ItemsSource = new[] { "底部", "左侧", "右侧" }, SelectedIndex = (int)_settings.Value.Sheets };
        AutomationProperties.SetName(placement, "工作表标签位置"); panel.Children.Add(placement);
        panel.Children.Add(new TextBlock { Text = "当前工作表缩放（随工作簿保存，可撤销）", Margin = new(0, 20, 0, 8) });
        var currentZoom = _document.Snapshot.Sheets[_sheetIndex].View.Zoom;
        var zoom = new ComboBox { ItemsSource = new[] { 10, 25, 50, 75, 100, 125, 150, 175, 200, 300, 400, currentZoom }.Distinct().Order().ToArray(), SelectedItem = currentZoom };
        AutomationProperties.SetName(zoom, "网格缩放百分比"); panel.Children.Add(zoom);
        var save = new Button { Content = "应用", IsDefault = true, Margin = new(0, 24, 0, 0), HorizontalAlignment = HorizontalAlignment.Right };
        panel.Children.Add(save);
        var dialog = ToolDialog("设置", panel);
        save.Click += (_, _) =>
        {
            var value = new DesktopPreferences((SheetNavigationPlacement)placement.SelectedIndex, _settings.Value.Zoom);
            SetWorksheetZoom((int)(zoom.SelectedItem ?? 100));
            _settings.Save(value); ApplyPreferences(value); dialog.Close();
            if (_settings.Error is { } error) ReportEditFailure(new IOException("设置未保存：" + error));
        };
        dialog.ShowDialog();
    }

    private void GoTo_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var address = PromptText("定位", "单元格地址", CellAddress.Text);
        if (address is null) return;
        try
        {
            var match = System.Text.RegularExpressions.Regex.Match(address.Trim(), "^([A-Za-z]{1,3})([1-9][0-9]{0,6})$");
            if (!match.Success) throw new ArgumentException("请输入有效地址，如 A1。");
            var column = 0;
            foreach (var c in match.Groups[1].Value.ToUpperInvariant()) column = checked(column * 26 + c - 'A' + 1);
            NavigateToCell(_sheetIndex, int.Parse(match.Groups[2].Value) - 1, column - 1);
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private void NavigateToCell(int sheet, int row, int column)
    {
        if (row < 0 || row >= QuickExcel.Domain.WorkbookLimits.MaxRows || column < 0 || column >= 256)
            throw new ArgumentException("当前界面支持 A–IV 列及 1–100000 行。");
        var target = _document.Snapshot.Sheets[sheet];
        if (target.Visibility != QuickExcel.Domain.SheetVisibility.Visible) throw new InvalidOperationException("目标工作表被隐藏，请先使用工作表显示操作。");
        ShowWorkbookPage();
        SheetPicker.SelectedIndex = Array.IndexOf(_visibleSheetIds, target.Id);
        if (target.Rows.GetValueOrDefault(row)?.Hidden == true || target.Columns.GetValueOrDefault(column)?.Hidden == true || CurrentView.IsFilteredOut(_document.Snapshot, sheet, row))
            throw new InvalidOperationException($"{target.Name}!{ColumnName(column)}{row + 1} 被隐藏，请先显示该行列或清除筛选。");
        _rowCount = Math.Max(_rowCount, row + 1); _columnCount = Math.Max(_columnCount, column + 1);
        RefreshSheet(true); GridSurface.Select(new(row, column));
    }

    private async void Recovery_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var dialog = new OpenFileDialog { Filter = "QuickExcel 副本 (*.xlsq)|*.xlsq", Title = "打开恢复副本" };
        if (dialog.ShowDialog(this) != true || !await ConfirmDiscardAsync()) return;
        var prior = _document.Session.SessionId;
        await OpenPathAsync(dialog.FileName, allowAutomaticScripts: false);
        if (_document.Session.SessionId == prior) return;
        // Recovery always opens an unbound copy. It cannot overwrite its source or carry authorization.
        _filePath = _versionPath = null; _fileVersion = null;
        _displayName = Path.GetFileNameWithoutExtension(dialog.FileName) + "-恢复.xlsq";
        _document.DisableScripts(); UpdateChrome();
        StatusText.Text = "已打开恢复副本，请另存。";
    }

    private Action<Window>? _smokeCompatibilityDriver;
    private async void Compatibility_Click(object sender, RoutedEventArgs e) => await ShowCompatibilityAsync();
    private async Task ShowCompatibilityAsync()
    {
        if (_busy || !CommitEditors()) return;
        var snapshot = _document.Snapshot;
        QuickExcel.Domain.WorkbookCompatibilityReport? report = null;
        if (!await WorkAsync("检查兼容性…", token => report = QuickExcel.Storage.TypedWorkbookStorage.InspectCompatibility(snapshot, token), cancellable: true)) return;
        var notices = new List<string>();
        if (snapshot.FileState?.RequiresMigration == true) notices.Add("XLSQ v1 → v2：首次保存为新副本，原文件保留。");
        foreach (var sheet in snapshot.Sheets)
        {
            if (sheet.UsedColumnCount > 256) notices.Add($"{sheet.Name}：IV 列之后的数据会保存，当前网格不显示。");
            if (sheet.View.ActiveCell.Column >= 256 || sheet.View.FrozenRows > 10 || sheet.View.FrozenColumns >= 256)
                notices.Add($"{sheet.Name}：原活动格/冻结设置会保存，但超出当前网格的 10 冻结行或 A–IV 显示范围；显示采用受限视图。");
        }
        var dialog = new CompatibilityDialog(this, report!, saveMode: false, notices);
        if (_smokeCompatibilityDriver is not null) dialog.Loaded += (_, _) => _smokeCompatibilityDriver?.Invoke(dialog);
        dialog.ShowDialog();
    }

    private void ExportOptions_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var panel = new StackPanel { Margin = new(24) };
        var values = new Button { Content = "值快照 .xlsx", Margin = new(0, 0, 0, 6), Padding = new(16) };
        panel.Children.Add(values); panel.Children.Add(new TextBlock { Text = "保留当前结果，所有公式转为值。", Margin = new(0, 0, 0, 22) });
        var formulas = new Button { Content = "可移植公式 .xlsx", Margin = new(0, 0, 0, 6), Padding = new(16) };
        panel.Children.Add(formulas); panel.Children.Add(new TextBlock { Text = "保留支持的标准公式，脚本依赖转为值。", TextWrapping = TextWrapping.Wrap });
        var dialog = ToolDialog("导出", panel, 470, 300);
        values.Click += (_, _) => { dialog.Close(); Export_Click(this, new()); };
        formulas.Click += (_, _) => { dialog.Close(); ExportPortable_Click(this, new()); };
        dialog.ShowDialog();
    }
}
