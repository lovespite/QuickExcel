using System.Runtime.Versioning;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Microsoft.Win32;
using Markdown.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckSettingsAsync(string directory, Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static void Click(Window w, string label) => w.GetVisualDescendants().OfType<Button>()
            .Single(b => Equals(b.Content, UiText.Get(label))).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        _dialogDriver = async window =>
        {
            await Layout();
            var english = UiText.Language == UiText.English;
            var nav = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsNavigation");
            check(nav.ItemCount == 6, "Settings exposes six separately navigable sections");
            nav.SelectedItem = UiText.Get("历史管理"); await Layout();
            var recent = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsRecentFiles");
            recent.SelectedIndex = 1; var before = recent.ItemCount; Click(window, "移除选中记录");
            check(recent.ItemCount == before - 1, "Remove updates only the selected history entry");
            Click(window, "清空历史记录"); await Layout();
            Click(window.OwnedWindows.Single(), "确定"); await Layout();
            check(recent.ItemCount == 0, "Confirmed history clear empties the settings list");
            nav.SelectedItem = UiText.Get("保存与计算"); await Layout();
            var manual = window.GetVisualDescendants().OfType<RadioButton>().Single(r => r.Content?.ToString()?.StartsWith(english ? "Manual" : "手动") == true);
            manual.IsChecked = true;
            check(!window.GetVisualDescendants().OfType<ToggleSwitch>().Single().IsEnabled, "Autosave is a disabled placeholder");
            await capture(window, "settings-calculation");
            nav.SelectedItem = UiText.Get("脚本与 API"); await Layout();
            var search = window.GetVisualDescendants().OfType<TextBox>().Single(t => t.Name == "SettingsApiSearch"); search.Text = "Range.getValues"; await Layout();
            var tree = window.GetVisualDescendants().OfType<TreeView>().Single(t => t.Name == "SettingsApiTree");
            var kind = window.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Name == "SettingsApiDetailKind");
            check(!kind.IsVisible, "The member kind stays hidden until a member is selected");
            check(tree.ItemCount == 1, "Manifest-backed tree filters to the matching object group");
            var group = (TreeViewItem)tree.Items[0]!;
            var type = (TreeViewItem)group.Items[0]!;
            var member = (TreeViewItem)type.Items[0]!;
            check(tree.ItemCount == 1 && group.Items.Count == 1 && type.Items.Count == 1, "Filtered tree keeps only the matching member");
            tree.SelectedItem = member; await Layout();
            var signature = window.GetVisualDescendants().OfType<TextBox>().Single(t => t.Name == "SettingsApiSignature");
            check(signature.Text == "getValues(): (string | number | boolean)[][]", "Selecting a member shows its declared signature");
            check(window.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Name == "SettingsApiDetailTitle").Text == "Range.getValues",
                "The details pane identifies the selected member");
            check(kind.IsEffectivelyVisible && kind.Text == UiText.Get("宿主调用"), "Selecting a host member reveals its call type");
            tree.SelectedItem = type; await Layout();
            check(!kind.IsVisible, "Selecting a type hides member details");
            foreach (var (query, label) in new[] { ("Worksheet.getRange", "本地成员"), ("BorderIndex.edgeTop", "枚举值"), ("Range.getValues", "宿主调用") })
            {
                search.Text = query; await Layout();
                var matchingGroup = (TreeViewItem)tree.Items[0]!;
                var matchingType = (TreeViewItem)matchingGroup.Items[0]!;
                tree.SelectedItem = matchingType.Items[0]; await Layout();
                check(kind.IsEffectivelyVisible && kind.Text == UiText.Get(label), "Changing selection reveals the correct kind: " + query);
            }
            await capture(window, "settings-scripts");
            nav.SelectedItem = UiText.Get("关于与许可"); await Layout();
            var icon = window.GetVisualDescendants().OfType<Image>().Single(i => i.Name == "AboutProductIcon");
            check(icon.Source is Bitmap && ReferenceEquals(icon.Source, Application.Current!.Resources[DesktopApp.ProductIconKey])
                && icon.Bounds.Size == new Size(32, 32), "About displays the application icon at 32 x 32");
            check(window.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text?.StartsWith("Copyright in the original code") == true),
                "Copyright is displayed in English in either UI language");
            await capture(window, "settings-about");
            var licenses = window.GetVisualDescendants().OfType<ComboBox>().Single();
            check(window.GetVisualDescendants().OfType<MarkdownScrollViewer>().Single().Markdown!.StartsWith("# Third-party components and licenses"),
                "The bundled third-party declaration is shown in English");
            licenses.SelectedItem = licenses.Items.OfType<string>().Single(p => p.EndsWith("OFFICE-SCRIPTS-NOTICE.md"));
            check(window.GetVisualDescendants().OfType<MarkdownScrollViewer>().Single().Markdown!.Contains("CC BY 4.0"), "Bundled license text and attribution can be read offline");
            nav.SelectedItem = english ? "Appearance" : "界面"; await Layout();
            var language = window.GetVisualDescendants().OfType<ComboBox>().Single(b => b.Name == "DisplayLanguage");
            check(language.SelectedItem?.ToString() == (english ? "English" : "简体中文"), "The language selector shows the persisted choice");
            await capture(window, "settings-language");
            language.SelectedIndex = english ? 0 : 1;
            Click(window, "关闭");
        };
        var item = ((MenuItem)_menu.Items[0]!).Items.OfType<MenuItem>().Single(i => Equals(i.Header, UiText.Get("设置")));
        item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation;
        check(_recent.Children.Single() is TextBlock && File.Exists(_workspace.FilePath), "Home history updates without deleting the workbook");
        var reloaded = new DesktopSettings(Path.Combine(directory, "settings", "settings.json")); reloaded.Load();
        check(reloaded.Value.Language == (UiText.Language == UiText.English ? UiText.Chinese : UiText.English),
            "Changing the language persists for the next launch");
        check(reloaded.Value.Calculation == DesktopCalculationMode.Manual, "Manual calculation selected in Settings survives reload");
        _workspace.New(); Document.SetCell(0, 0, 0, "2"); Document.SetCell(0, 0, 1, "=A1*3");
        await CalculateAsync(); Document.SetCell(0, 0, 0, "4"); await FlushEditsAsync();
        check(!_calculationTimer.IsEnabled && _calculationPending && Document.GetCellValue(0, 0, 1) != "12", "Manual mode retains pending edits through an ordinary flush");
        var source = Path.Combine(directory, "原始 数据.csv");
        var csv = "001,公式\r\n\"引号\"\"与,逗号\n换行\",=1+2\r\n";
        var dialogs = new SmokeFiles { NextSave = Path.Combine(directory, "manual.xlsq") }; _files = dialogs;
        check(await SaveAsync(false) && _services.Storage.Open(dialogs.NextSave).Sheets[0].GetCell(0, 1)!.Value.NumberValue == 12,
            "Manual save writes the recalculated result");
        await File.WriteAllTextAsync(source, csv); await OpenPathAsync(source);
        check(Document.GetCellValue(0, 0, 0) == "001" && Document.GetCellValue(0, 1, 0) == "引号\"与,逗号\n换行" &&
            Document.GetCellValue(0, 1, 1) == "=1+2", "CSV opens via normal file workflow with quoted multiline text, leading zeros and literal formulas");
        dialogs.NextSave = Path.Combine(directory, "imported.xlsq");
        check(await SaveAsync(false) && _workspace.FilePath == dialogs.NextSave && await File.ReadAllTextAsync(source) == csv,
            "Saving imported CSV creates a workbook and preserves the source");
        _settings.Save(_settings.Value with { Calculation = DesktopCalculationMode.Automatic });
        if (OperatingSystem.IsWindows()) CheckAssociationRegistration(check);
    }

    [SupportedOSPlatform("windows")]
    private static void CheckAssociationRegistration(Action<bool, string> check)
    {
        var name = @"Software\QuickExcel-Association-Test-" + Guid.NewGuid().ToString("N");
        try
        {
            using var root = Registry.CurrentUser.CreateSubKey(name);
            check(WindowsFileAssociations.ReadSelection(root).SequenceEqual([".xlsq"]), "Initial association selection defaults to XLSQ");
            var executable = Path.Combine(AppContext.BaseDirectory, "quickxl.exe");
            check(File.Exists(executable), "File associations target the packaged quickxl.exe launcher");
            WindowsFileAssociations.Register(root, executable, [".xlsq", ".csv"]);
            var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "QuickExcel-XLSQ.ico");
            using var icon = root.OpenSubKey(@"Software\Classes\QuickExcel.xlsq\DefaultIcon");
            check(File.Exists(iconPath) && Equals(icon?.GetValue(""), $"\"{iconPath}\",0"),
                "The XLSQ file type uses the packaged workbook icon");
            using var command = root.OpenSubKey(@"Software\Classes\QuickExcel.csv\shell\open\command");
            check(Equals(command?.GetValue(""), $"\"{executable}\" \"%1\""), "Registered CSV command quotes both the executable and selected file");
            WindowsFileAssociations.Register(root, executable, [".xlsx"]);
            using var xlsxIcon = root.OpenSubKey(@"Software\Classes\QuickExcel.xlsx\DefaultIcon");
            using var csvIcon = root.OpenSubKey(@"Software\Classes\QuickExcel.csv\DefaultIcon");
            check(xlsxIcon is null && csvIcon is null, "The XLSQ file icon is exclusive to the native workbook format");
            check(WindowsFileAssociations.ReadSelection(root).SequenceEqual([".xlsx"]), "Reapplying associations removes deselected advertisements");
        }
        finally { Registry.CurrentUser.DeleteSubKeyTree(name, throwOnMissingSubKey: false); }
    }
}
