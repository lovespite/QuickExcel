using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Markup;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private bool _updatingRibbon;
    private sealed record RibbonFont(string Source, string Name);
    private readonly List<RibbonFont> _ribbonFonts = Fonts.SystemFontFamilies.Select(f => new RibbonFont(f.Source,
        f.Source switch { "Microsoft YaHei UI" => "微软雅黑 UI", "Microsoft YaHei" => "微软雅黑",
            _ => f.FamilyNames.TryGetValue(XmlLanguage.GetLanguage("zh-CN"), out var name) ? name : f.Source }))
        .OrderBy(font => font.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
    private readonly List<double> _ribbonSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

    private void UpdateRibbonState()
    {
        if (_document is null || RibbonFontFamily is null) return;
        _updatingRibbon = true;
        try
        {
            var snapshot = _document.Snapshot;
            var readOnly = snapshot.FileState?.FormulaReadOnly == true;
            ScriptEditor.IsReadOnly = readOnly;
            FormulaBar.IsReadOnly = readOnly || snapshot.Sheets[_sheetIndex].GetCell(_selectedRow, _selectedColumn)?.Content.RichText is not null;
            RibbonPasteButton.IsEnabled = !readOnly;
            RibbonCutButton.IsEnabled = !readOnly;
            foreach (var group in HomeRibbonContent.Children.OfType<GroupBox>().Where(g => g.Header as string is "字体" or "对齐方式" or "数字" or "单元格"))
                group.IsEnabled = !readOnly;
            var sheet = snapshot.Sheets[_sheetIndex];
            var style = snapshot.Styles[sheet.GetStyleId(_selectedRow, _selectedColumn)];
            var font = style.FontName ?? "Microsoft YaHei UI";
            var size = style.FontSize ?? 11;
            if (!_ribbonFonts.Any(f => f.Source == font)) { _ribbonFonts.Add(new(font, font)); RibbonFontFamily.ItemsSource = null; }
            if (!_ribbonSizes.Contains(size)) { _ribbonSizes.Add(size); _ribbonSizes.Sort(); RibbonFontSize.ItemsSource = null; }
            RibbonFontFamily.ItemsSource = _ribbonFonts;
            RibbonFontSize.ItemsSource = _ribbonSizes;
            RibbonFontFamily.SelectedValue = font;
            RibbonFontSize.SelectedItem = size;
            RibbonBold.IsChecked = style.Bold;
            RibbonItalic.IsChecked = style.Italic;
            RibbonWrap.IsChecked = style.WrapText;
            RibbonAlignLeft.IsChecked = style.Horizontal is 0 or 1;
            RibbonAlignCenter.IsChecked = style.Horizontal == 2;
            RibbonAlignRight.IsChecked = style.Horizontal == 3;
            RibbonAlignTop.IsChecked = style.Vertical == 0;
            RibbonAlignMiddle.IsChecked = style.Vertical == 1;
            RibbonAlignBottom.IsChecked = style.Vertical == 2;
            var number = style.NumberFormatCode == "yyyy-mm-dd" ? "date" : style.NumberFormatCode is not null ? null :
                style.NumberFormatId switch { 0 => "general", 2 => "number", 10 => "percent", _ => null };
            RibbonNumberFormat.SelectedItem = RibbonNumberFormat.Items.OfType<ComboBoxItem>().FirstOrDefault(i => i.Tag as string == number);
            RibbonNumberFormat.ToolTip = number is null ? "当前格式：" + (style.NumberFormatCode ?? style.NumberFormatId.ToString(CultureInfo.InvariantCulture)) : "数字格式";
        }
        finally { _updatingRibbon = false; }
    }

    private void ApplyRibbonStyle(CellStylePatch patch)
    {
        var snapshot = _document.Snapshot;
        CurrentView.RequireCurrentFilter(snapshot);
        var sheet = snapshot.Sheets[_sheetIndex];
        var selection = GridSurface.Selection.Count == 0 ? [new CellPosition(_selectedRow, _selectedColumn)] : GridSurface.Selection;
        var addresses = selection.Where(p => !CurrentView.IsFilteredOut(snapshot, _sheetIndex, p.Row))
            .Select(p => new CellAddress(sheet.Id, p.Row, p.Column));
        ExecuteWorkbookCommand(new SetCellStyleCommand(addresses, patch));
    }

    private void RibbonFont_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingRibbon || _document is null) return;
        var selected = ((ComboBox)sender).SelectedItem;
        try
        {
            if (_busy || !CommitEditors()) return;
            if (ReferenceEquals(sender, RibbonFontFamily) && selected is RibbonFont family)
                ApplyRibbonStyle(new(FontName: family.Source));
            else if (ReferenceEquals(sender, RibbonFontSize) && selected is double size)
                ApplyRibbonStyle(new(FontSize: size));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { UpdateRibbonState(); }
    }

    private void RibbonNumberFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingRibbon || _document is null) return;
        if (RibbonNumberFormat.SelectedItem is ComboBoxItem item) Style_Click(item, new());
    }

    private void RibbonTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!ReferenceEquals(e.Source, RibbonTabs)) return;
        if (RibbonTabs.Template.FindName("RibbonScroller", RibbonTabs) is ScrollViewer scroller) scroller.ScrollToLeftEnd();
    }

    private void RibbonMenu_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition()) return;
        if (sender is Button { ContextMenu: { } menu } button)
        {
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
        }
    }

    private void RibbonCopy_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition()) return;
        try { CopySelection(); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void RibbonPaste_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition()) return;
        try
        {
            await (_gridInput = PasteClipboardAsync((sender as FrameworkElement)?.Tag as string ?? "all"));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void RibbonClear_Click(object sender, RoutedEventArgs e)
    {
        try { await (_gridInput = HandleGridKeyAsync(Key.Delete, ModifierKeys.None, false)); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
