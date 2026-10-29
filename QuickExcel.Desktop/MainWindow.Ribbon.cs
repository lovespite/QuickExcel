using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private bool _updatingRibbon;
    private readonly List<string> _ribbonFonts = Fonts.SystemFontFamilies.Select(f => f.Source)
        .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase).ToList();
    private readonly List<double> _ribbonSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

    private void UpdateRibbonState()
    {
        if (_document is null || RibbonFontFamily is null) return;
        _updatingRibbon = true;
        try
        {
            var snapshot = _document.Snapshot;
            var sheet = snapshot.Sheets[_sheetIndex];
            var style = snapshot.Styles[sheet.GetStyleId(_selectedRow, _selectedColumn)];
            var font = style.FontName ?? "Microsoft YaHei UI";
            var size = style.FontSize ?? 11;
            if (!_ribbonFonts.Contains(font)) { _ribbonFonts.Add(font); RibbonFontFamily.ItemsSource = null; }
            if (!_ribbonSizes.Contains(size)) { _ribbonSizes.Add(size); _ribbonSizes.Sort(); RibbonFontSize.ItemsSource = null; }
            RibbonFontFamily.ItemsSource = _ribbonFonts;
            RibbonFontSize.ItemsSource = _ribbonSizes;
            RibbonFontFamily.SelectedItem = font;
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
            if (ReferenceEquals(sender, RibbonFontFamily) && selected is string family)
                ApplyRibbonStyle(new(FontName: family));
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
            if (Clipboard.ContainsText())
                await (_gridInput = PasteSelectionAsync(Clipboard.GetText(), (sender as FrameworkElement)?.Tag as string == "formulas"));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void RibbonClear_Click(object sender, RoutedEventArgs e)
    {
        try { await (_gridInput = HandleGridKeyAsync(Key.Delete, ModifierKeys.None, false)); }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
