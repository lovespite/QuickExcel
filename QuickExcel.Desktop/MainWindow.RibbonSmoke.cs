using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunRibbonSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var cases = new List<string>();
        void Check(bool condition, string message)
        {
            if (!condition) throw new Exception("Ribbon: " + message);
            cases.Add(message);
        }
        async Task Idle() { await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle); UpdateLayout(); }
        void Click(ButtonBase button) => button.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        _suppressAutoCalculation = true;
        try
        {
            InstallProductFixture();
            Width = 1360; Height = 860;
            await Idle();
            Check(HomeRibbonTab.IsSelected && RibbonPasteButton.IsVisible, "Home ribbon opens with the workbook");
            Check(HomeRibbonContent.ActualWidth <= CommandBar.ActualWidth, "All Home groups fit the default window");
            var revision = _document.Snapshot.Revision;
            foreach (TabItem tab in RibbonTabs.Items)
            {
                tab.IsSelected = true; await Idle();
                Check(((FrameworkElement)tab.Content).IsVisible, $"Tab {tab.Header} exposes its commands");
            }
            Check(_document.Snapshot.Revision == revision, "Switching ribbon tabs never changes the document");
            HomeRibbonTab.IsSelected = true; await Idle();
            GridSurface.Select(new(3, 1)); await Idle();
            GridSurface.SelectRange(new(3, 1, 4, 1), new(3, 1));
            Click(RibbonBold); await Idle();
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(3, 1)].Bold &&
                _document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(4, 1)].Bold && RibbonBold.IsChecked == true,
                "Bold ribbon button formats the selected range and reflects document state");
            await UndoAsync(false); await Idle();
            Check(RibbonBold.IsChecked == false, "Undo refreshes the ribbon state");
            await UndoAsync(true); await Idle();
            Check(RibbonBold.IsChecked == true, "Redo refreshes the ribbon state");
            GridSurface.Select(new(3, 1)); await Idle();
            FormulaBar.Focus(); FormulaBar.Text = "123";
            RibbonFontSize.SelectedItem = 18d; await Idle();
            Check(_document.GetCellInput(0, 3, 1) == "123" &&
                _document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(3, 1)].FontSize == 18,
                "Changing font size commits the original-cell draft and keeps the requested size");
            RibbonFontFamily.SelectedValue = "Consolas"; await Idle();
            Click(RibbonWrap); Click(RibbonAlignMiddle); Click(RibbonAlignCenter); await Idle();
            var style = _document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(3, 1)];
            Check(style.FontName == "Consolas" && style.WrapText && style.Vertical == 1 && style.Horizontal == 2,
                "Font, wrap and alignment ribbon controls execute document commands");
            var column = SheetGrid.Columns[1];
            var text = column.GetCellContent(_rows[3]);
            Check(text is TextBlock { FontSize: 24, TextWrapping: TextWrapping.Wrap, VerticalAlignment: VerticalAlignment.Center } block &&
                block.FontFamily.Source == "Consolas" || text is TextBox { FontSize: 24, TextWrapping: TextWrapping.Wrap } editor &&
                editor.FontFamily.Source == "Consolas" && editor.VerticalContentAlignment == VerticalAlignment.Center,
                "Grid renders the requested family, point size, wrapping and vertical alignment");
            RibbonNumberFormat.SelectedIndex = 2; await Idle();
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(3, 1)].NumberFormatId == 10,
                "Number format selector applies the chosen format");
            GridSurface.Select(new(4, 2)); await Idle();
            Check(RibbonBold.IsChecked == false && RibbonWrap.IsChecked == false && RibbonFontSize.SelectedItem is 11d,
                "Selection change restores the new cell's formatting state");
            var formatMenu = ProductChildren<Button>(HomeRibbonContent).Single(b => b.Content as string == "格式 ▾");
            Click(formatMenu); await Idle();
            Check(formatMenu.ContextMenu.IsOpen, "Ribbon dropdown opens from its button");
            formatMenu.ContextMenu.IsOpen = false;
            ScriptPage_Click(this, new()); await Idle();
            Check(ScriptPanel.IsVisible && !RibbonTabs.IsEnabled, "Workbook ribbon is disabled while editing scripts");
            WorkbookPage_Click(this, new()); await Idle();
            Check(RibbonTabs.IsEnabled, "Workbook ribbon reactivates when returning to the sheet");
            Check(ProductChildren<ButtonBase>(HomeRibbonContent).Where(b => b.ReadLocalValue(Ribbon.IconProperty) != DependencyProperty.UnsetValue)
                .All(b => !string.IsNullOrWhiteSpace(AutomationProperties.GetName(b))),
                "Every ribbon action has an accessible name");
            InstallProductFixture(); await Idle();
            GridSurface.Select(new(3, 1)); await Idle();
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-ribbon.png"));
            Width = 900; await Idle();
            var scroll = (ScrollViewer)RibbonTabs.Template.FindName("RibbonScroller", RibbonTabs);
            Check(scroll.ScrollableWidth > 0 && scroll.ComputedHorizontalScrollBarVisibility == Visibility.Visible,
                "Narrow window exposes horizontal scrolling without wrapping groups");
            scroll.ScrollToRightEnd(); await Idle();
            var last = ProductChildren<Button>(HomeRibbonContent).Last();
            var bounds = last.TransformToAncestor(scroll).TransformBounds(new Rect(last.RenderSize));
            Check(bounds.Left >= 0 && bounds.Right <= scroll.ActualWidth, "Last ribbon command remains reachable at minimum width");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-ribbon-narrow.png"));
            File.WriteAllText(Path.Combine(folder, "ribbon-cases.json"), JsonSerializer.Serialize(cases, new JsonSerializerOptions { WriteIndented = true }));
            File.WriteAllText(Path.Combine(folder, "ribbon-smoke-result.txt"), $"PASS: {cases.Count} WPF ribbon assertions; layout, tabs, formatting, undo/redo, draft ownership and accessibility metadata.");
        }
        finally { _suppressAutoCalculation = false; }
    }
}
