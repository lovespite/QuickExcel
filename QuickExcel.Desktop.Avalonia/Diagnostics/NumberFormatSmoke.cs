using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunNumberFormatSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static void Click(Control root, string name) => Named<Button>(root, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Idle()
        {
            for (var i = 0; i < 100; i++)
            {
                await _operation; await Layout();
                if (!_working && !_workspace.IsBusy) { if (!_calculationPending) return; await PerformAsync(CalculatePendingAsync); }
                await Task.Delay(10);
            }
            throw new TimeoutException("Number format checks did not become idle.");
        }
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowBook(); await Layout();
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("1.25\n-2.5\n=1+2", new(0, 0, 0, 0), false)); Synchronize(); await FlushEditsAsync();
            await ExecuteAsync(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 2, 0), CellContent.FromFormula("=1+2"))]));
            await Idle();
            _grid.SelectRange(new(0, 0, 2, 0), new(0, 0));
            var dropdown = Named<DropDownButton>(this, "NumberFormatDropdown"); var flyout = (MenuFlyout)dropdown.Flyout!;
            flyout.ShowAt(dropdown); await Layout();
            var presets = flyout.Items.OfType<MenuItem>().Where(i => i.Header is StackPanel).ToArray();
            Check(presets.Length == 11, "Number dropdown contains all eleven quick formats");
            Check(presets.All(p => TopLevel.GetTopLevel(p) is not null), "Dropdown entries are attached to the visible popup");
            Check(presets.All(i => i.Header is StackPanel p && p.Children.OfType<TextBlock>().Count() == 2), "Every format shows a name and live preview");
            if (TopLevel.GetTopLevel(presets[0]) is { } popup) await Capture(popup, "number-format-menu");
            var fraction = presets.Single(i => AutomationProperties.GetName(i) == "分数"); flyout.Hide();
            fraction.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await Idle();
            Check(CurrentStyle.NumberFormatCode == "# ?/?" && _numberFormatLabel.Text == "分数", "Dropdown applies fraction format and updates current label");
            Check(Document.Snapshot.Sheets[0].GetCell(0, 0)!.Value == CellValue.FromNumber(1.25) && Document.Snapshot.Sheets[0].GetCell(2, 0)!.Content.FormulaText == "=1+2", "Formatting preserves values and formulas");
            await History(false); await Idle(); Check(CurrentStyle.NumberFormatCode is null, "Formatting a range undoes as one operation");
            await History(true); await Idle(); Check(CurrentStyle.NumberFormatCode == "# ?/?", "Formatting a range redoes as one operation");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                await Layout(); Check(Named<ListBox>(window, "FormatCategories").SelectedIndex == (int)NumberFormatCategory.Fraction, "Dialog initializes from current format");
                foreach (var category in Enum.GetValues<NumberFormatCategory>())
                {
                    Named<ListBox>(window, "FormatCategories").SelectedIndex = (int)category; await Layout();
                    Check(Named<Button>(window, "FormatConfirm").IsEnabled, "Category has a supported default: " + category);
                }
                var code = Named<TextBox>(window, "FormatCode"); code.Text = "[DBNum1]0"; await Layout();
                Check(!Named<Button>(window, "FormatConfirm").IsEnabled && !string.IsNullOrEmpty(Named<TextBlock>(window, "FormatError").Text), "Unsupported custom code reports inline validation");
                code.Text = "#,##0.000;[红色](#,##0.000)"; await Layout();
                Check(Named<TextBlock>(window, "FormatPreview").Text == "1.250", "Custom format updates the sample");
                Check(Document.Snapshot.Revision == revision, "Category and code previews never mutate the document");
                await Capture(window, "number-format-custom"); Click(window, "FormatCancel");
            };
            await NumberFormatAsync(); Check(Document.Snapshot.Revision == revision, "Cancel leaves document and history unchanged");
            _dialogDriver = async window =>
            {
                await Layout(); Named<ListBox>(window, "FormatCategories").SelectedIndex = (int)NumberFormatCategory.Currency; await Layout();
                Named<NumericUpDown>(window, "FormatDecimals").Value = 3; Named<ComboBox>(window, "FormatSymbol").SelectedItem = "€"; Named<ComboBox>(window, "FormatNegative").SelectedIndex = 3; await Layout();
                Check(Named<TextBlock>(window, "FormatPreview").Text == "€1.250", "Currency settings update live sample");
                await Capture(window, "number-format-currency"); Click(window, "FormatConfirm");
            };
            await NumberFormatAsync(); await Idle();
            Check(CurrentStyle.NumberFormatCode == "\"€\"#,##0.000;[Red](\"€\"#,##0.000)", "Currency settings commit through the style command");
            _dialogDriver = async window =>
            {
                await Layout();
                Check(Named<ListBox>(window, "FormatCategories").SelectedIndex == (int)NumberFormatCategory.Currency && Named<NumericUpDown>(window, "FormatDecimals").Value == 3 && Equals(Named<ComboBox>(window, "FormatSymbol").SelectedItem, "€") && Named<ComboBox>(window, "FormatNegative").SelectedIndex == 3,
                    "Reopening restores category, decimals, symbol and negative style");
                Click(window, "FormatConfirm");
            };
            revision = Document.Snapshot.Revision; await CellFormatAsync(); Check(Document.Snapshot.Revision == revision, "Confirming unchanged cell format does not edit mixed styles");
            _dialogDriver = async window =>
            {
                await Layout(); Named<ListBox>(window, "FormatCategories").SelectedIndex = (int)NumberFormatCategory.Custom; await Layout();
                Named<TextBox>(window, "FormatCode").Text = "0.00;[红色](0.00)"; await Layout(); Click(window, "FormatConfirm");
            };
            await NumberFormatAsync(); await Idle(); Check(CurrentStyle.NumberFormatCode == "0.00;[Red](0.00)", "Localized custom code is stored in portable XLSX form");
            _grid.Select(new(3, 0)); await ApplyStyleAsync(new(NumberFormatId: 49));
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("123", new(3, 0, 3, 0), false)); Synchronize();
            Check(Document.Snapshot.Sheets[0].GetCell(3, 0)!.Value == CellValue.FromText("123"), "New entries in text-formatted cells remain text");
            _grid.Select(new(0, 0)); await Capture(this, "number-format-workspace");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
