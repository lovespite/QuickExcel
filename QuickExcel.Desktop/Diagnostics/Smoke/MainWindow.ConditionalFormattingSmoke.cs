using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunConditionalFormattingSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool value, string message) { if (!value) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        async Task Rules(params ConditionalFormatRule[] rules)
        {
            _smokeConditionalFormats = rules; ConditionalFormatting_Click(this, new()); await _workspaceInput; _smokeConditionalFormats = null; await Idle();
        }
        StyleDefinition Style(int row, int col) => ConditionalPresentation.GetStyle(_document.Snapshot, _document.Snapshot.Sheets[0], row, col);
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "条件格式验收");
            ApplyValues([(0, 0, "Amount"), (1, 0, "1"), (2, 0, "5"), (3, 0, "10"), (0, 1, "Dependent"), (1, 1, "=A2*2"),
                (0, 2, "Text"), (1, 2, "'alpha"), (2, 2, "'ALPHA"), (3, 2, "'beta")]);
            var id = _document.Snapshot.Sheets[0].Id;
            var dialog = new ConditionalRuleDialog(this, new(1, 0, 3, 0));
            _ = Dispatcher.BeginInvoke(new Action(() =>
            {
                dialog.OperatorPicker.SelectedIndex = (int)ConditionalComparison.GreaterThan; dialog.First.Text = "3";
                dialog.NumberFormat.Text = "0.00\" units\""; dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }), DispatcherPriority.ApplicationIdle);
            Check(dialog.ShowDialog() == true && dialog.Result?.Formula1 == "=3", "Real rule editor produces a typed comparison and differential number format");
            var rule = dialog.Result!; await Rules(rule);
            Check(_document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 1 && Style(2, 0).FillColor == "FFFFC7CE", "Ribbon handler installs and paints the comparison rule");
            Check(_grid.DisplayText(2, 0) == "5.00 units" && _grid.DisplayText(1, 0) == "1", "Grid text uses the matching conditional number format only");
            var lower = new ConditionalFormatRule(rule.Ranges, ConditionalFormatKind.Expression, 2, new(Italic: true, FontColor: "FF0000FF"), "=TRUE");
            await Rules(rule, lower); Check(Style(2, 0).Italic && Style(2, 0).FontColor == "FF9C0006", "Priority resolves overlapping differential properties");
            var stopped = new ConditionalFormatRule(rule.Ranges, rule.Kind, 1, rule.Style, rule.Formula1, comparison: rule.Comparison, stopIfTrue: true);
            await Rules(stopped, lower); Check(!Style(2, 0).Italic && Style(1, 0).Italic, "Stop-if-true suppresses lower rules only on matching cells");
            await UndoAsync(false); await Idle(); Check(Style(2, 0).Italic, "Undo restores the prior rule collection");
            var dependent = new ConditionalFormatRule([new(1, 1, 1, 1)], ConditionalFormatKind.Expression, 3, new(FillColor: "FFC6EFCE"), "=$B$2>5");
            var duplicate = new ConditionalFormatRule([new(1, 2, 3, 2)], ConditionalFormatKind.DuplicateValues, 4, new(FillColor: "FFFFFF00"));
            await Rules(rule, lower, dependent, duplicate);
            Check(Style(1, 2).FillColor == "FFFFFF00" && Style(3, 2).FillColor is null, "Duplicate rules compare text without case sensitivity");
            _editor.SetCell(0, 2, 2, "unique"); await Idle(); Check(Style(1, 2).FillColor is null, "Editing one duplicate updates the other matching cell");
            _editor.SetCell(0, 1, 0, "4"); await RecalculateAsync(); await Idle();
            Check(Style(1, 1).FillColor == "FFC6EFCE" && _grid.DisplayText(1, 0) == "4.00 units", "Dependency edits and recalculation invalidate conditional presentation");
            var manager = new ConditionalFormatDialog(this, _document.Snapshot.Sheets[0], new(1, 0, 3, 0));
            _ = Dispatcher.BeginInvoke(new Action(() => manager.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent))), DispatcherPriority.ApplicationIdle);
            Check(manager.ShowDialog() == true && manager.Result.Select(r => r.Priority).SequenceEqual([1, 2, 3, 4]), "Real manager preserves ordered rules on apply");
            await Rules(); Check(Style(1, 0).FillColor is null, "Clearing the collection immediately restores base styles");
            var baseFit = _grid.MeasureAutoFit(SheetAxis.Column, 0);
            await UndoAsync(false); await Idle(); Check(_document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 4, "One undo restores all removed rules");
            var fit = _grid.MeasureAutoFit(SheetAxis.Column, 0);
            Check(fit > baseFit, "Auto-fit measures conditional text and font properties");
            _editor.Execute(new SetAxisMetadataCommand(id, SheetAxis.Column, 0, 1, (Math.Max(180, fit) - 5) / 7));
            var output = Path.Combine(folder, "conditional-ui.xlsx"); new TypedWorkbookStorage().Save(output, _document.Snapshot);
            await OpenPathAsync(output); await RecalculateAsync(); await Idle();
            Check(_document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 4 && _grid.DisplayText(1, 0) == "4.00 units", "Normal save and reopen retain rule behavior and formatted text");
            RibbonTabs.SelectedIndex = 2; await Idle(); CaptureGridEvidence(Path.Combine(folder, "conditional-formatting.png"));
            var artwork = _grid.Artwork; var dpi = VisualTreeHelper.GetDpi(artwork);
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(artwork.ActualWidth * dpi.DpiScaleX), (int)Math.Ceiling(artwork.ActualHeight * dpi.DpiScaleY),
                dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32); bitmap.Render(artwork);
            var bounds = _grid.Control.GetModelCellBounds(1, 0)!.Value;
            var crop = new CroppedBitmap(bitmap, new Int32Rect((int)(bounds.X * dpi.DpiScaleX), (int)(bounds.Y * dpi.DpiScaleY), (int)(bounds.Width * dpi.DpiScaleX), (int)(bounds.Height * dpi.DpiScaleY)));
            var pixels = new byte[crop.PixelWidth * crop.PixelHeight * 4]; crop.CopyPixels(pixels, crop.PixelWidth * 4, 0);
            var dark = Enumerable.Range(0, pixels.Length / 4).Count(i => pixels[i * 4 + 3] > 200 && pixels[i * 4 + 1] < 80 && pixels[i * 4 + 2] is > 100 and < 210);
            Check(dark > 5, "Conditional glyphs remain visible after resize and reopening");
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); await Idle();
                Check(_document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 8, "Native Excel comparisons enter the modeled rule collection");
                Check(ConditionalPresentation.GetStyle(_document.Snapshot, _document.Snapshot.Sheets[0], 5, 0).Bold, "Native differential font reaches WPF presentation");
            }
            File.WriteAllLines(Path.Combine(folder, "conditional-formatting-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _smokeConditionalFormats = null; _suppressAutoCalculation = false; }
    }
}
