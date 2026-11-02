using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunOutlineSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool value, string message) { if (!value) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        void Pick(SheetAxis axis, int first, int last, int level) => OutlineGroups.SelectedItem = OutlineGroups.Items.Cast<OutlineItem>().Single(i => i.Group.Axis == axis && i.Group.First == first && i.Group.Last == last && i.Group.Level == level);
        async Task Toggle(bool collapsed) { if (collapsed) CollapseOutline_Click(this, new()); else ExpandOutline_Click(this, new()); await _workspaceInput; await RecalculateAsync(); await Idle(); }
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "分类汇总验收");
            ApplyValues([(0, 0, "Group"), (0, 1, "Amount"), (0, 2, "Other"), (1, 0, "a"), (2, 0, "A"), (3, 0, "b"), (4, 0, "b"), (5, 0, "b"), (6, 0, "c"), (7, 0, "a"),
                (1, 1, "1"), (2, 1, "2"), (3, 1, "3"), (4, 1, "4"), (5, 1, "5"), (6, 1, "6"), (7, 1, "7"),
                (1, 2, "=B2*2"), (2, 2, "=B3*2"), (3, 2, "=B4*2"), (4, 2, "=B5*2"), (5, 2, "=B6*2"), (6, 2, "=B7*2"), (7, 2, "=B8*2"), (0, 4, "=B8"), (9, 0, "tail")]);
            await RecalculateAsync(); await Idle(); var book = _document.Snapshot; var id = book.Sheets[0].Id;
            var dialog = new SubtotalDialog(this, book, id, new(0, 0, 7, 2));
            _ = Dispatcher.BeginInvoke(new Action(async () =>
            {
                dialog.TotalColumns.Text = "2,3"; dialog.PreviewButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                Check(dialog.Preview?.Groups.Length == 4 && dialog.Preview.InsertedRows == 5, "Real dialog previews contiguous groups and five whole-row insertions");
                Check(_document.Snapshot.Revision == book.Revision, "Preview does not commit any edits");
                dialog.Function.SelectedIndex = 0; Check(dialog.Preview is null && !dialog.ApplyButton.IsEnabled, "Changing the function invalidates the candidate");
                dialog.Function.SelectedIndex = 8; dialog.PreviewButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                await Dispatcher.InvokeAsync(dialog.UpdateLayout, DispatcherPriority.ApplicationIdle);
                var dpi = VisualTreeHelper.GetDpi(dialog); var bitmap = new RenderTargetBitmap((int)(dialog.ActualWidth * dpi.DpiScaleX), (int)(dialog.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
                bitmap.Render(dialog); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using (var file = File.Create(Path.Combine(folder, "subtotal-preview.png"))) encoder.Save(file);
                dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }), DispatcherPriority.ApplicationIdle);
            Check(dialog.ShowDialog() == true, "Preview dialog applies the reviewed candidate"); _smokeSubtotalPreview = dialog.Preview;
            Subtotal_Click(this, new()); await _workspaceInput; _smokeSubtotalPreview = null;
            Check(StatusText.Text.StartsWith("已插入 5 行"), "Subtotal handler commits: " + StatusText.Text); await RecalculateAsync(); await Idle();
            Check(_document.GetCellValue(0, 12, 1) == "28" && _document.GetCellValue(0, 12, 2) == "56", "Grand totals exclude the inserted nested subtotal formulas");
            Check(_document.Snapshot.Sheets[0].GetCell(0, 4)!.Content.FormulaText == "=B11" && _document.GetCellValue(0, 14, 0) == "tail", "Whole-row insertion moves outside references and later data");
            await UndoAsync(false); Check(_document.GetCellValue(0, 9, 0) == "tail" && !WorksheetOutline.HasGroups(_document.Snapshot.Sheets[0]), "One undo removes all generated rows and outline metadata");
            await UndoAsync(true); await RecalculateAsync(); await Idle();
            _editor.SetCell(0, 0, 6, "=SUBTOTAL(109,B2:B11)"); await RecalculateAsync(); await Idle();
            Check(OutlinePanel.Visibility == Visibility.Visible && OutlineGroups.Items.Count == 5, "Imported/generated outline levels have a visible group selector");
            var collapseButton = OutlinePanel.Children.OfType<Button>().Single(b => (b.Content as string)?.Contains("折叠") == true);
            var expandButton = OutlinePanel.Children.OfType<Button>().Single(b => (b.Content as string)?.Contains("展开") == true);
            var panelHeight = OutlinePanel.ActualHeight; var collapseHeight = collapseButton.ActualHeight; var expandHeight = expandButton.ActualHeight;
            collapseButton.Focus(); await Idle();
            Check(Math.Abs(OutlinePanel.ActualHeight - panelHeight) < .1 && Math.Abs(collapseButton.ActualHeight - collapseHeight) < .1,
                "Focusing Collapse draws its focus ring without changing the outline panel height");
            expandButton.Focus(); await Idle();
            Check(Math.Abs(OutlinePanel.ActualHeight - panelHeight) < .1 && Math.Abs(expandButton.ActualHeight - expandHeight) < .1,
                "Focusing Expand keeps the outline panel height stable (no layout flicker)");
            Pick(SheetAxis.Row, 1, 2, 2); await Toggle(true);
            Check(!_grid.Control.GetVisibleModelRows().Contains(1) && _document.GetCellValue(0, 0, 6) == "25", "Collapsed detail rows disappear and SUBTOTAL 109 recalculates");
            Pick(SheetAxis.Row, 1, 11, 1); await Toggle(true);
            Check(_document.GetCellValue(0, 0, 6) == "0" && _document.GetCellValue(0, 12, 1) == "28", "Outer collapse distinguishes SUBTOTAL 9 and 109");
            await Toggle(false);
            Check(!_grid.Control.GetVisibleModelRows().Contains(1) && _grid.Control.GetVisibleModelRows().Contains(4), "Expanding the parent retains the inner collapse");
            GridSurface.SelectRange(new(0, 9, 0, 11), new(0, 9)); _smokeOutlineAxis = SheetAxis.Column; GroupOutline_Click(this, new()); await _workspaceInput; await RecalculateAsync(); await Idle();
            Pick(SheetAxis.Column, 9, 11, 1); await Toggle(true);
            Check(_document.Snapshot.Sheets[0].Columns[10].Hidden, "Column group handler collapses the selected columns");
            await Toggle(false); GridSurface.SelectRange(new(0, 9, 0, 11), new(0, 9)); UngroupOutline_Click(this, new()); await _workspaceInput; await RecalculateAsync(); await Idle();
            Check(!WorksheetOutline.Groups(_document.Snapshot.Sheets[0], SheetAxis.Column).Any(), "Ungroup handler removes one column level");
            _editor.Execute(new SetAxisMetadataCommand(id, SheetAxis.Column, 0, 7, 24)); await RecalculateAsync();
            var path = Path.Combine(folder, "outline-ui.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot); await OpenPathAsync(path); await RecalculateAsync(); await Idle();
            Check(_document.Snapshot.Sheets[0].Rows[1].Hidden && _document.GetCellValue(0, 0, 6) == "25", "Ordinary save/reopen preserves nested collapse and visibility calculation");
            RibbonTabs.SelectedIndex = 2; GridSurface.Select(new(0, 0)); await Idle(); CaptureGridEvidence(Path.Combine(folder, "outline-grid.png"));
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); await RecalculateAsync(); await Idle();
                Check(_document.Snapshot.Sheets.Any(s => s.Data.Outline is { SummaryBelow: false, SummaryRight: false }), "Native above/left outline settings import intact");
            }
            File.WriteAllLines(Path.Combine(folder, "outline-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _smokeSubtotalPreview = null; _smokeOutlineAxis = null; _suppressAutoCalculation = false; }
    }
}
