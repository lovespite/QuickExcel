using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFormatEfficiencySmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool ok, string message) { if (!ok) throw new Exception(message); checks.Add(message); }
        async Task Stats() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await _selectionStatisticsTask; }
        CellSnapshot Literal(CellValue value) => new(CellContent.FromLiteral(value));
        var cells = new Dictionary<CellPosition, CellSnapshot> {
            [new(0, 0)] = Literal(CellValue.FromNumber(123.456)), [new(1, 0)] = Literal(CellValue.FromNumber(10)), [new(2, 0)] = Literal(CellValue.FromNumber(20)),
            [new(3, 0)] = Literal(CellValue.FromNumber(30)), [new(4, 0)] = Literal(CellValue.FromError(CellErrorCode.NotAvailable)),
            [new(5, 0)] = Literal(CellValue.FromText("")), [new(6, 0)] = Literal(CellValue.FromBoolean(true)), [new(0, 1)] = new(CellContent.FromFormula("=A1*2")),
            [new(0, 2)] = Literal(CellValue.FromText("00123")) };
        for (var row = 1; row <= 6; row++) cells[new(row, 1)] = Literal(CellValue.FromText(row == 3 ? "no" : "yes"));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Efficiency", cells, rows: [new(2, new(Hidden: true))]);
        try
        {
            InstallDocument(new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]), _services), null, "格式效率验收.xlsx");
            await RecalculateAsync();
            GridSurface.SelectRange(new(0, 0, 6, 0), new(0, 0)); await Stats();
            Check(SelectionText.Text.Contains("163.456") && SelectionText.Text.Contains("计数：6"), "Selection statistics exclude manually hidden rows and count errors/text/booleans");
            ApplyFilter(new(0, 0, 6, 1), 1, "yes"); GridSurface.SelectRange(new(0, 0, 6, 0), new(0, 0)); await Stats();
            Check(SelectionText.Text.Contains("133.456") && SelectionText.Text.Contains("66.728") && SelectionText.Text.Contains("计数：5"), "Filtered selection statistics use visible model coordinates");
            Check(SelectionText.ToolTip.ToString()!.Contains("错误 1") && SelectionText.ToolTip.ToString()!.Contains("不转换文本"), "Statistics explain excluded errors and numeric coercion policy");
            ClearFilter_Click(this, new()); GridSurface.Select(new(0, 0));
            DecimalPlaces_Click(new Button { Tag = "increase" }, new());
            var current = _document.Snapshot; var style = current.Styles[current.Sheets[0].GetStyleId(0, 0)];
            Check(style.NumberFormatCode == "0.0000" && current.Sheets[0].GetCell(0, 0)!.Value == CellValue.FromNumber(123.456), "Increase decimals changes only display precision through the real handler");
            DecimalPlaces_Click(new Button { Tag = "decrease" }, new());
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].NumberFormatCode == "0.000", "Decrease decimals removes one placeholder");
            await UndoAsync(false); Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].NumberFormatCode == "0.0000", "One undo restores display precision");
            await UndoAsync(true);
            Style_Click(new Button { Tag = "currency" }, new()); Style_Click(new Button { Tag = "bold" }, new());
            FormatBrush_Click(RibbonFormatBrush, new()); Check(_formatBrush is not null && RibbonFormatBrush.Content.ToString() == "应用格式", "Format brush captures the source style without editing");
            GridSurface.Select(new(0, 1)); FormatBrush_Click(RibbonFormatBrush, new());
            current = _document.Snapshot;
            Check(_formatBrush is null && current.Sheets[0].GetStyleId(0, 0) == current.Sheets[0].GetStyleId(0, 1), "Format brush applies captured style to the chosen target");
            Check(current.Sheets[0].GetCell(0, 1)!.Content.FormulaText == "=A1*2" && current.Sheets[0].GetCell(0, 1)!.Value == CellValue.FromNumber(246.912), "Format brush preserves formula and fresh cached value");
            await UndoAsync(false); Check(_document.Snapshot.Sheets[0].GetStyleId(0, 1) == 0, "Format brush is one undo transaction"); await UndoAsync(true);
            Fill_Click(new Button { Tag = "clear-formats" }, new()); await _gridInput;
            Check(_document.Snapshot.Sheets[0].GetStyleId(0, 1) == 0 && _document.Snapshot.Sheets[0].GetCell(0, 1)!.Content.FormulaText == "=A1*2", "Existing clear-formats command removes the brush style and retains the formula");
            await UndoAsync(false);
            GridSurface.Select(new(0, 0)); FormatBrush_Click(RibbonFormatBrush, new());
            Check(_formatBrush is not null, "Format brush can be sampled again after undo");
            var escape = new System.Windows.Input.KeyEventArgs(System.Windows.Input.Keyboard.PrimaryDevice, PresentationSource.FromVisual(this), 0, System.Windows.Input.Key.Escape)
                { RoutedEvent = System.Windows.Input.Keyboard.KeyDownEvent };
            Window_KeyDown(this, escape); Check(escape.Handled && _formatBrush is null, "Escape cancels a pending brush without applying it");
            GridSurface.Select(new(0, 2)); Style_Click(new Button { Tag = "thousands" }, new());
            Check(_document.Snapshot.Sheets[0].GetCell(0, 2)!.Value == CellValue.FromText("00123"), "Common numeric format does not convert a text ID");
            Style_Click(new Button { Tag = "scientific" }, new());
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 2)].NumberFormatId == 11, "Scientific format entry applies its built-in format");
            var colorMenu = BuildColorMenu("fill");
            Check(ProductChildren<Button>((StackPanel)colorMenu.Items[0]!).Count() >= 26, "Fill color dropdown exposes sample swatches plus clear and custom-color commands");
            var picker = new ColorPickerDialog(this, "颜色验收", "FF12AB34");
            Check(picker.SelectedColor == "FF000000", "Custom color dialog starts without writing until confirmed");
            GridSurface.Select(new(0, 0));
            ApplyRibbonColor("fill", "FFFF0000");
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].FillColor == "FFFF0000", "Fill color dropdown applies the chosen swatch");
            ApplyRibbonColor("font", "FF0070C0");
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].FontColor == "FF0070C0", "Font color dropdown applies the chosen swatch");
            ApplyRibbonColor("fill", null);
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].FillColor is null, "No-fill command clears the cell fill without touching the font color");
            ApplyRibbonColor("font", null);
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].FontColor is null, "Automatic color command restores the inherited font color");
            for (var i = 0; i < 4; i++) await UndoAsync(false);
            GridSurface.Select(new(0, 0)); FormatBrush_Click(RibbonFormatBrush, new());
            _document.SetCell(0, 9, 3, "lease changed"); RefreshSheet(); GridSurface.Select(new(1, 0)); var before = _document.Snapshot;
            FormatBrush_Click(RibbonFormatBrush, new()); Check(ReferenceEquals(before, _document.Snapshot) && _formatBrush is null, "An intervening edit expires the format brush without partial writes");
            GridSurface.Select(new(0, 1)); await Stats(); Check(SelectionText.Text.Contains("待重算"), "Statistics label old formula caches without triggering calculation");
            await RecalculateAsync(); GridSurface.SelectRange(new(0, 0, 1, 2), new(0, 0)); await Stats();
            var path = Path.Combine(folder, "format-efficiency.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            var saved = new TypedWorkbookStorage().Open(path);
            Check(saved.Sheets[0].GetCell(0, 1)!.Content.FormulaText == "=A1*2" && saved.Sheets[0].GetCell(0, 2)!.Value == CellValue.FromText("00123"), "Normal XLSX save/reopen retains formatted formulas and text types");
            await WaitForFastGridPaintAsync(_grid.Control); CaptureGridEvidence(Path.Combine(folder, "format-efficiency.png"));
            File.WriteAllLines(Path.Combine(folder, "format-efficiency-result.txt"), new[] { $"PASS: {checks.Count} checks" }.Concat(checks));
        }
        finally { CancelFormatBrush(); _suppressAutoCalculation = false; }
    }
}
