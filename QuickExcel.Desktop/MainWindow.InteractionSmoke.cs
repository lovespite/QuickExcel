using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunInteractionSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var cases = new List<string>();
        void Check(bool condition, string scenario)
        {
            if (!condition) throw new Exception("M5 interaction: " + scenario);
            cases.Add(scenario);
        }
        async Task RoutedKeyAsync(Key key)
        {
            var source = Keyboard.FocusedElement as UIElement ?? SheetGrid;
            var args = new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(this), 0, key)
            { RoutedEvent = Keyboard.PreviewKeyDownEvent };
            source.RaiseEvent(args);
            if (!args.Handled) throw new Exception("Grid did not handle routed key: " + key);
            await _gridInput;
        }
        var selectionDragReady = false;
        async Task ClickCellAsync(int row, int column)
        {
            SheetGrid.ScrollIntoView(_rows[row], SheetGrid.Columns[column]);
            SheetGrid.UpdateLayout();
            var content = SheetGrid.Columns[column].GetCellContent(_rows[row]);
            // Prepared text inputs are transparent to hit testing, just as on the real grid.
            var source = content.IsHitTestVisible ? content : (UIElement)content.Parent;
            var args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                { RoutedEvent = Mouse.PreviewMouseDownEvent };
            source.RaiseEvent(args);
            // The native DataGrid starts dragging only when the clicked cell does not
            // already own keyboard focus. Synthetic input has no OS mouse capture provider.
            var cell = content.Parent as DataGridCell ?? source as DataGridCell;
            selectionDragReady = !args.Handled && cell is { IsKeyboardFocusWithin: false };
            args.RoutedEvent = Mouse.MouseDownEvent;
            source.RaiseEvent(args);
            SheetGrid.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                { RoutedEvent = Mouse.MouseUpEvent });
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        }
        async Task DoubleClickCellAsync(int row, int column)
        {
            await ClickCellAsync(row, column);
            var content = SheetGrid.Columns[column].GetCellContent(_rows[row]);
            SheetGrid.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
            { RoutedEvent = Control.PreviewMouseDoubleClickEvent, Source = content.Parent });
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        }
        InstallDocument(WorkbookDocument.CreateBlank(_services), Path.Combine(folder, "interaction.xlsq"), "交互验收.xlsq");
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        GridSurface.Select(new(0, 0));
        await RoutedKeyAsync(Key.F2);
        var editor = SheetGrid.Columns[0].GetCellContent(_rows[0]) as TextBox ?? throw new Exception("F2 editor missing.");
        editor.Text = "42";
        await RoutedKeyAsync(Key.Tab);
        Check(_document.GetCellValue(0, 0, 0) == "42" && _selectedColumn == 1 && _selectedRow == 0, "F2 / Tab commits once and moves right");
        await HandleGridKeyAsync(Key.Tab, ModifierKeys.Shift, false);
        Check(_selectedColumn == 0, "Shift+Tab moves left");
        await HandleGridKeyAsync(Key.F2, ModifierKeys.None, false);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[0]);
        editor.Text = "discard";
        await HandleGridKeyAsync(Key.Escape, ModifierKeys.None, true);
        Check(_document.GetCellValue(0, 0, 0) == "42", "Escape cancels grid draft");
        await HandleGridKeyAsync(Key.Enter, ModifierKeys.None, false);
        Check(_selectedRow == 1, "Enter moves down");
        await HandleGridKeyAsync(Key.Enter, ModifierKeys.Shift, false);
        Check(_selectedRow == 0, "Shift+Enter moves up");
        await HandleGridKeyAsync(Key.Right, ModifierKeys.Shift, false);
        await HandleGridKeyAsync(Key.Down, ModifierKeys.Shift, false);
        Check(GridRange.Rectangle(GridSurface.Selection) == new GridRange(0, 0, 1, 1), "Shift+arrows extends rectangle from stable anchor");
        var revision = _document.Snapshot.Revision;
        await PasteSelectionAsync("7", false);
        Check(_document.Snapshot.Revision == revision + 1 && _document.GetCellValue(0, 1, 1) == "7", "Scalar paste fills selection in one transaction");
        await UndoAsync(false);
        Check(_document.GetCellValue(0, 0, 0) == "42" && _document.GetCellValue(0, 1, 1) == "", "One undo restores entire pasted rectangle");
        await UndoAsync(true);
        GridSurface.Select(new(0, 0));
        await PasteSelectionAsync("=1+2\t00123\n\"多行\n文本\"\tTRUE", false);
        Check(_document.GetCellInput(0, 0, 0) == "'=1+2" && _document.GetCellValue(0, 1, 0) == "多行\n文本", "Normal paste preserves formulas as text and quoted multiline fields");
        GridSurface.Select(new(2, 0));
        await PasteSelectionAsync("=1+2", true);
        Check(_document.GetCellInput(0, 2, 0) == "=1+2" && _document.GetCellValue(0, 2, 0) == "3" && !_document.ScriptsEnabled,
            "Explicit formula paste calculates without granting script authorization");
        revision = _document.Snapshot.Revision;
        try { await PasteSelectionAsync("ok\t" + new string('x', 32768), false); throw new Exception("Invalid paste accepted."); }
        catch (ArgumentException) { }
        Check(_document.Snapshot.Revision == revision, "Invalid paste leaves entire document unchanged");
        SheetGrid.SelectedCells.Clear();
        SheetGrid.SelectedCells.Add(new(_rows[0], SheetGrid.Columns[0]));
        SheetGrid.SelectedCells.Add(new(_rows[1], SheetGrid.Columns[1]));
        try { CopySelection(); throw new Exception("Discontiguous copy accepted."); }
        catch (ArgumentException) { }
        Check(_document.Snapshot.Revision == revision, "Discontiguous copy refuses unselected intervening data");
        GridSurface.SelectRange(new(0, 0, 1, 1), new(0, 0));
        await HandleGridKeyAsync(Key.Delete, ModifierKeys.None, false);
        Check(_document.GetCellValue(0, 0, 0) == "" && _document.GetCellValue(0, 1, 1) == "", "Delete clears full selection");
        await UndoAsync(false);
        Check(_document.GetCellValue(0, 0, 0) == "=1+2" && _document.GetCellValue(0, 1, 1) == "TRUE", "Delete has one undo");

        GridSurface.Select(new(3, 0));
        FormulaBar.Focus();
        Check(!FormulaBar.CanUndo, "Formula bar cannot undo programmatic display text from an earlier cell");
        FormulaBar.SelectedText = "temporary draft";
        FormulaBar.Undo();
        Check(FormulaBar.Text == "" && _document.GetCellValue(0, 0, 0) == "=1+2",
            "Formula bar undo affects only its current draft and preserves workbook history");
        FormulaBar.Text = "公式栏中文";
        await NavigateGridAsync(Key.Tab, ModifierKeys.None);
        Check(_document.GetCellValue(0, 3, 0) == "公式栏中文" && _selectedColumn == 1, "Formula bar shares Tab commit/navigation");
        GridSurface.Select(new(4, 0));
        await HandleGridKeyAsync(Key.F2, ModifierKeys.None, false);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[4]);
        var composition = new TextComposition(InputManager.Current, editor, "中文");
        editor.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, composition) { RoutedEvent = TextCompositionManager.PreviewTextInputStartEvent });
        editor.Text = "中文组合";
        revision = _document.Snapshot.Revision;
        Check(_editState.IsComposing && !CommitEditors() && !await SaveAsync(false) && _document.Snapshot.Revision == revision,
            "Routed IME composition blocks partial commit and save");
        var switchEditor = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
            { RoutedEvent = Mouse.PreviewMouseDownEvent };
        FormulaBar.RaiseEvent(switchEditor);
        Check(switchEditor.Handled && editor.IsKeyboardFocused,
            "Composition rejects mouse focus switch before Windows can commit an unfinished Pinyin buffer");
        await HandleGridKeyAsync(Key.Tab, ModifierKeys.None, true);
        Check(_selectedRow == 4 && _selectedColumn == 0, "Composition blocks premature navigation");
        editor.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, composition) { RoutedEvent = TextCompositionManager.PreviewTextInputEvent });
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        await HandleGridKeyAsync(Key.Enter, ModifierKeys.None, true);
        Check(!_editState.IsComposing && _document.GetCellValue(0, 4, 0) == "中文组合" && _selectedRow == 5,
            "Completed composition commits Chinese text once");
        Check(await SaveAsync(false), "Edited workbook saves through actual desktop path");
        using (var reopened = WorkbookDocument.Open(_filePath!, _services))
            Check(reopened.GetCellValue(0, 4, 0) == "中文组合" && reopened.GetCellInput(0, 0, 0) == "'=1+2", "Save/reopen preserves text and formula input policy");
        var invalid = _document.Snapshot.Revision;
        await HandleGridKeyAsync(Key.F2, ModifierKeys.None, false);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[5]);
        editor.Text = new string('x', 32768);
        await HandleGridKeyAsync(Key.Tab, ModifierKeys.None, true);
        Check(_selectedRow == 5 && _selectedColumn == 0 && editor.Text.Length == 32768 && _document.Snapshot.Revision == invalid,
            "Rejected grid draft remains editable and selection stays at its address");
        await HandleGridKeyAsync(Key.Escape, ModifierKeys.None, true);
        ApplyValues([(6, 0, "=1+2")]);
        await RecalculateAsync();
        GridSurface.Select(new(6, 0));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[6]);
        Check(_gridEditorArmed && !_editState.GridEditing && editor.Text == "3",
            "Focused cell prepares native IME editor while showing calculated value without a draft");
        revision = _document.Snapshot.Revision;
        await RoutedKeyAsync(Key.Right);
        Check(_selectedColumn == 1 && _document.Snapshot.Revision == revision && _document.GetCellInput(0, 6, 0) == "=1+2",
            "Navigation from prepared editor preserves original formula and document history");
        GridSurface.Select(new(6, 0));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        await RoutedKeyAsync(Key.F2);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[6]);
        Check(editor.Text == "=1+2" && editor.CaretIndex == editor.Text.Length && editor.IsHitTestVisible && editor.SelectionOpacity > 0,
            "F2 initializes raw formula before focus and places caret at end");
        await HandleGridKeyAsync(Key.Escape, ModifierKeys.None, true);
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        var inputTarget = (UIElement)Keyboard.FocusedElement;
        inputTarget.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice,
            new TextComposition(InputManager.Current, inputTarget, "n"))
            { RoutedEvent = TextCompositionManager.TextInputEvent });
        editor = SheetGrid.Columns[0].GetCellContent(_rows[6]) as TextBox ?? throw new Exception("Direct input editor missing.");
        Check(editor.Text == "n" && !_gridEditorArmed && editor.IsHitTestVisible && editor.SelectionOpacity > 0,
            "Direct text input exposes the editor, retains initiating character and replaces original formula");
        await RoutedKeyAsync(Key.Enter);
        Check(_document.GetCellValue(0, 6, 0) == "n" && _selectedRow == 7,
            "Direct text input commits exactly once and moves down");
        FormulaBar.Focus();
        await RecalculateAsync();
        Check(FormulaBar.IsKeyboardFocused,
            "Background recalculation restores formula bar keyboard focus after disabling the workspace");
        FormulaBar.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice,
            new TextComposition(InputManager.Current, FormulaBar, "after calculation"))
            { RoutedEvent = TextCompositionManager.TextInputEvent });
        Check(FormulaBar.Text == "after calculation" && _formulaPending,
            "Text after recalculation reaches the visible formula editor and remains a draft");
        _editState.EndFormula();
        UpdateFormulaBar();
        InstallDocument(WorkbookDocument.Open(_filePath!, _services), _filePath, "重开交互验收.xlsq", fileVersion: QuickExcel.Storage.FileVersions.Read(_filePath!));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        Check(_selectedRow == 0 && _selectedColumn == 0 &&
            GridSurface.Selection.SequenceEqual(new[] { new CellPosition(0, 0) }) &&
            SheetGrid.CurrentItem == _rows[0] && SheetGrid.CurrentColumn == SheetGrid.Columns[0] && CellAddress.Text == "A1",
            "Reopening the same sheet identity resets visual selection, current editor and formula address together");

        ApplyValues([(6, 0, "=1+2")]);
        await RecalculateAsync();
        revision = _document.Snapshot.Revision;
        await ClickCellAsync(6, 0);
        editor = SheetGrid.Columns[0].GetCellContent(_rows[6]) as TextBox ?? throw new Exception("Selected cell input context missing.");
        Check(_selectedRow == 6 && _selectedColumn == 0 && _gridEditorArmed && !_editState.GridEditing &&
            editor.Text == "3" && _document.Snapshot.Revision == revision,
            "First mouse click selects a different cell and displays its calculated value without starting a draft");
        Check(!editor.IsHitTestVisible && editor.CaretBrush == Brushes.Transparent && editor.SelectionOpacity == 0 &&
            editor.BorderThickness == new Thickness(0) && editor.IsKeyboardFocused,
            "Selection hides the text caret, text selection and editor border while preserving native keyboard input context");
        Check(editor.Parent is DataGridCell { BorderThickness.Left: 1, BorderBrush: SolidColorBrush border } &&
            border.Color == Color.FromRgb(24, 117, 83),
            "Selected cell retains a visible green cell border independent of its workbook border style");
        CaptureGridEvidence(Path.Combine(folder, "QuickExcel-click-selected.png"));
        await ClickCellAsync(6, 0);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[6]);
        Check(_gridEditorArmed && !_editState.GridEditing && !editor.IsHitTestVisible && editor.Text == "3" &&
            selectionDragReady && _document.Snapshot.Revision == revision,
            "Repeated single click preserves formula and reaches native drag selection without editing");
        await DoubleClickCellAsync(6, 0);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[6]);
        Check(!_gridEditorArmed && _editState.GridEditing && editor.Text == "=1+2" && editor.IsHitTestVisible &&
            editor.SelectionOpacity > 0 && editor.CaretBrush != Brushes.Transparent && _document.Snapshot.Revision == revision,
            "Double click starts editing the original formula without committing it");
        CaptureGridEvidence(Path.Combine(folder, "QuickExcel-click-editing.png"));
        await RoutedKeyAsync(Key.Escape);
        Check(_document.GetCellInput(0, 6, 0) == "=1+2" && _document.Snapshot.Revision == revision,
            "Escape after mouse editing preserves the original formula and history");
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        await DoubleClickCellAsync(6, 0);
        editor = (TextBox)SheetGrid.Columns[0].GetCellContent(_rows[6]);
        editor.Text = "=2+3";
        await ClickCellAsync(6, 1);
        Check(_document.GetCellInput(0, 6, 0) == "=2+3" && _document.Snapshot.Revision == revision + 1 &&
            _selectedColumn == 1 && _gridEditorArmed && !_editState.GridEditing,
            "Clicking another cell commits the previous draft once and only selects the destination");
        GridSurface.SelectRange(new(6, 0, 7, 1), new(6, 0));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        await ClickCellAsync(6, 0);
        Check(GridSurface.Selection.SequenceEqual(new[] { new CellPosition(6, 0) }) && _gridEditorArmed && !_editState.GridEditing,
            "First click on an active cell in a multi-cell selection collapses the range without editing");
        await ClickCellAsync(6, 0);
        Check(!_editState.GridEditing && _gridEditorArmed && selectionDragReady,
            "Another single click after collapsing a range still reaches native drag selection");
        await DoubleClickCellAsync(6, 0);
        Check(_editState.GridEditing && !_gridEditorArmed, "Double click after collapsing a range enters cell editing");
        await RoutedKeyAsync(Key.Escape);
        await RunCellLayoutSmokeAsync(folder, Check, ClickCellAsync, DoubleClickCellAsync);
        RefreshSheet();
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        CaptureGridEvidence(Path.Combine(folder, "QuickExcel-interaction.png"));
        File.WriteAllText(Path.Combine(folder, "interaction-cases.json"), JsonSerializer.Serialize(new
        {
            cases, imeEvidence = "Synthetic WPF routed composition; physical IME candidate window not automated."
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "interaction-smoke-result.txt"), $"PASS: {cases.Count} WPF assertions; mouse/keyboard/edit/selection/paste/undo/composition/save/reopen. Mouse and IME composition routing is synthetic.");
    }

    private void CaptureGridEvidence(string path)
    {
        UpdateLayout();
        var root = (FrameworkElement)Content;
        var dpi = VisualTreeHelper.GetDpi(root);
        var bitmap = new RenderTargetBitmap((int)(root.ActualWidth * dpi.DpiScaleX), (int)(root.ActualHeight * dpi.DpiScaleY),
            96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(root);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using var file = File.Create(path);
        encoder.Save(file);
    }

    private async Task RunGridBenchmarkAsync(string folder)
    {
        _suppressAutoCalculation = true;
        try { await MeasureGridAsync(folder); }
        finally { _suppressAutoCalculation = false; _calculationPending = false; }
    }

    private async Task MeasureGridAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var workloads = new List<object>();
        foreach (var sparse in new[] { false, true })
        {
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            var count = sparse ? 20_000 : 100_000;
            for (var i = 0; i < count; i++)
            {
                var position = sparse ? new CellPosition(i * 4999 % 100_000, i * 17 % 256) : new CellPosition(i / 10, i % 10);
                cells[position] = new(CellContent.FromLiteral(CellValue.FromNumber(i)));
            }
            var document = new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new(SheetId.New(), sparse ? "稀疏20k" : "密集100k", cells)]), _services);
            var before = GC.GetTotalMemory(true);
            var timer = Stopwatch.StartNew();
            InstallDocument(document, null, "M5 网格测量.xlsq");
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            var bindMs = timer.Elapsed.TotalMilliseconds;
            var after = GC.GetTotalMemory(true);
            var firstRows = _rows.CreatedCount;
            if (firstRows > 1000 || _grid.RealizedRows > 1000)
                throw new Exception("Grid bind eagerly realized too many rows.");
            var scroll = new List<double>();
            var horizontal = new List<double>();
            var input = new List<double>();
            var command = new List<double>();
            var refresh = new List<double>();
            var selection = new List<double>();
            for (var i = 0; i < 35; i++)
            {
                timer.Restart();
                SheetGrid.ScrollIntoView(_rows[(i * 271) % _rowCount], SheetGrid.Columns[0]);
                await Dispatcher.InvokeAsync(() => { SheetGrid.UpdateLayout(); }, DispatcherPriority.ApplicationIdle);
                if (i >= 5) scroll.Add(timer.Elapsed.TotalMilliseconds);
                timer.Restart();
                SheetGrid.ScrollIntoView(_rows[(i * 271) % _rowCount], SheetGrid.Columns[(i * 17) % _columnCount]);
                await Dispatcher.InvokeAsync(() => { SheetGrid.UpdateLayout(); }, DispatcherPriority.ApplicationIdle);
                if (i >= 5) horizontal.Add(timer.Elapsed.TotalMilliseconds);
                timer.Restart();
                GridSurface.SelectRange(new(0, 0, 9, 9), new(0, 0));
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                if (i >= 5) selection.Add(timer.Elapsed.TotalMilliseconds);
                timer.Restart();
                ApplyValues([(0, 0, (i + 100_000).ToString(System.Globalization.CultureInfo.InvariantCulture))]);
                var commandMs = timer.Elapsed.TotalMilliseconds;
                RefreshSheet();
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                if (i >= 5)
                {
                    var totalMs = timer.Elapsed.TotalMilliseconds;
                    input.Add(totalMs);
                    command.Add(commandMs);
                    refresh.Add(totalMs - commandMs);
                }
            }
            static object Samples(List<double> values) => new { rawMs = values, p50Ms = values.Order().ElementAt(14), p95Ms = values.Order().ElementAt(28) };
            workloads.Add(new
            {
                name = sparse ? "20k sparse / 100k rows x 256 columns" : "P02 100k cells / 10k rows x 10 columns",
                cellCount = cells.Count, logicalRows = _rows.Count, logicalColumns = _columnCount, bindMs,
                viewManagedBytesDelta = after - before, createdRowsAfterBind = firstRows, createdRowsAfterRun = _rows.CreatedCount,
                realizedRows = _grid.RealizedRows, refreshedRows = _grid.LastRefreshRows,
                scroll = Samples(scroll), horizontal = Samples(horizontal), input = Samples(input), command = Samples(command),
                refresh = Samples(refresh), selection100 = Samples(selection)
            });
            CaptureGridEvidence(Path.Combine(folder, sparse ? "grid-sparse.png" : "grid-dense.png"));
        }
        using var process = Process.GetCurrentProcess();
        File.WriteAllText(Path.Combine(folder, "grid-benchmark.json"), JsonSerializer.Serialize(new
        {
            generatedUtc = DateTime.UtcNow, os = Environment.OSVersion.ToString(), runtime = Environment.Version.ToString(),
            processorCount = Environment.ProcessorCount, dpi = VisualTreeHelper.GetDpi(this).PixelsPerInchX,
            window = new { ActualWidth, ActualHeight }, renderTier = RenderCapability.Tier >> 16,
            peakWorkingSetBytes = process.PeakWorkingSet64, privateBytes = process.PrivateMemorySize64,
            method = "Release WPF; 5 warmups + 30 samples; operation to Dispatcher.ApplicationIdle after layout, not display presentation latency; no calculation in input samples; exploration, not frozen QX-B01 Beta gate.",
            workloads
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "grid-benchmark-result.txt"), "PASS: both workloads completed; raw samples and environment recorded. Performance budgets require interpretation of grid-benchmark.json.");
    }
}
