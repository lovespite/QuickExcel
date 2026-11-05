using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class MemoryClipboard : IClipboardExchange
    {
        public ClipboardExchange Value = new(null);
        public Task<ClipboardExchange> ReadAsync() => Task.FromResult(Value);
        public Task WriteAsync(ClipboardExchange value) { Value = value; return Task.CompletedTask; }
    }
    private async Task RunGridSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? error = null;
        void Check(bool result, string label) { if (!result) throw new InvalidOperationException(label); checks.Add(label); }
        void PressKey(Key key, KeyModifiers modifiers = KeyModifiers.None) => _grid.CellEditor.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = key, KeyModifiers = modifiers });
        async Task<string> Capture(string name)
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var file = Path.Combine(directory, name + ".png");
            using var image = new RenderTargetBitmap(new PixelSize((int)_grid.Bounds.Width, (int)_grid.Bounds.Height), new Vector(96, 96));
            image.Render(_grid); image.Save(file, PngBitmapEncoderOptions.Default); return file;
        }
        try
        {
            _clipboard = new MemoryClipboard(); ShowBook();
            var id = SheetId.New();
            var cells = new Dictionary<CellPosition, CellSnapshot>
            {
                [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromText("合并标题"), 1)),
                [new(5, 0)] = new(CellInput.Parse("10")), [new(5, 1)] = new(CellInput.Parse("=A6+1")),
                [new(10, 3)] = new(CellInput.Parse("远端"))
            };
            var sheet = new WorksheetSnapshot(id, "网格", cells, rows: new Dictionary<int, RowMetadata> { [2] = new(Hidden: true) },
                mergedCells: ["A1:C2"], view: new WorksheetView { FrozenRows = 1, FrozenColumns = 1, PaneTopLeftCell = new(1, 1) });
            var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { FillColor = "FFFF0000" }]);
            var path = Path.Combine(directory, "grid.xlsq"); new TypedWorkbookStorage().Save(path, snapshot); await OpenPathAsync(path);
            _grid.ScrollTo(40, 15);
            var rendered = await Capture("frozen-merge");
            using (var bitmap = SKBitmap.Decode(rendered))
            {
                foreach (var (x, y) in new[] { (75, 40), (235, 40), (75, 68), (235, 68) })
                {
                    var color = bitmap.GetPixel(x, y);
                    Check(color.Red > 180 && color.Green < 80 && color.Blue < 80, $"Merged fill reaches frozen/scrolling quadrant at {x},{y}");
                }
            }
            await RunMergedCellInputChecksAsync(Check);
            _grid.Select(new(0, 0));
            Check(_grid.SelectionRanges.Single() == new GridRange(0, 0, 1, 2), "Merged selection expands without per-cell storage");
            PressKey(Key.Right); Check(_grid.Current == new CellPosition(0, 3), "Right arrow exits entire merged range");
            _grid.Select(new(0, 0)); PressKey(Key.Down); Check(_grid.Current.Row == 3, "Down skips merge extent and hidden row");
            _grid.ScrollTo(0, 0); await Capture("before-selection"); var renderedCount = _grid.ContentRenderCount; var recordings = _grid.CellDrawingCount;
            _grid.Select(new(5, 0)); await Capture("after-selection");
            Check(_grid.ContentRenderCount == renderedCount && _grid.CellDrawingCount == recordings, "Selection-only change reuses content raster and recorded cells");
            _grid.SelectRanges([new(5, 0, 5, 1), new(8, 0, 8, 1)], new(5, 0)); Check(_grid.SelectionRanges.Count == 2, "Disjoint selections remain rectangular ranges");
            _grid.SelectRange(new(0, 0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1), new(0, 0));
            Check(_grid.SelectionRanges.Sum(r => r.Count) == (long)WorkbookLimits.MaxRows * WorkbookLimits.MaxColumns, "Select-all stores a bounded range");
            try { _ = _grid.Selection; throw new InvalidOperationException("Unbounded enumeration succeeded"); }
            catch (InvalidOperationException ex) when (ex.Message != "Unbounded enumeration succeeded") { checks.Add("Unbounded cell enumeration is refused"); }
            await CalculateAsync();
            _grid.SelectRange(new(5, 0, 5, 1), new(5, 0)); await CopyAsync(false);
            _grid.Select(new(7, 0)); await PasteAsync("all");
            Check(Document.GetCellInput(0, 7, 1) == "=(A8+1)" && Document.GetCellValue(0, 7, 1) == "11", "Internal paste preserves types and relocates formula references");
            await History(false); Check(Document.GetCellInput(0, 7, 0) == "" && Document.GetCellInput(0, 7, 1) == "", "Paste is one undo entry");
            var memory = (MemoryClipboard)_clipboard; memory.Value = new("=1+2\t'文本");
            _grid.Select(new(8, 0)); await PasteAsync("all");
            Check(Document.Snapshot.Sheets[0].GetCell(8, 0)!.Content.Kind == CellContentKind.Literal && Document.GetCellValue(0, 8, 0) == "=1+2", "External normal paste keeps formulas literal");
            Check(Document.GetCellValue(0, 8, 1) == "'文本", "External apostrophe remains data");
            _grid.Select(new(9, 0)); await PasteAsync("formulas"); Check(Document.GetCellValue(0, 9, 0) == "3", "Explicit formula paste opts into parsing");
            _grid.Select(new(5, 0)); await CopyAsync(true); _workspace.Editor.SetCell(0, 15, 0, "changed"); Synchronize();
            _grid.Select(new(16, 0)); var before = Document.Snapshot;
            try { await PasteAsync("all"); throw new InvalidOperationException("Stale cut succeeded"); }
            catch (InvalidOperationException ex) when (ex.Message != "Stale cut succeeded") { Check(ReferenceEquals(before, Document.Snapshot), "Stale cut rejects atomically"); }
            var revision = Document.Snapshot.Revision;
            await FillAsync(new(5, 0, 5, 1), new(5, 0, 7, 1), revision);
            Check(Document.GetCellInput(0, 7, 1) == "=(A8+1)", "Fill relocates references through shared command");
            await History(false); Check(Document.GetCellInput(0, 7, 1) == "", "Fill is one undo entry");
            await RunFillHandleChecksAsync(Check);
            _grid.SelectRange(new(0, 3, 3, 3), new(0, 3)); memory.Value = new("value"); before = Document.Snapshot;
            try { await PasteAsync("all"); throw new InvalidOperationException("Hidden paste succeeded"); }
            catch (InvalidOperationException ex) when (ex.Message != "Hidden paste succeeded") { Check(ReferenceEquals(before, Document.Snapshot), "Paste never remaps across hidden rows"); }
            await Capture("final-grid"); Check(_grid.CachedNativeBytes <= 96L * 1024 * 1024, "Native content caches remain within budgets");
            _workspace.Editor.SetCell(0,24,5,"=SEQUENCE(3,2)"); await CalculateAsync();
            _grid.Select(new(26,6)); UpdateFormula();
            Check(_grid.IsCellReadOnly(_grid.Current) && !_grid.BeginEdit(),"Spill child is read-only through native editor");
            Check(_formula.IsReadOnly && _formula.Text == "=SEQUENCE(3,2)" && Document.GetCellValue(0,26,6) == "6","Spill child shows owner formula and derived value");
            await Capture("dynamic-array-child");
            _grid.Select(new(24,5)); Check(_grid.BeginEdit(),"Spill anchor remains editable");
            _grid.CellEditor.Text = "=SEQUENCE(2)"; Check(await _grid.CommitAsync(),"Spill anchor uses normal command commit"); await CalculateAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(26,6) is null,"Shrinking array clears former derived cells");
            await History(false); Check(Document.GetCellValue(0,26,6) == "6","Undo restores entire array");
            _workspace.Editor.SetCells(0,[(0,9,"Region"),(0,10,"Product"),(0,11,"Amount"),(1,9,"North"),(1,10,"A"),(1,11,"10"),
                (2,9,"South"),(2,10,"A"),(2,11,"20"),(3,9,"North"),(3,10,"B"),(3,11,"5"),(4,9,"South"),(4,10,"B"),(4,11,"0"),
                (5,9,"East"),(5,10,"B"),(5,11,"-3"),(6,9,"East"),(6,10,"A")]);
            await ExecuteAsync(new CreateTableCommand(id,new(0,9,6,11),"PivotSource"));
            _grid.Select(new(9,9));
            _dialogDriver=window=>
            {
                var dialog=(FieldsDialog)window;
                ((ComboBox)dialog.Input("column")).SelectedIndex=2;
                ((TextBox)dialog.Input("name")).Text="GridPivot";
                dialog.Validate!(); window.Close(true); return Task.CompletedTask;
            };
            await CreatePivotAsync();
            Check(CurrentPivot?.Name=="GridPivot" && Document.GetCellValue(0,14,12)=="32","Native pivot dialog creates row/column sum report");
            Check(_grid.IsCellReadOnly(new(11,10)) && !_pivotStatus.IsVisible,"Pivot results are read-only and initially current");
            _workspace.Editor.SetCell(0,1,11,"12"); Synchronize();
            Check(_pivotStatus.IsVisible && CurrentPivot!.Stale && Document.GetCellValue(0,14,12)=="32","Source edit retains old pivot values and shows persistent stale notice");
            await RefreshPivotAsync();
            Check(!CurrentPivot!.Stale && !_pivotStatus.IsVisible && Document.GetCellValue(0,14,12)=="34","Native refresh updates report and clears stale notice");
            await History(false); Check(CurrentPivot!.Stale && Document.GetCellValue(0,14,12)=="32","Pivot refresh is one undo transaction");
            await History(true); await CalculateAsync(); _grid.ScrollTo(900,240); await Capture("pivot-report");
            var pivotPath=Path.Combine(directory,"pivot.xlsq"); new TypedWorkbookStorage().Save(pivotPath,Document.Snapshot);
            Check(new TypedWorkbookStorage().Open(pivotPath).Sheets[0].Data.Pivots.Single().Name=="GridPivot","Native pivot workbook saves and reopens with definition and cache");
            _dialogDriver=window=>{window.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="确定").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));return Task.CompletedTask;}; await ConvertPivotToValuesAsync();
            Check(CurrentPivot is null && Document.GetCellValue(0,14,12)=="34","Converting pivot to values keeps displayed report");
            foreach (var (aggregation,index,expected) in new[] {(PivotAggregation.Average,2,6.8),(PivotAggregation.Min,3,-3d),(PivotAggregation.Max,4,20d)})
            {
                _grid.Select(new(39,9));
                _dialogDriver=window=>
                {
                    var dialog=(FieldsDialog)window;
                    ((ComboBox)dialog.Input("column")).SelectedIndex=2;
                    ((ComboBox)dialog.Input("function")).SelectedIndex=index;
                    ((TextBox)dialog.Input("name")).Text="Grid"+aggregation;
                    dialog.Validate!(); window.Close(true); return Task.CompletedTask;
                };
                await CreatePivotAsync();
                var report=CurrentPivot!;
                Check(report.Aggregation==aggregation && Document.Snapshot.GetSheet(id).GetCell(report.Range.Bottom,report.Range.Right)!.Value==CellValue.FromNumber(expected),
                    $"Native pivot dialog routes {aggregation} and computes grand total from source records");
                await History(false);
                Check(!Document.Snapshot.GetSheet(id).Data.Pivots.Any(),$"Undo removes the complete {aggregation} report");
            }
            var validation = new WorksheetValidation([new(20, 0, 20, 0)], ValidationKind.Whole, "=1", "=5", alert: ValidationAlert.Warning);
            await ExecuteAsync(new SetValidationCommand(id, validation)); _grid.Select(new(20, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "9";
            var answer = new TaskCompletionSource<bool>(); _validationChoice = _ => answer.Task;
            var pending = _grid.CommitAsync(); Check(!pending.IsCompleted && _grid.IsCommitPending, "Native validation awaits confirmation without blocking dispatcher");
            _workspace.Editor.SetCell(0, 21, 0, "concurrent"); answer.SetResult(true);
            Check(!await pending && Document.GetCellInput(0, 20, 0) == "" && _grid.IsEditing, "Concurrent edit rejects old validation commit and retains draft");
            _validationChoice = _ => Task.FromResult(false);
            Check(!await _grid.CommitAsync() && _grid.IsEditing, "Cancelled validation keeps original cell draft");
            _validationChoice = _ => Task.FromResult(true);
            Check(await _grid.CommitAsync() && Document.GetCellValue(0, 20, 0) == "9", "Confirmed native validation commits original owner");
            _validationChoice = null; await History(false);
            Check(Document.GetCellInput(0, 20, 0) == "" && Document.GetCellInput(0, 21, 0) == "concurrent", "Validation override contributes one history transaction");
            _grid.Select(new(30, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "=SU"; _grid.CellEditor.CaretIndex = 3;
            _formulaSuggestions.SelectedItem = "SUM";
            Check(AcceptCompletion() && _grid.CellEditor.Text == "=SUM(", "Native formula completion uses shared bounded suggestions");
            _grid.CellEditor.Text = "=A6"; _grid.CellEditor.SelectionStart = _grid.CellEditor.SelectionEnd = 3;
            PressKey(Key.F4); Check(_grid.CellEditor.Text == "=$A$6", "Routed F4 cycles absolute references in native editor");
            _grid.Cancel(); _grid.Select(new(30, 0)); _formula.Focus(); _formula.Text = "=SUM()"; _formula.CaretIndex = 5; _formula.SelectionStart = _formula.SelectionEnd = 5;
            Check(PickFormulaReference(new(5, 0, 5, 0)) && PickFormulaReference(new(5, 0, 7, 1)) && _formula.Text == "=SUM(A6:B8)", "Formula-bar reference drag replaces one draft slot");
            _formulaGesture = null; Check(await CommitFormulaAsync() && Document.GetCellInput(0, 30, 0) == "=SUM(A6:B8)", "Picked formula commits to original owner");
            await RunFormulaReferenceChecksAsync(Check);
            _grid.Select(new(35, 0)); _workspace.Editor.SetCell(0, 35, 0, "自动列宽需要完整测量远端内容"); Synchronize();
            var oldWidth = DrawingGeometry.ColumnSize(Document.Snapshot.Sheets[0], 0); await AutoFitAsync(SheetAxis.Column, 0);
            Check(DrawingGeometry.ColumnSize(Document.Snapshot.Sheets[0], 0) > oldWidth, "Auto-fit measures offscreen stored content");
            await History(false); Check(DrawingGeometry.ColumnSize(Document.Snapshot.Sheets[0], 0) == oldWidth, "Auto-fit is undoable dimension metadata");
            var bold = _mutationControls.OfType<Button>().Single(b => global::Avalonia.Automation.AutomationProperties.GetName(b) == "加粗");
            bold.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation;
            Check(CurrentStyle.Bold, "Native ribbon click reaches shared style command");
            await History(false); Check(!CurrentStyle.Bold, "Ribbon style is one undo entry");
            _grid.SelectRange(new(35, 0, 36, 1), new(35, 0));
            await ApplyStyleAsync(new(Horizontal: 3, Vertical: 3, Bold: true));
            bool Aligned(int vertical) => new[] { new CellPosition(35, 0), new(35, 1), new(36, 0), new(36, 1) }
                .All(p => Document.Snapshot.Styles[Document.Snapshot.Sheets[SheetIndex].GetStyleId(p.Row, p.Column)]
                    is { Horizontal: 3, Bold: true } style && style.Vertical == vertical);
            foreach (var (title, vertical) in new[] { ("顶端对齐", 0), ("垂直居中", 1), ("底端对齐", 2) })
            {
                var button = _mutationControls.OfType<Button>().Single(b => global::Avalonia.Automation.AutomationProperties.GetName(b) == title);
                button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation;
                Check(Aligned(vertical), $"{title} formats the whole selection without changing horizontal alignment or font");
                await History(false); Check(Aligned(3), $"{title} undoes the entire selection in one step");
                await History(true); Check(Aligned(vertical), $"{title} redoes the entire selection in one step");
                await History(false);
            }
            await History(false); _grid.Select(new(35, 0));
            await FreezeAsync(2, 2); Check(Document.Snapshot.Sheets[0].View.FrozenRows == 2 && Document.Snapshot.Sheets[0].View.FrozenColumns == 2, "Native freeze command persists both axes");
            _grid.ScrollTo(180, 100); Check(Document.Snapshot.Sheets[0].View.PaneTopLeftCell.Row >= 2, "Scroll captures persisted pane navigation");
            await Capture("native-ribbon");
            await RunDataFeatureChecksAsync(Check);
            await RunFileObjectChecksAsync(directory, Check);
            await Capture("data-features");
            await RunScrollRangeChecksAsync(directory, Check);
            await Capture("bounded-scrollbars");
            await RunCellEditorAlignmentChecksAsync(directory, Check);
            await RunCellEditorAppearanceChecksAsync(directory, Check);
        }
        catch (Exception ex) { error = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = error is null, checks, error = error?.ToString(),
            scaling = RenderScaling, contentRenders = _grid.ContentRenderCount, recordings = _grid.CellDrawingCount, cachedNativeBytes = _grid.CachedNativeBytes }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}
