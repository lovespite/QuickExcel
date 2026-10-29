using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunProductSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var cases = new List<string>();
        void Check(bool condition, string message) { if (!condition) throw new Exception("M5 product: " + message); cases.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle); UpdateLayout(); }
        _suppressAutoCalculation = true;
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "未命名.xlsq");
            _hasWorkbook = false; ShowStartPage(); await Idle();
            Check(StartPage.IsVisible && !ResumeButton.IsVisible && !Workspace.IsVisible, "Clean start page has no demo workbook or script authorization");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-start.png"));
            New_Click(this, new()); await Idle();
            Check(WorkbookPage.IsVisible && !ScriptPanel.IsVisible && _hasWorkbook, "New workbook opens the full-width grid");
            Check(DockPanel.GetDock(SheetNavigation) == Dock.Bottom, "Worksheet navigation defaults to bottom");
            ApplyValues([(0, 0, "apple"), (1, 0, "apple"), (2, 0, "pear")]);
            FormulaBar.Focus(); FormulaBar.Text = "draft";
            ScriptPage_Click(this, new()); await Idle();
            Check(ScriptPanel.IsVisible && !WorkbookPage.IsVisible && _document.GetCellInput(0, 0, 0) == "draft", "Switching to scripts commits the original-cell draft");
            ScriptEditor.Text = "quickExcel.functions.register('TWICE', x => x * 2);";
            WorkbookPage_Click(this, new()); await Idle();
            Check(_document.ScriptSource.Contains("TWICE") && WorkbookPage.IsVisible, "Script draft survives page switch without execution");
            Check(!_document.ScriptsEnabled, "Source edit never grants execution permission");
            for (var i = 0; i < 5; i++) AddSheet_Click(this, new());
            Check(SheetPicker.Items.Count == 6, "All worksheets are exposed as tabs");
            foreach (var placement in Enum.GetValues<SheetNavigationPlacement>())
            {
                ApplyPreferences(new(placement)); await Idle();
                Check(DockPanel.GetDock(SheetNavigation) == (placement == SheetNavigationPlacement.Bottom ? Dock.Bottom : placement == SheetNavigationPlacement.Left ? Dock.Left : Dock.Right), $"Navigation placement {placement}");
                SheetPicker.SelectedIndex = 0; await Idle();
                Check(_sheetIndex == 0 && _document.GetCellInput(0, 0, 0) == "draft", $"Sheet switch uses document identity with {placement} navigation");
            }
            _settings.Save(new(SheetNavigationPlacement.Right, 125));
            var loaded = new DesktopSettings(Path.Combine(folder, "settings.json")); loaded.Load();
            Check(loaded.Value.Sheets == SheetNavigationPlacement.Right && loaded.Value.Zoom == 125, "Navigation and zoom persist in isolated settings");
            ApplyPreferences(new());
            var matches = WorkbookSearch.Find(_document.Snapshot, 0, new("apple", SearchScope.Workbook));
            var revision = _document.Snapshot.Revision;
            _editor.Execute(WorkbookSearch.Replace(matches, new("apple"), "orange")); RefreshSheet();
            Check(_document.Snapshot.Revision == revision + 1 && _document.GetCellInput(0, 1, 0) == "orange", "Desktop replacement commits one transaction");
            await UndoAsync(false);
            Check(_document.GetCellInput(0, 1, 0) == "apple", "One desktop undo restores replaced content");
            _smokeSearchDriver = dialog =>
            {
                try
                {
                    var fields = ProductChildren<TextBox>(dialog).ToArray();
                    fields.Single(t => AutomationProperties.GetName(t) == "查找").Text = "apple";
                    fields.Single(t => AutomationProperties.GetName(t) == "替换为").Text = "replaced";
                    var buttons = ProductChildren<Button>(dialog).ToArray();
                    buttons.Single(b => b.Content as string == "查找全部").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    var results = ProductChildren<ListBox>(dialog).Single();
                    Check(results.Items.Count == 1, "Actual search dialog finds displayed values");
                    buttons.Single(b => b.Content as string == "全部替换").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    Check(_document.GetCellInput(0, 1, 0) == "replaced", "Actual Replace All button commits requested text");
                }
                finally { dialog.Close(); }
            };
            Find_Click(this, new()); _smokeSearchDriver = null;
            await UndoAsync(false);
            Check(_document.GetCellInput(0, 1, 0) == "apple", "Dialog replacement is undone as one document operation");
            var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
            {
                ["main"] = new("scripts/main.js", "import { twice } from 'helpers';\nquickExcel.functions.register('TWICE', twice);"),
                ["helpers"] = new("scripts/helpers.js", "export const twice = value => value * 2;")
            }, ["main"], scriptApiVersion: 2);
            _editor.Execute(new UpdateScriptModulesCommand(modules));
            _refreshing = true; LoadScriptModules(); _refreshing = false;
            ScriptPage_Click(this, new()); ModulePicker.SelectedItem = "helpers";
            ScriptEditor.Text = "export const twice = value => value * 3;";
            ModulePicker.SelectedItem = "main";
            Check(_document.Snapshot.ScriptModules.Modules["helpers"].Source.Contains("* 3") && ScriptEditor.Text.Contains("import"), "Module edits are committed to their own logical ID");
            ModulePicker.SelectedItem = "helpers";
            ScriptEditor.Text = new string('x', 1024 * 1024 + 1);
            ModulePicker.SelectedItem = "main";
            Check(_selectedModule == "helpers" && (string)ModulePicker.SelectedItem == "helpers" && !_document.ScriptsEnabled, "Rejected module draft blocks switching and revokes authorization");
            ScriptEditor.Text = "export const twice = value => value * 2;";
            ModulePicker.SelectedItem = "main";
            Check(AutomationProperties.GetName(ScriptEditor).Length > 0 && AutomationProperties.GetName(SheetPicker).Length > 0, "Editor and navigation expose accessible names");
            ScriptEditor.CaretIndex = ScriptEditor.Text.Length;
            Check(ScriptPosition.Text.StartsWith("2:"), "Script caret line and column are visible");
            _editor.SetCell(0, 0, 1, "=TWICE(5)");
            await RunScriptsAsync();
            Check(_document.GetCellValue(0, 0, 1) == "10", "Edited multi-module sources run through the isolated worker");
            ModulePicker.SelectedItem = "helpers"; ScriptEditor.Text = "export const twice = ;";
            await RunScriptsAsync();
            Check(DiagnosticsList.Items.Count > 0 && !_document.ScriptsEnabled, "Syntax error produces visible diagnostics and revokes authorization");
            File.WriteAllText(Path.Combine(folder, "module-diagnostic.txt"), string.Join("\n", _document.ScriptErrors));
            DiagnosticsList.SelectedIndex = 0; ActivateDiagnostic();
            Check(_selectedModule == "helpers", "Syntax diagnostic locates the failing module");
            ScriptEditor.Text = "export const twice = value => value * 2;";
            await RunScriptsAsync(); ModulePicker.SelectedItem = "main";
            Check(_document.GetCellValue(0, 0, 1) == "10", "Corrected module can be explicitly re-enabled");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-scripts.png"));
            WorkbookPage_Click(this, new());
            InstallProductFixture(); await Idle();
            GridSurface.Select(new(0, 0));
            Check(_artwork is not null && _document.Snapshot.Sheets[0].Images.Length == 1, "Artwork renderer is attached to actual WPF grid");
            await Idle();
            Check(ProductCellPixel(SheetGrid, 0, 1) == Color.FromRgb(230, 241, 235), "Merged background paints across the second cell");
            Check(ProductCellPixel(SheetGrid, 2, 6) == Color.FromRgb(22, 99, 75), "Embedded image pixels are visible above the grid");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-workbook.png"));
            ApplyPreferences(new(SheetNavigationPlacement.Left)); await Idle();
            Check(ProductCellPixel(SheetGrid, 2, 6) == Color.FromRgb(22, 99, 75), "Image remains visible with left worksheet navigation");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-sheets-left.png"));
            ApplyPreferences(new(SheetNavigationPlacement.Right)); await Idle();
            Check(ProductCellPixel(SheetGrid, 0, 1) == Color.FromRgb(230, 241, 235) && ProductCellPixel(SheetGrid, 2, 6) == Color.FromRgb(22, 99, 75), "Merge and image remain visible with right worksheet navigation");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-sheets-right.png"));
            ApplyPreferences(new(SheetNavigationPlacement.Bottom, 150)); await Idle();
            Check(((ScaleTransform)SheetPageContent.LayoutTransform).ScaleX == 1.5, "Grid zoom uses layout-aware WPF scaling");
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-zoom.png"));
            ApplyPreferences(new());
            var sheetId = _document.Snapshot.Sheets[0].Id;
            Check(ExecuteWorkbookCommand(new ShiftAxisCommand(sheetId, SheetAxis.Row, 0, 1)), "Insert row transforms artwork through desktop command path");
            Check(_document.Snapshot.Sheets[0].MergedCells[0] == "A2:D2" && _document.Snapshot.Sheets[0].Images[0].RowIndex == 3, "Merge and image follow inserted row");
            await UndoAsync(false);
            Check(_document.Snapshot.Sheets[0].MergedCells[0] == "A1:D1" && _document.Snapshot.Sheets[0].Images[0].RowIndex == 2, "Undo restores merge and image anchors");
            GridSurface.Select(new(4, 2)); Freeze_Click(new MenuItem(), new()); await Idle();
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-artwork-frozen.png"));
            Check(FrozenGrid.IsVisible && _frozenArtwork is not null, "Frozen rows have their own artwork projection");
            Check(ProductCellPixel(FrozenGrid, 0, 1) == Color.FromRgb(230, 241, 235), "Frozen merge is painted in the frozen viewport");
            Freeze_Click(new MenuItem { Tag = "clear" }, new());
            var path = Path.Combine(folder, "product-workbook.xlsq");
            if (!await SaveToPathAsync(path, false, new(null))) throw new Exception("Product save failed: " + LogText.Text);
            Check(true, "Product workbook saves through fingerprint-checked storage");
            await OpenPathAsync(path); await Idle();
            Check(_document.Snapshot.Sheets[0].MergedCells[0] == "A1:D1" && _document.Snapshot.Sheets[0].Images[0].RowIndex == 2 && !_document.ScriptsEnabled,
                "Saved artwork reopens with stable anchors and scripts disabled");
            Home_Click(this, new()); await Idle();
            Check(StartPage.IsVisible && ResumeButton.IsVisible && RecentList.Items.Count > 0, "Home retains current document and exposes recent files");
            WorkbookPage_Click(this, new());
            File.WriteAllText(Path.Combine(folder, "product-cases.json"), JsonSerializer.Serialize(new { cases,
                scope = "Actual WPF controls, pages, layout, command and storage paths. No physical IME, screen-reader or clean-machine claims." }, new JsonSerializerOptions { WriteIndented = true }));
            File.WriteAllText(Path.Combine(folder, "product-smoke-result.txt"), $"PASS: {cases.Count} WPF assertions; product shell, scripts, search, artwork, accessibility metadata.");
        }
        finally { _smokeSearchDriver = null; _refreshing = false; _suppressAutoCalculation = false; ApplyPreferences(new()); }
    }

    private void InstallProductFixture()
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        void Put(int r, int c, string value, int style = 0) => cells[new(r, c)] = new(CellInput.Parse(value, style));
        Put(0, 0, "销售明细", 1);
        Put(2, 0, "产品", 1); Put(2, 1, "单价", 1); Put(2, 2, "数量", 1); Put(2, 3, "金额", 1);
        string[] names = ["笔记本", "键盘", "显示器", "扩展坞", "鼠标", "耳机", "摄像头", "支架"];
        int[] prices = [6800, 299, 1899, 399, 159, 499, 329, 199];
        for (var i = 0; i < names.Length; i++)
        {
            Put(i + 3, 0, names[i]); Put(i + 3, 1, prices[i].ToString()); Put(i + 3, 2, (i + 2).ToString());
            Put(i + 3, 3, "=B" + (i + 4) + "*C" + (i + 4));
        }
        var drawing = new DrawingVisual();
        using (var dc = drawing.RenderOpen())
        {
            dc.DrawRoundedRectangle(new SolidColorBrush(Color.FromRgb(22, 99, 75)), null, new Rect(0, 0, 160, 90), 8, 8);
            dc.DrawText(new FormattedText("QuickExcel", System.Globalization.CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                new Typeface("Segoe UI"), 24, Brushes.White, 1), new Point(19, 26));
        }
        var bitmap = new RenderTargetBitmap(160, 90, 96, 96, PixelFormats.Pbgra32); bitmap.Render(drawing);
        var png = new PngBitmapEncoder(); png.Frames.Add(BitmapFrame.Create(bitmap)); using var image = new MemoryStream(); png.Save(image);
        var sheet = new WorksheetSnapshot(SheetId.New(), "销售明细", cells, mergedCells: ["A1:D1"], images: [new SheetImage(image.ToArray(), "png", 2, 6, 3, 2)]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [sheet, new(SheetId.New(), "产品目录"), new(SheetId.New(), "月度汇总")], "quickExcel.functions.register('TOTAL', (price, quantity) => price * quantity);",
            [new(), new() { Bold = true, FillColor = "FFE6F1EB", FontColor = "FF16634B" }]);
        var document = new WorkbookDocument(snapshot, _services); document.Calculate();
        InstallDocument(document, null, "销售台账.xlsq");
    }

    private Color ProductCellPixel(DataGrid grid, int row, int column)
    {
        UpdateLayout();
        var cell = ProductChildren<DataGridCell>(grid).Single(c => c.DataContext is SheetRow r && r.Index == row && c.Column.DisplayIndex == column);
        var root = (FrameworkElement)Content;
        var point = cell.TranslatePoint(new Point(cell.ActualWidth / 2, cell.ActualHeight / 2), root);
        var dpi = VisualTreeHelper.GetDpi(root);
        var bitmap = new RenderTargetBitmap((int)(root.ActualWidth * dpi.DpiScaleX), (int)(root.ActualHeight * dpi.DpiScaleY),
            96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(root);
        var pixel = new byte[4];
        bitmap.CopyPixels(new Int32Rect((int)(point.X * dpi.DpiScaleX), (int)(point.Y * dpi.DpiScaleY), 1, 1), pixel, 4, 0);
        return Color.FromArgb(pixel[3], pixel[2], pixel[1], pixel[0]);
    }

    private static IEnumerable<T> ProductChildren<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) yield return match;
            foreach (var nested in ProductChildren<T>(child)) yield return nested;
        }
    }
}
