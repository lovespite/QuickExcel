using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow : Window
{
    private readonly WorkbookWorkspace _workspace;
    private readonly WorkbookGrid _grid = new();
    private readonly CompositionTextBox _formula = new() { MinHeight = 30 };
    private readonly TextBlock _address = new() { Width = 86, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _status = new() { Text = "就绪", Margin = new(12, 5) };
    private readonly ListBox _sheets = new() { MinHeight = 36, SelectionMode = SelectionMode.Single };
    private readonly DockPanel _book = new();
    private readonly StackPanel _start = new() { Margin = new(60), Spacing = 18 };
    private readonly StackPanel _recent = new() { Spacing = 6 };
    private readonly Menu _menu = new();
    private readonly RecentFileStore _recentFiles;
    private readonly DesktopSettings _settings;
    private ILocalFileDialogs _files;
    private Guid _boundSession;
    private SheetId _boundSheet;
    private bool _synchronizing, _formulaDirty, _allowClose, _closePending, _working;
    private CellAddress? _formulaOwner;
    private Task _operation = Task.CompletedTask;
    private Func<Task<SaveDecision>>? _discardChoice;
    private WorkbookDocument Document => _workspace.Document;
    private int SheetIndex => Document.Snapshot.Sheets.IndexOf(Document.Snapshot.GetSheet(Document.Snapshot.ActiveSheetId));

    public MainWindow(string userDataDirectory)
    {
        Title = "QuickExcel · Avalonia"; Width = 1280; Height = 820; MinWidth = 760; MinHeight = 480;
        FontFamily = new FontFamily("Segoe UI, Noto Sans CJK SC, sans-serif"); FontSize = 13;
        var services = DesktopServices.Create(); _workspace = new(services, new WorkbookFileAccess(services.Storage));
        _files = new AvaloniaFileDialogs(this);
        _recentFiles = new(Path.Combine(userDataDirectory, "recent-files.json")); _recentFiles.Load();
        _settings = new(Path.Combine(userDataDirectory, "settings.json")); _settings.Load();
        BuildView();
        _grid.ResolveStyle = ConditionalPresentation.GetStyle;
        _grid.CommitCell = (index, owner, text) => TryEdit(() => _workspace.Editor.SetCell(index, owner.Row, owner.Column, text));
        _grid.BeforeSelection = CommitFormula;
        _grid.SelectionChanged += () =>
        {
            if (_synchronizing) return;
            Document.Session.UpdateNavigation(Document.Snapshot.ActiveSheetId, _grid.Current); UpdateFormula(); UpdateTitle();
        };
        _formula.GotFocus += (_, _) => _formulaOwner ??= new(Document.Snapshot.ActiveSheetId, _grid.Current.Row, _grid.Current.Column);
        _formula.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty && !_synchronizing) _formulaDirty = true; };
        _formula.AddHandler(KeyDownEvent, (_, e) =>
        {
            if (_formula.IsComposing) return;
            if (e.Key == Key.Enter) { e.Handled = true; if (CommitFormula()) _grid.Select(_grid.Current); }
            if (e.Key == Key.Escape) { e.Handled = true; _formulaDirty = false; _formulaOwner = null; UpdateFormula(); _grid.Select(_grid.Current); }
        }, RoutingStrategies.Tunnel);
        _sheets.SelectionChanged += (_, _) =>
        {
            if (_synchronizing || _sheets.SelectedItem is not SheetItem selected || selected.Id == Document.Snapshot.ActiveSheetId) return;
            if (!CommitEditors()) { Synchronize(); return; }
            Document.Session.UpdateNavigation(selected.Id, Document.Snapshot.GetSheet(selected.Id).View.ActiveCell); Synchronize();
        };
        _workspace.Changed += Synchronize;
        AddHandler(KeyDownEvent, Shortcuts, RoutingStrategies.Tunnel);
        Closing += OnClosing;
        Closed += (_, _) => { _grid.Dispose(); _workspace.Dispose(); };
        Opened += async (_, _) =>
        {
            if (Program.SmokeDirectory is { } smoke) { await RunWorkspaceSmokeAsync(smoke); return; }
            if (Program.InitialFile is { } path) await PerformAsync(() => OpenPathAsync(path));
        };
        Synchronize(); ShowHome();
    }
    private sealed record SheetItem(SheetId Id, string Name) { public override string ToString() => Name; }
    private void BuildView()
    {
        MenuItem Item(string text, Func<Task> action) { var item = new MenuItem { Header = text }; item.Click += async (_, _) => await PerformAsync(action); return item; }
        _menu.Items.Add(new MenuItem { Header = "文件", Items =
        {
            Item("新建  Ctrl+N", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }),
            Item("打开…  Ctrl+O", OpenAsync), Item("保存  Ctrl+S", async () => await SaveAsync(false)),
            Item("另存为…", async () => await SaveAsync(true)), Item("最近文件", async () => { if (await ConfirmDiscardAsync()) ShowHome(); }),
            Item("关闭工作簿", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }), Item("退出", () => { Close(); return Task.CompletedTask; })
        } });
        _menu.Items.Add(new MenuItem { Header = "编辑", Items =
        { Item("撤销  Ctrl+Z", () => History(false)), Item("重做  Ctrl+Y", () => History(true)), Item("清除内容", ClearAsync) } });
        _menu.Items.Add(new MenuItem { Header = "工作表", Items =
        { Item("添加工作表", AddSheetAsync), Item("重命名…", RenameSheetAsync), Item("删除工作表…", DeleteSheetAsync) } });
        _menu.Items.Add(new MenuItem { Header = "视图", Items =
        { Item("100%", () => ZoomAsync(100)), Item("125%", () => ZoomAsync(125)), Item("150%", () => ZoomAsync(150)),
          Item("标签在底部", () => PlaceSheets(SheetNavigationPlacement.Bottom)), Item("标签在左侧", () => PlaceSheets(SheetNavigationPlacement.Left)), Item("标签在右侧", () => PlaceSheets(SheetNavigationPlacement.Right)) } });
        _menu.Items.Add(Item("重算  F9", CalculateAsync));
        var formulaRow = new DockPanel { Margin = new(8, 4) }; DockPanel.SetDock(_address, Dock.Left); formulaRow.Children.Add(_address); formulaRow.Children.Add(_formula);
        DockPanel.SetDock(formulaRow, Dock.Top); _book.Children.Add(formulaRow); _book.Children.Add(_sheets); _book.Children.Add(_grid); ApplySheetPlacement();
        _start.Children.Add(new TextBlock { Text = "QuickExcel", FontSize = 36, FontWeight = FontWeight.SemiBold });
        _start.Children.Add(new TextBlock { Text = "工作簿与脚本", FontSize = 18, Foreground = Brushes.Gray });
        void StartButton(string text, Func<Task> action) { var button = new Button { Content = text, MinWidth = 220 }; button.Click += async (_, _) => await PerformAsync(action); _start.Children.Add(button); }
        StartButton("新建工作簿", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } });
        StartButton("打开工作簿…", OpenAsync);
        StartButton("打开示例", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(true); ShowBook(); } });
        StartButton("返回当前工作簿", () => { ShowBook(); return Task.CompletedTask; });
        _start.Children.Add(new TextBlock { Text = "最近文件", FontSize = 18 }); _start.Children.Add(_recent);
        var content = new global::Avalonia.Controls.Grid(); content.Children.Add(_book); content.Children.Add(new ScrollViewer { Content = _start });
        var root = new DockPanel(); DockPanel.SetDock(_menu, Dock.Top); DockPanel.SetDock(_status, Dock.Bottom);
        root.Children.Add(_menu); root.Children.Add(_status); root.Children.Add(content); Content = root;
    }
    private void ShowHome()
    {
        _book.IsVisible = false; ((Control)_start.Parent!).IsVisible = true; _recent.Children.Clear();
        foreach (var path in _recentFiles.Paths)
        {
            var button = new Button { Content = Path.GetFileName(path), HorizontalAlignment = HorizontalAlignment.Left }; ToolTip.SetTip(button, path);
            button.Click += async (_, _) => await PerformAsync(async () => { if (await ConfirmDiscardAsync()) await OpenPathAsync(path); }); _recent.Children.Add(button);
        }
    }
    private void ShowBook() { ((Control)_start.Parent!).IsVisible = false; _book.IsVisible = true; }
    private void Synchronize()
    {
        if (_synchronizing) return;
        _synchronizing = true;
        try
        {
            _book.IsEnabled = !_workspace.IsBusy; _menu.IsEnabled = !_workspace.IsBusy;
            if (_boundSession != Document.Session.SessionId || _boundSheet != Document.Snapshot.ActiveSheetId)
            {
                _formulaDirty = false; _formulaOwner = null;
                _boundSession = Document.Session.SessionId; _boundSheet = Document.Snapshot.ActiveSheetId;
                _grid.Bind(Document.Snapshot, SheetIndex, WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns);
            }
            else _grid.Refresh(Document.Snapshot);
            var items = Document.Snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Visible).Select(s => new SheetItem(s.Id, s.Name)).ToArray();
            _sheets.ItemsSource = items; _sheets.SelectedItem = items.FirstOrDefault(s => s.Id == Document.Snapshot.ActiveSheetId);
        }
        finally { _synchronizing = false; }
        UpdateFormula(); UpdateTitle();
    }
    private void UpdateTitle()
    {
        Title = (Document.IsDirty ? "* " : "") + (_workspace.FilePath is { } path ? Path.GetFileName(path) : "未命名.xlsq") + " — QuickExcel";
        _address.Text = WorkbookGrid.ColumnName(_grid.Current.Column) + (_grid.Current.Row + 1);
    }
    private void UpdateFormula()
    {
        if (_formulaDirty || _formula.IsComposing) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        _synchronizing = true; _formula.Text = ProtectionSafety.HidesFormula(Document.Snapshot, sheet, _grid.Current.Row, _grid.Current.Column) ? "" : Document.GetCellInput(SheetIndex, _grid.Current.Row, _grid.Current.Column); _synchronizing = false;
        _formulaOwner = null; _formula.IsReadOnly = Document.Snapshot.FileState?.FormulaReadOnly == true || !ProtectionSafety.CanEdit(Document.Snapshot, sheet, _grid.Current.Row, _grid.Current.Column);
    }
    private bool CommitFormula()
    {
        if (_formula.IsComposing) return false;
        if (!_formulaDirty) return true;
        var owner = _formulaOwner ?? new(Document.Snapshot.ActiveSheetId, _grid.Current.Row, _grid.Current.Column);
        var index = Document.Snapshot.Sheets.IndexOf(Document.Snapshot.GetSheet(owner.SheetId));
        if (!TryEdit(() => _workspace.Editor.SetCell(index, owner.Row, owner.Column, _formula.Text ?? ""))) return false;
        _formulaDirty = false; _formulaOwner = null; UpdateFormula(); return true;
    }
    private bool CommitEditors() => !_workspace.IsBusy && _grid.Commit() && CommitFormula();
    private bool TryEdit(Func<CommandResult> edit)
    {
        if (_workspace.IsBusy) return false;
        try { edit(); Synchronize(); _status.Text = "已修改"; return true; }
        catch (Exception ex) { Error(ex); return false; }
    }
    private void Error(Exception ex) => _status.Text = ex.Message;
    private async Task PerformAsync(Func<Task> action)
    {
        if (_working) return;
        _working = true;
        try { _operation = action(); await _operation; }
        catch (Exception ex) { Error(ex); }
        finally { _working = false; }
    }
    private async Task<bool> ConfirmDiscardAsync()
    {
        if (!CommitEditors()) return false;
        if (!Document.IsDirty) return true;
        var choice = _discardChoice is null ? await Dialogs.Unsaved(this) : await _discardChoice();
        return choice == SaveDecision.Discard || choice == SaveDecision.Save && await SaveAsync(false);
    }
    private async Task OpenAsync()
    {
        if (!CommitEditors()) return;
        var selected = await _files.OpenAsync(new("打开工作簿", [new("工作簿", ["*.xlsx", "*.xlsq"])]));
        if (selected is not null && await ConfirmDiscardAsync()) await OpenPathAsync(selected.Path);
    }
    private async Task OpenPathAsync(string path)
    {
        await _workspace.OpenAsync(new(path)); _recentFiles.Remember(path); ShowBook(); _status.Text = "已打开 " + Path.GetFileName(path);
    }
    private async Task<bool> SaveAsync(bool saveAs)
    {
        if (!CommitEditors()) return false;
        var snapshot = Document.Snapshot;
        var issues = TypedWorkbookStorage.SaveIssues(snapshot); var discard = false;
        if (issues.Count > 0)
        {
            if (!await Dialogs.Confirm(this, "保存受支持副本", "以下内容不能原样保存，确定放弃这些内容并另存副本？\n\n" + string.Join("\n", issues))) return false;
            discard = true; saveAs = true;
        }
        var migration = snapshot.FileState?.RequiresMigration == true;
        var needsScripts = TypedWorkbookStorage.RequiresScriptContainer(snapshot);
        var current = _workspace.FilePath;
        if (migration || needsScripts && current is not null && Path.GetExtension(current).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) saveAs = true;
        LocalFileSelection? selected = !saveAs && current is not null ? new(current) : null;
        if (selected is null)
        {
            var extension = needsScripts || migration ? ".xlsq" : ".xlsx";
            FileDialogType[] types = needsScripts ? [new("QuickExcel 工作簿（含脚本）", ["*.xlsq"])] : [new("Excel 工作簿", ["*.xlsx"]), new("QuickExcel 工作簿", ["*.xlsq"])];
            selected = await _files.SaveAsync(new("保存工作簿", types, Path.GetFileNameWithoutExtension(current ?? "未命名") + (migration ? "-v2" : discard ? "-受支持副本" : "") + extension, extension));
        }
        if (selected is null) return false;
        try { await _workspace.SaveAsync(selected, discard); _recentFiles.Remember(selected.Path); _status.Text = "已保存 " + Path.GetFileName(selected.Path); return true; }
        catch (Exception ex) { Error(ex); return false; }
    }
    private Task History(bool redo) { if (CommitEditors()) TryEdit(() => redo ? _workspace.Editor.Redo() : _workspace.Editor.Undo()); return Task.CompletedTask; }
    private Task ClearAsync()
    {
        if (CommitEditors()) TryEdit(() => _workspace.Editor.ClearCells(SheetIndex, _grid.Selection.Select(p => (p.Row, p.Column))));
        return Task.CompletedTask;
    }
    private Task AddSheetAsync()
    {
        if (!CommitEditors()) return Task.CompletedTask;
        var sheet = Document.AddSheet(); Document.Session.UpdateNavigation(sheet.Id, new(0, 0)); Synchronize(); return Task.CompletedTask;
    }
    private async Task RenameSheetAsync()
    {
        if (!CommitEditors()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex]; var name = await Dialogs.Text(this, "重命名工作表", "工作表名称", sheet.Name);
        if (name is not null) TryEdit(() => _workspace.Editor.Execute(new RenameSheetCommand(sheet.Id, name)));
    }
    private async Task DeleteSheetAsync()
    {
        if (!CommitEditors()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        if (await Dialogs.Confirm(this, "删除工作表", $"删除“{sheet.Name}”？引用此表的公式将变为 #REF!，可撤销。")) TryEdit(() => _workspace.Editor.Execute(new DeleteSheetCommand(sheet.Id)));
    }
    private Task ZoomAsync(int value)
    {
        if (CommitEditors()) TryEdit(() => _workspace.Editor.Execute(new SetSheetViewCommand(Document.Snapshot.ActiveSheetId, Document.Snapshot.Sheets[SheetIndex].View with { Zoom = value })));
        return Task.CompletedTask;
    }
    private Task PlaceSheets(SheetNavigationPlacement placement) { _settings.Save(_settings.Value with { Sheets = placement }); ApplySheetPlacement(); return Task.CompletedTask; }
    private void ApplySheetPlacement()
    {
        var placement = _settings.Value.Sheets; var vertical = placement != SheetNavigationPlacement.Bottom;
        DockPanel.SetDock(_sheets, placement switch { SheetNavigationPlacement.Left => Dock.Left, SheetNavigationPlacement.Right => Dock.Right, _ => Dock.Bottom });
        _sheets.Width = vertical ? 160 : double.NaN;
        _sheets.ItemsPanel = new FuncTemplate<Panel?>(() => new StackPanel { Orientation = vertical ? Orientation.Vertical : Orientation.Horizontal });
    }
    private async Task CalculateAsync()
    {
        if (!CommitEditors()) return;
        await _workspace.CalculateAsync(); _status.Text = "重算完成";
    }
    private async void Shortcuts(object? sender, KeyEventArgs e)
    {
        if (_grid.IsComposing || _formula.IsComposing) return;
        var command = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        Func<Task>? action = command ? e.Key switch { Key.S => async () => await SaveAsync(e.KeyModifiers.HasFlag(KeyModifiers.Shift)), Key.O => OpenAsync,
            Key.N => async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }, Key.Z => () => History(false), Key.Y => () => History(true), _ => null } : e.Key == Key.F9 ? CalculateAsync : null;
        if (action is null) return; e.Handled = true; await PerformAsync(action);
    }
    private async void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_allowClose) return;
        e.Cancel = true; if (_closePending) return; _closePending = true;
        try
        {
            await _operation;
            if (!await ConfirmDiscardAsync()) return;
            _allowClose = true; Close();
        }
        catch (Exception ex) { Error(ex); }
        finally { _closePending = false; }
    }
}
