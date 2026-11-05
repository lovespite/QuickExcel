using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly DockPanel _sheetStrip = new();
    private readonly DockPanel _sheetTabsAndAdd = new();
    private readonly StackPanel _sheetTools = new() { Orientation = Orientation.Horizontal };
    private readonly global::Avalonia.Controls.Grid _sheetTabArea = new() { ClipToBounds = true };
    private readonly Canvas _sheetDropOverlay = new() { IsHitTestVisible = false };
    private readonly Border _sheetDropMarker = new() { Background = Brushes.SeaGreen, IsVisible = false };
    private readonly Button _addSheetButton = TabButton(UiText.Get("新建工作表"), "+");
    private readonly Button _scrollSheetsBack = TabButton(UiText.Get("向前滚动标签"), "‹");
    private readonly Button _scrollSheetsForward = TabButton(UiText.Get("向后滚动标签"), "›");
    private readonly DispatcherTimer _sheetDragScroll = new() { Interval = TimeSpan.FromMilliseconds(100) };
    private readonly Cursor _moveSheetCursor = new(StandardCursorType.DragMove);
    private readonly Cursor _copySheetCursor = new(StandardCursorType.DragCopy);
    private SheetPress? _sheetPress;
    private bool _draggingSheet;
    private Point _sheetPointerPosition;
    private bool VerticalTabs => _settings.Value.Sheets != SheetNavigationPlacement.Bottom;
    private sealed record SheetPress(IPointer Pointer, SheetId Id, Point Start, Guid Session, long Revision);

    private static Button TabButton(string title, object icon)
    {
        var button = new Button { Content = icon, Width = 30, Height = 28, MinHeight = 0, Padding = new(0),
            FontSize = 20, HorizontalContentAlignment = HorizontalAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center,
            Background = Brushes.Transparent, BorderThickness = new(0) };
        ToolTip.SetTip(button, title); AutomationProperties.SetName(button, title);
        return button;
    }

    private DockPanel BuildSheetStrip()
    {
        // A table with an inset gear in its lower-right corner.
        var icon = new PathIcon { Width = 19, Height = 19, Data = Geometry.Parse(
            "M1 2 H17 V8 H15.5 V6.5 H10 V9.5 H8.5 V6.5 H2.5 V10 H7 V11.5 H2.5 V15 H7 V16.5 H1 Z " +
            "M2.5 3.5 V5 H8.5 V3.5 Z M10 3.5 V5 H15.5 V3.5 Z " +
            "M13 9 H15 L15.4 10.4 L16.5 10.9 L17.9 10.4 L19 12.2 L18 13.3 V14.6 L19 15.7 L18 17.5 L16.5 17.1 L15.4 17.7 L15 19 H13 L12.6 17.7 L11.5 17.1 L10 17.5 L9 15.7 L10 14.6 V13.3 L9 12.2 L10.1 10.4 L11.5 10.9 L12.6 10.4 Z " +
            "M16.2 14 A2.2 2.2 0 1 0 11.8 14 A2.2 2.2 0 1 0 16.2 14 Z") };
        var manager = TabButton(UiText.Get("工作表管理器"), icon);
        manager.Click += async (_, _) => await PerformAsync(ShowSheetManagerAsync);
        _scrollSheetsBack.Click += (_, _) => ScrollSheetTabs(-1);
        _scrollSheetsForward.Click += (_, _) => ScrollSheetTabs(1);
        _addSheetButton.Click += async (_, _) => await PerformAsync(AddSheetAsync);
        _sheetTools.Children.Add(manager); _sheetTools.Children.Add(_scrollSheetsBack); _sheetTools.Children.Add(_scrollSheetsForward);
        _sheetDropOverlay.Children.Add(_sheetDropMarker);
        _sheetTabArea.Children.Add(_sheets); _sheetTabArea.Children.Add(_sheetDropOverlay);
        _sheetTabsAndAdd.Children.Add(_addSheetButton); _sheetTabsAndAdd.Children.Add(_sheetTabArea);
        _sheetStrip.Children.Add(_sheetTools); _sheetStrip.Children.Add(_sheetTabsAndAdd);
        _sheets.AddHandler(ScrollViewer.ScrollChangedEvent, (_, _) => UpdateSheetArrows());
        _sheetStrip.AddHandler(PointerPressedEvent, SheetTabPressed, RoutingStrategies.Tunnel);
        _sheetStrip.AddHandler(PointerMovedEvent, SheetTabMoved, RoutingStrategies.Tunnel);
        _sheetStrip.AddHandler(PointerReleasedEvent, SheetTabReleased, RoutingStrategies.Tunnel);
        _sheetStrip.PointerCaptureLost += (_, e) => { if (e.Source == _sheetStrip) EndSheetGesture(); };
        Deactivated += (_, _) => EndSheetGesture();
        _sheetDragScroll.Tick += (_, _) =>
        {
            var direction = SheetEdgeDirection();
            if (direction == 0) { _sheetDragScroll.Stop(); return; }
            ScrollSheetTabs(direction, 28);
            UpdateSheetDropMarker();
        };
        return _sheetStrip;
    }

    private void ApplySheetPlacement()
    {
        EndSheetGesture();
        var placement = _settings.Value.Sheets;
        var vertical = VerticalTabs;
        DockPanel.SetDock(_sheetStrip, placement switch { SheetNavigationPlacement.Left => Dock.Left, SheetNavigationPlacement.Right => Dock.Right, _ => Dock.Bottom });
        _sheetStrip.Width = vertical ? 160 : double.NaN;
        DockPanel.SetDock(_sheetTools, vertical ? Dock.Top : Dock.Left);
        DockPanel.SetDock(_addSheetButton, vertical ? Dock.Bottom : Dock.Right);
        _sheetTabsAndAdd.HorizontalAlignment = vertical ? HorizontalAlignment.Stretch : HorizontalAlignment.Left;
        _sheetTabsAndAdd.VerticalAlignment = vertical ? VerticalAlignment.Top : VerticalAlignment.Stretch;
        _sheets.VerticalAlignment = vertical ? VerticalAlignment.Stretch : VerticalAlignment.Center;
        _scrollSheetsBack.Content = vertical ? "⌃" : "‹"; _scrollSheetsForward.Content = vertical ? "⌄" : "›";
        ScrollViewer.SetHorizontalScrollBarVisibility(_sheets, vertical ? ScrollBarVisibility.Disabled : ScrollBarVisibility.Hidden);
        ScrollViewer.SetVerticalScrollBarVisibility(_sheets, vertical ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Disabled);
        _sheets.ItemsPanel = new FuncTemplate<Panel?>(() => new StackPanel { Orientation = vertical ? Orientation.Vertical : Orientation.Horizontal });
        UpdateSheetArrows();
    }

    private void UpdateSheetArrows()
    {
        var scroll = _sheets.Scroll;
        var offset = VerticalTabs ? scroll?.Offset.Y : scroll?.Offset.X;
        var maximum = VerticalTabs ? scroll?.Extent.Height - scroll?.Viewport.Height : scroll?.Extent.Width - scroll?.Viewport.Width;
        _scrollSheetsBack.IsEnabled = offset > 0.5;
        _scrollSheetsForward.IsEnabled = offset < maximum - 0.5;
    }

    private void ScrollSheetTabs(int direction, double? distance = null)
    {
        if (_sheets.Scroll is not { } scroll) return;
        var size = VerticalTabs ? scroll.Viewport.Height : scroll.Viewport.Width;
        var step = direction * (distance ?? Math.Max(32, size * 0.65));
        scroll.Offset += VerticalTabs ? new Vector(0, step) : new Vector(step, 0);
    }

    private void SheetTabPressed(object? sender, PointerPressedEventArgs e)
    {
        if (Document.IsReadOnly) return; // Let the ListBox perform ordinary sheet navigation without a drag gesture.
        if (e.Pointer.Type != PointerType.Mouse || e.GetCurrentPoint(_sheetStrip).Properties.PointerUpdateKind != PointerUpdateKind.LeftButtonPressed ||
            _working || _workspace.IsBusy || _execution.IsRunning || e.Source is not Visual visual) return;
        var container = visual as ListBoxItem ?? visual.FindAncestorOfType<ListBoxItem>();
        if (container?.DataContext is not SheetItem sheet) return;
        e.Handled = true;
        _sheetPress = new(e.Pointer, sheet.Id, e.GetPosition(_sheetStrip), Document.Session.SessionId, Document.Snapshot.Revision);
        e.Pointer.Capture(_sheetStrip);
    }

    private void SheetTabMoved(object? sender, PointerEventArgs e)
    {
        if (_sheetPress is not { } press || press.Pointer != e.Pointer) return;
        _sheetPointerPosition = e.GetPosition(_sheetStrip);
        if (!_draggingSheet && Math.Abs(_sheetPointerPosition.X - press.Start.X) < 5 && Math.Abs(_sheetPointerPosition.Y - press.Start.Y) < 5) return;
        _draggingSheet = true; e.Handled = true;
        _sheetStrip.Cursor = e.KeyModifiers.HasFlag(KeyModifiers.Control) ? _copySheetCursor : _moveSheetCursor;
        UpdateSheetDropMarker();
        if (SheetEdgeDirection() != 0) _sheetDragScroll.Start(); else _sheetDragScroll.Stop();
    }

    private int SheetEdgeDirection()
    {
        if (_sheetStrip.TranslatePoint(_sheetPointerPosition, _sheetTabArea) is not { } p || !new Rect(_sheetTabArea.Bounds.Size).Contains(p)) return 0;
        var axis = VerticalTabs ? p.Y : p.X;
        var size = VerticalTabs ? _sheetTabArea.Bounds.Height : _sheetTabArea.Bounds.Width;
        if (axis < 18 && _scrollSheetsBack.IsEnabled) return -1;
        return axis > size - 18 && _scrollSheetsForward.IsEnabled ? 1 : 0;
    }

    // Return an insertion boundary in workbook order, including hidden sheets.
    private (int Index, double Marker)? SheetDropAt(Point point)
    {
        if (!new Rect(_sheetStrip.Bounds.Size).Contains(point) || _sheetTools.Bounds.Contains(point) ||
            _sheetStrip.TranslatePoint(point, _sheetTabArea) is not { } position) return null;
        var snapshot = Document.Snapshot;
        var axis = VerticalTabs ? position.Y : position.X;
        var extent = VerticalTabs ? _sheetTabArea.Bounds.Height : _sheetTabArea.Bounds.Width;
        if (axis >= extent) return (snapshot.Sheets.Length, extent);
        double end = 0;
        for (var i = 0; i < _sheets.ItemCount; i++)
        {
            if (_sheets.ContainerFromIndex(i) is not ListBoxItem { DataContext: SheetItem item } tab ||
                tab.TranslatePoint(default, _sheetTabArea) is not { } origin) continue;
            var start = VerticalTabs ? origin.Y : origin.X;
            var length = VerticalTabs ? tab.Bounds.Height : tab.Bounds.Width;
            end = start + length;
            if (axis >= end) continue;
            var before = axis < start + length / 2;
            var index = snapshot.Sheets.IndexOf(snapshot.GetSheet(item.Id));
            return (before ? index : index + 1, before ? start : end);
        }
        return (snapshot.Sheets.Length, end);
    }

    private void UpdateSheetDropMarker()
    {
        var drop = SheetDropAt(_sheetPointerPosition);
        _sheetDropMarker.IsVisible = drop.HasValue;
        if (drop is not { } target) return;
        var size = _sheetTabArea.Bounds.Size;
        var marker = Math.Clamp(target.Marker - 1, 0, Math.Max(0, (VerticalTabs ? size.Height : size.Width) - 2));
        Canvas.SetLeft(_sheetDropMarker, VerticalTabs ? 0 : marker); Canvas.SetTop(_sheetDropMarker, VerticalTabs ? marker : 0);
        _sheetDropMarker.Width = VerticalTabs ? size.Width : 2; _sheetDropMarker.Height = VerticalTabs ? 2 : size.Height;
    }

    private async void SheetTabReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (_sheetPress is not { } press || press.Pointer != e.Pointer || e.InitialPressMouseButton != MouseButton.Left) return;
        var point = e.GetPosition(_sheetStrip);
        var dragged = _draggingSheet;
        var drop = IsCurrentSheetPress(press) ? SheetDropAt(point) : null;
        EndSheetGesture(); e.Handled = true;
        if (drop is null) return;
        await PerformAsync(async () =>
        {
            if (!await CommitEditorsAsync()) return;
            var snapshot = Document.Snapshot;
            var source = snapshot.GetSheet(press.Id);
            if (!dragged)
            {
                Document.Session.UpdateNavigation(source.Id, source.View.ActiveCell); Synchronize(); _sheets.Focus();
            }
            else if (e.KeyModifiers.HasFlag(KeyModifiers.Control))
                await ExecuteAsync(new CopySheetCommand(source.Id, NextCopyName(snapshot, source.Name), drop.Value.Index), snapshot.Revision);
            else
            {
                var from = snapshot.Sheets.IndexOf(source);
                var to = drop.Value.Index > from ? drop.Value.Index - 1 : drop.Value.Index;
                if (to != from) await ExecuteAsync(new ReorderSheetCommand(source.Id, to), snapshot.Revision);
            }
        });
    }

    private bool IsCurrentSheetPress(SheetPress press) => press.Session == Document.Session.SessionId && press.Revision == Document.Snapshot.Revision;

    private void EndSheetGesture()
    {
        var pointer = _sheetPress?.Pointer;
        _sheetPress = null; _draggingSheet = false; _sheetDragScroll.Stop();
        _sheetDropMarker.IsVisible = false; _sheetStrip.ClearValue(CursorProperty);
        if (pointer?.Captured == _sheetStrip) pointer.Capture(null);
    }
}
