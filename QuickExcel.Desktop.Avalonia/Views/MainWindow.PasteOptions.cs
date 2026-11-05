using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Primitives.PopupPositioning;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed record LastPaste(Guid Session, SheetId Sheet, long Revision, WorkbookSnapshot Before,
        ClipboardExchange Data, WorkbookClipboard? Cells, GridRange Target, GridRange Output, string Mode, bool HasFormats);
    private LastPaste? _lastPaste;
    private Popup? _pasteIconPopup, _pasteOptionsPopup;
    private readonly Button _pasteOptionsButton = new() { Padding = new(5, 3), MinHeight = 0, MinWidth = 0 };
    private WorkbookSnapshot? _pasteReplacementBase;
    private bool _cellPasteApplied, _pasteHasFormats;
    private long _cellPasteRevision;
    private WorkbookSnapshot PasteSnapshot => _pasteReplacementBase ?? Document.Snapshot;

    private void InstallPasteOptions()
    {
        var icon = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };
        icon.Children.Add(RibbonIcons.Create(RibbonIcon.Paste));
        icon.Children.Add(new TextBlock { Text = "▾", VerticalAlignment = VerticalAlignment.Center });
        _pasteOptionsButton.Content = icon;
        _pasteOptionsButton.Background = new SolidColorBrush(Color.Parse("#EDF6F0"));
        _pasteOptionsButton.BorderBrush = new SolidColorBrush(Color.Parse("#81B599"));
        _pasteOptionsButton.BorderThickness = new(1);
        AutomationProperties.SetName(_pasteOptionsButton, UiText.Get("粘贴选项"));
        ToolTip.SetTip(_pasteOptionsButton, UiText.Get("粘贴选项"));
        _pasteOptionsButton.Click += (_, _) => TogglePasteOptions();
        _grid.DraftStarted += DismissPasteOptions;
        _grid.SelectionChanged += DismissPasteOptions;
        _grid.ViewportChanged += PositionPasteOptions;
        _grid.SizeChanged += (_, _) => PositionPasteOptions();
        _formula.GotFocus += (_, _) => DismissPasteOptions();
        _formula.PropertyChanged += (_, e) =>
        { if (e.Property == TextBox.TextProperty && !_synchronizing) DismissPasteOptions(); };
        _formula.CompositionChanged += () => { if (_formula.IsComposing) DismissPasteOptions(); };
        Closed += (_, _) => DismissPasteOptions();
    }

    private async Task<bool> ExecuteCellPasteAsync(DocumentCommand command, long revision)
    {
        if (!await CommitEditorsAsync()) return false;
        try
        {
            var result = await _workspace.Editor.ExecuteCellPasteAsync(command, revision, _pasteReplacementBase is not null, ConfirmValidationAsync);
            _cellPasteApplied |= result.Changed;
            if (result.Changed) _cellPasteRevision = result.Revision;
        }
        catch (DataValidationException ex) when (ex.CanOverride) { Error(ex); return false; }
        Synchronize(); return true;
    }

    private bool IsPasteOptionsCurrent() => _lastPaste is { } paste && paste.Session == Document.Session.SessionId &&
        paste.Sheet == Document.Snapshot.ActiveSheetId && paste.Revision == Document.Snapshot.Revision &&
        paste.Revision > Document.Session.SavedRevision && !_grid.IsEditing && !_formulaDirty && _grid.SelectedDrawingObject is null;

    private void ShowPasteOptions(LastPaste paste)
    {
        DismissPasteOptions(); _lastPaste = paste;
        _pasteIconPopup = new Popup
        {
            Child = _pasteOptionsButton, PlacementTarget = _grid, Placement = PlacementMode.AnchorAndGravity,
            PlacementAnchor = PopupAnchor.TopLeft, PlacementGravity = PopupGravity.BottomRight,
            IsLightDismissEnabled = false
        };
        _book.Children.Insert(0, _pasteIconPopup);
        PositionPasteOptions();
    }

    private void PositionPasteOptions()
    {
        if (_pasteIconPopup is not { } popup) return;
        if (!IsPasteOptionsCurrent()) { DismissPasteOptions(); return; }
        if (_grid.GetVisibleRangeBounds(_lastPaste!.Output) is not { } bounds)
        { ClosePasteOptionsPanel(); popup.IsOpen = false; return; }
        var x = Math.Clamp(bounds.Right + 4, 0, Math.Max(0, _grid.Bounds.Width - 64));
        var y = Math.Clamp(bounds.Bottom + 4, 0, Math.Max(0, _grid.Bounds.Height - 48));
        popup.PlacementRect = new Rect(x, y, 1, 1);
        popup.IsOpen = true;
    }

    private void TogglePasteOptions()
    {
        if (!IsPasteOptionsCurrent()) { DismissPasteOptions(); return; }
        if (_pasteOptionsPopup is not null) { ClosePasteOptionsPanel(); return; }
        CloseFormulaHelp(); CloseFilterPanel(); CloseChartFilter();
        var paste = _lastPaste!;
        var panel = new PasteOptionsPanel(paste.Mode, paste.HasFormats);
        panel.Selected += mode => _ = PerformAsync(() => ChangePasteModeAsync(mode));
        var popup = new Popup
        {
            Child = panel, PlacementTarget = _pasteOptionsButton, Placement = PlacementMode.BottomEdgeAlignedLeft,
            IsLightDismissEnabled = true
        };
        popup.Closed += (_, _) =>
        {
            if (_pasteOptionsPopup == popup) _pasteOptionsPopup = null;
            popup.Child = null; _book.Children.Remove(popup);
        };
        panel.KeyDown += (_, e) =>
        { if (e.Key == Key.Escape) { e.Handled = true; DismissPasteOptions(); _grid.Focus(); } };
        _pasteOptionsPopup = popup; _book.Children.Insert(0, popup); popup.IsOpen = true;
        panel.FocusFirst();
    }

    private async Task ChangePasteModeAsync(string mode)
    {
        if (!IsPasteOptionsCurrent()) { DismissPasteOptions(); return; }
        var paste = _lastPaste!;
        DismissPasteOptions();
        if (mode == paste.Mode) { _grid.Focus(); return; }
        _pasteReplacementBase = paste.Before;
        try
        {
            await PasteCapturedAsync(paste.Data, paste.Cells, mode, paste.Target, paste.Sheet, paste.Revision);
        }
        finally { _pasteReplacementBase = null; _grid.Focus(); }
    }

    private void ClosePasteOptionsPanel()
    {
        var popup = _pasteOptionsPopup; _pasteOptionsPopup = null;
        if (popup is null) return;
        popup.IsOpen = false; popup.Child = null; _book.Children.Remove(popup);
    }

    private void DismissPasteOptions()
    {
        _lastPaste = null; ClosePasteOptionsPanel();
        var popup = _pasteIconPopup; _pasteIconPopup = null;
        if (popup is null) return;
        popup.IsOpen = false; popup.Child = null; _book.Children.Remove(popup);
    }
}
