namespace QuickExcel.Domain;

public enum SheetVisibility { Visible, Hidden, VeryHidden }
public enum SheetPane { TopLeft, TopRight, BottomLeft, BottomRight }
public sealed record SheetSelection(CellPosition ActiveCell, CellRectangle Range, bool OmitActiveCell = false, bool OmitRange = false)
{
    public static SheetSelection At(CellPosition cell) => new(cell, new(cell.Row, cell.Column, cell.Row, cell.Column));
    internal void Validate()
    {
        WorkbookLimits.ValidatePosition(ActiveCell.Row, ActiveCell.Column);
        WorkbookLimits.ValidatePosition(Range.Top, Range.Left);
        WorkbookLimits.ValidatePosition(Range.Bottom, Range.Right);
        if (Range.Bottom < Range.Top || Range.Right < Range.Left) throw new ArgumentException("选区顺序无效。");
        if (!Range.Contains(ActiveCell.Row, ActiveCell.Column)) throw new ArgumentException("活动格必须位于选区内。");
    }
}

/// <summary>Normal worksheet view. Split panes, multi-window and noncontiguous selections remain separately diagnosed.</summary>
public sealed record WorksheetView
{
    public int FrozenRows { get; init; }
    public int FrozenColumns { get; init; }
    public bool FrozenSplit { get; init; }
    public int Zoom { get; init; } = 100;
    public int? NormalZoom { get; init; }
    public CellPosition TopLeftCell { get; init; }
    public CellPosition PaneTopLeftCell { get; init; }
    public SheetPane ActivePane { get; init; }
    public SheetSelection? TopLeft { get; init; }
    public SheetSelection? TopRight { get; init; }
    public SheetSelection? BottomLeft { get; init; }
    public SheetSelection? BottomRight { get; init; }
    public bool HasFreeze => FrozenRows > 0 || FrozenColumns > 0;
    public SheetSelection? Selection => ActivePane switch
    { SheetPane.TopRight => TopRight, SheetPane.BottomLeft => BottomLeft, SheetPane.BottomRight => BottomRight, _ => TopLeft };
    public CellPosition ActiveCell => Selection?.ActiveCell ?? default;
    public WorksheetView WithActiveCell(CellPosition cell)
    {
        var right = FrozenColumns > 0 && cell.Column >= FrozenColumns;
        var bottom = FrozenRows > 0 && cell.Row >= FrozenRows;
        var pane = (right, bottom) switch { (true, true) => SheetPane.BottomRight, (true, false) => SheetPane.TopRight, (false, true) => SheetPane.BottomLeft, _ => SheetPane.TopLeft };
        var view = this with { ActivePane = pane };
        return pane switch
        {
            SheetPane.TopRight => view with { TopRight = SheetSelection.At(cell) },
            SheetPane.BottomLeft => view with { BottomLeft = SheetSelection.At(cell) },
            SheetPane.BottomRight => view with { BottomRight = SheetSelection.At(cell) },
            _ => view with { TopLeft = SheetSelection.At(cell) }
        };
    }
    internal void Validate()
    {
        if (FrozenRows < 0 || FrozenRows >= WorkbookLimits.MaxRows || FrozenColumns < 0 || FrozenColumns >= WorkbookLimits.MaxColumns ||
            Zoom is < 10 or > 400 || NormalZoom is < 10 or > 400 || !Enum.IsDefined(ActivePane) || FrozenSplit && !HasFreeze)
            throw new ArgumentException("工作表视图参数超出支持范围。");
        WorkbookLimits.ValidatePosition(TopLeftCell.Row, TopLeftCell.Column);
        WorkbookLimits.ValidatePosition(PaneTopLeftCell.Row, PaneTopLeftCell.Column);
        if (HasFreeze && (PaneTopLeftCell.Row < FrozenRows || PaneTopLeftCell.Column < FrozenColumns))
            throw new ArgumentException("可滚动窗格的首格不能位于冻结区域内。");
        if (!HasFreeze && (ActivePane != SheetPane.TopLeft || PaneTopLeftCell != default || TopRight is not null || BottomLeft is not null || BottomRight is not null))
            throw new ArgumentException("非冻结视图不能包含其他窗格。");
        foreach (var selection in new[] { TopLeft, TopRight, BottomLeft, BottomRight }) selection?.Validate();
    }
}

/// <summary>One workbook window; coordinates are persisted Office twips, not desktop window positioning commands.</summary>
public sealed record WorkbookWindow(int? X = null, int? Y = null, uint? Width = null, uint? Height = null);
