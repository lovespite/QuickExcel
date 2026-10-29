using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>View-only boundary. All writes remain in EditorController/DocumentSession.</summary>
internal interface IGridSurface
{
    void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns);
    void Refresh(WorkbookSnapshot snapshot);
    IReadOnlyList<CellPosition> Selection { get; }
    IReadOnlyList<GridRange> SelectionRanges { get; }
    void Select(CellPosition position, bool extend = false);
    void SelectRange(GridRange range, CellPosition current);
    bool Commit();
    void Cancel();
    bool BeginEdit();
}
