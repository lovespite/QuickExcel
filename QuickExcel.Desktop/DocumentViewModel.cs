using System.ComponentModel;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Read-only presentation state; authoritative document state always belongs to the session.</summary>
internal sealed class DocumentViewModel(WorkbookDocument document) : INotifyPropertyChanged
{
    public WorkbookDocument Document { get; } = document;
    public bool IsDirty => Document.IsDirty;
    public bool CanUndo => Document.CanUndo;
    public bool CanRedo => Document.CanRedo;
    public long Revision => Document.Snapshot.Revision;
    public int DiscardedHistoryCount => Document.Session.DiscardedHistoryCount;
    public int StaleCacheCount => Document.Snapshot.Sheets.Sum(sheet => sheet.Cells.Values.Count(cell =>
        cell.Content.Kind == CellContentKind.Formula && cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error)));
    public event PropertyChangedEventHandler? PropertyChanged;
    public void Refresh() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(null));
}
