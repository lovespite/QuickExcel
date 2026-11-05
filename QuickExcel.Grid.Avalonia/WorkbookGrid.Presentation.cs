namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    /// <summary>Refresh time-dependent presentation without changing the snapshot or active edit ownership.</summary>
    public void InvalidatePresentation()
    {
        ClearScene(); InvalidateContent();
        if (!_editing) PrepareEditor(false);
    }
}
