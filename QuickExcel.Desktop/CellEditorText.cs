using System.Windows.Controls;

namespace QuickExcel.Desktop;

internal static class CellEditorText
{
    internal static void InitializeText(TextBox editor, string text)
    {
        // Display changes must never enter the user's current-cell undo stack.
        editor.IsUndoEnabled = false;
        editor.Text = text;
        editor.IsUndoEnabled = true;
    }
}
