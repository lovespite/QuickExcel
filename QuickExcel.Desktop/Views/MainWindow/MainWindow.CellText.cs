using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private void InstallCellTextHandlers()
    {
        var menu = new ContextMenu();
        var flatten = new MenuItem { Header = "转换当前格为普通文本（移除分段格式）" };
        flatten.Click += (_, _) => FlattenCurrentRichText();
        menu.Items.Add(flatten);
        menu.Opened += (_, _) => flatten.IsEnabled = !_busy && _document.Snapshot.FileState?.FormulaReadOnly != true &&
            _document.Snapshot.Sheets[_sheetIndex].GetCell(_selectedRow, _selectedColumn)?.Content.RichText is not null;
        _grid.Control.ContextMenu = menu;
    }
    private void FlattenCurrentRichText()
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            _editor.Execute(new FlattenRichTextCommand([new(sheet.Id, _selectedRow, _selectedColumn)]));
            RefreshSheet(); StatusText.Text = "已转换为普通文本；分段格式可通过撤销恢复。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
