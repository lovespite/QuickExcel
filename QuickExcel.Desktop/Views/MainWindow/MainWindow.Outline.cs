using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private sealed record OutlineItem(OutlineGroup Group, string Label);
    private SheetAxis? _smokeOutlineAxis;
    private SubtotalPreview? _smokeSubtotalPreview;
    private async void GroupOutline_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = GroupOutlineAsync(false));
    private async void UngroupOutline_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = GroupOutlineAsync(true));
    private async Task GroupOutlineAsync(bool ungroup)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var selected = SelectedRange(); SheetAxis axis;
            if (_smokeMode) axis = _smokeOutlineAxis ?? SheetAxis.Row;
            else
            {
                var picker = new ComboBox { ItemsSource = new[] { $"行 {selected.Top + 1}–{selected.Bottom + 1}", $"列 {ColumnName(selected.Left)}–{ColumnName(selected.Right)}" }, SelectedIndex = 0, Margin = new(0, 8, 0, 12) };
                var panel = new StackPanel { Margin = new(20) }; panel.Children.Add(new TextBlock { Text = ungroup ? "取消选区的一层分组，不删除数据。" : "对选区增加一层分组，最多 7 层。" }); panel.Children.Add(picker);
                var ok = new Button { Content = "确定", IsDefault = true, MinWidth = 90, HorizontalAlignment = HorizontalAlignment.Right }; panel.Children.Add(ok);
                var dialog = new Window { Owner = this, Title = ungroup ? "取消分组" : "分组", Content = panel, SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner, ResizeMode = ResizeMode.NoResize };
                ok.Click += (_, _) => dialog.DialogResult = true; if (dialog.ShowDialog() != true) return; axis = (SheetAxis)picker.SelectedIndex;
            }
            ExecuteWorkbookCommand(new GroupAxisCommand(_document.Snapshot.Sheets[_sheetIndex].Id, axis, axis == SheetAxis.Row ? selected.Top : selected.Left,
                axis == SheetAxis.Row ? selected.Rows : selected.Columns, ungroup));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private void RefreshOutlineView()
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var selected = (OutlineGroups.SelectedItem as OutlineItem)?.Group;
        var groups = WorksheetOutline.Groups(sheet, SheetAxis.Row).Concat(WorksheetOutline.Groups(sheet, SheetAxis.Column)).ToArray();
        var items = groups.Take(2048).Select(g => new OutlineItem(g, $"{(g.Collapsed ? "+" : "−")} {(g.Axis == SheetAxis.Row ? $"行 {g.First + 1}–{g.Last + 1}" : $"列 {ColumnName(g.First)}–{ColumnName(g.Last)}")} · {g.Level} 级")).ToArray();
        OutlineGroups.ItemsSource = items; OutlineGroups.SelectedItem = items.FirstOrDefault(i => i.Group.Axis == selected?.Axis && i.Group.First == selected.First && i.Group.Last == selected.Last && i.Group.Level == selected.Level) ?? items.FirstOrDefault();
        OutlineHint.Text = groups.Length > 2048 ? $"共 {groups.Length} 组（列出前 2048）" : $"{groups.Length} 个分组";
        OutlinePanel.Visibility = groups.Length > 0 && sheet.Data.Outline?.ShowSymbols != false ? Visibility.Visible : Visibility.Collapsed;
    }
    private async void CollapseOutline_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = ChangeOutlineAsync(true));
    private async void ExpandOutline_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = ChangeOutlineAsync(false));
    private async Task ChangeOutlineAsync(bool collapsed)
    {
        if (_busy || !await FlushEditsAsync() || OutlineGroups.SelectedItem is not OutlineItem item) return;
        try
        {
            var group = item.Group;
            ExecuteWorkbookCommand(new SetOutlineCollapsedCommand(_document.Snapshot.Sheets[_sheetIndex].Id, group.Axis, group.First, group.Last, group.Level, collapsed));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async void Subtotal_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = SubtotalAsync());
    private async Task SubtotalAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            SubtotalPreview? preview;
            if (_smokeMode) preview = _smokeSubtotalPreview;
            else
            {
                var selected = SelectedRange(); var book = _document.Snapshot;
                var dialog = new SubtotalDialog(this, book, book.Sheets[_sheetIndex].Id, new(selected.Top, selected.Left, selected.Bottom, selected.Right));
                if (dialog.ShowDialog() != true) return; preview = dialog.Preview;
            }
            if (preview is not null && ExecuteWorkbookCommand(new ApplySubtotalsCommand(preview))) StatusText.Text = $"已插入 {preview.InsertedRows} 行分类汇总；可一次撤销。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
