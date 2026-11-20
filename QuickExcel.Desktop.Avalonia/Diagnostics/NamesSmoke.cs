using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckNamesAsync(string directory, Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static void Click(Control root, string name) => Named<Button>(root, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static void Accept(FieldsDialog dialog) => dialog.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, UiText.Get("确定"))).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static void Text(FieldsDialog dialog, string key, string value) => ((TextBox)dialog.Input(key)).Text = value;
        try
        {
            _workspace.New(); ShowBook(); var sheet = Document.Snapshot.ActiveSheetId;
            var second = Document.AddSheet("范围测试").Id;
            Document.Session.UpdateNavigation(sheet, new(2, 1)); Synchronize();
            await ExecuteAsync(new SetCellsCommand([new(new(sheet, 2, 1), CellInput.Parse("20")), new(new(sheet, 2, 2), CellInput.Parse("30")),
                new(new(sheet, 3, 1), CellInput.Parse("40")), new(new(sheet, 0, 0), CellInput.Parse("=SUM(ELN)"))]));
            _grid.SelectRange(new(2, 1, 6, 2), new(2, 1));
            _dialogDriver = async window =>
            {
                await Layout(); var editor = (FieldsDialog)window;
                check(editor.String("formula") == "='Sheet1'!$B$3:$C$7", "Selection-created name defaults to an absolute sheet-qualified rectangle");
                check(((ComboBox)editor.Input("scope")).ItemCount == 3, "Name editor offers workbook and every worksheet scope");
                Text(editor, "name", "ELN"); Text(editor, "comment", "目标配额\n供脚本访问");
                await capture(editor, "name-editor"); Accept(editor);
            };
            BuildGridContextMenu(_grid.ContextMenu!, null);
            var create = _grid.ContextMenu!.Items.OfType<MenuItem>().Single(m => Equals(m.Header, UiText.Get("创建命名区域…")));
            create.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation; await Layout();
            var definition = Document.Snapshot.DefinedNames.Single(n => n.Name == "ELN");
            check(definition.Scope is null && definition.Comment == "目标配额\n供脚本访问", "Right-click create stores the workbook name and multiline comment");
            check(Document.GetCellValue(0, 0, 0) == "90", "Created range participates in ordinary formula calculation");
            await ExecuteAsync(new DefineNameCommand(new("Baseline", "=12", Comment: "参考值")));
            var managerCompleted = false;
            _dialogDriver = async window =>
            {
                await Layout(); var manager = (NameManagerDialog)window;
                var list = Named<ListBox>(manager, "NameManagerList"); var search = Named<TextBox>(manager, "NameManagerSearch");
                var filter = Named<ComboBox>(manager, "NameManagerFilter");
                check(list.ItemCount == 2, "Name manager lists existing definitions in one table");
                search.Text = "供脚本"; await Layout();
                check(list.ItemCount == 1 && Named<TextBox>(manager, "NameManagerReference").Text == definition.Formula,
                    "Name manager searches comments and shows the complete reference");
                check(list.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text == "{20, 30, 40, , , , …}"),
                    "Name manager displays a bounded value preview for a multi-cell named range");
                check(NamePreview(Document.Snapshot, new("WholeColumnPreview", "='Sheet1'!$B:$B")).EndsWith(", …}", StringComparison.Ordinal),
                    "Whole-column name previews remain bounded to six cells");
                _dialogDriver = async editorWindow =>
                {
                    await Layout(); var editor = (FieldsDialog)editorWindow;
                    check(!editor.Input("scope").IsEnabled, "Editing a name preserves its original scope");
                    Text(editor, "name", "Baseline"); Accept(editor); await Layout();
                    check(editor.IsVisible && Document.Snapshot.DefinedNames.Any(n => n.Name == "ELN"), "Duplicate names stay in the editor without overwriting an existing name");
                    Text(editor, "name", "Allocation"); Text(editor, "comment", "脚本配额区域"); Accept(editor);
                };
                Click(manager, "NameManagerEdit"); await manager.PendingOperation; await Layout();
                check(Document.Snapshot.DefinedNames.Single(n => n.Name == "Allocation").Comment == "脚本配额区域" &&
                    Document.GetCellInput(0, 0, 0) == "=SUM(Allocation)", "Manager rename updates formulas and comments atomically");
                search.Text = ""; await Layout();
                _dialogDriver = async editorWindow =>
                {
                    await Layout(); var editor = (FieldsDialog)editorWindow;
                    Text(editor, "name", "Allocation"); ((ComboBox)editor.Input("scope")).SelectedIndex = 2;
                    Text(editor, "comment", "工作表局部名称"); Accept(editor);
                };
                Click(manager, "NameManagerNew"); await manager.PendingOperation; await Layout();
                check(Document.Snapshot.DefinedNames.Any(n => n.Name == "Allocation" && n.Scope == second), "Manager creates a same-spelling name in another worksheet scope");
                filter.SelectedIndex = 1; await Layout(); check(list.ItemCount == 2, "Workbook filter excludes sheet-local names");
                filter.SelectedIndex = 3; await Layout(); check(list.ItemCount == 1, "Worksheet filter selects the requested local scope");
                filter.SelectedIndex = 0; search.Text = ""; await Layout(); await capture(manager, "name-manager");
                filter.SelectedIndex = 3; await Layout();
                Click(manager, "NameManagerDelete"); await manager.PendingOperation; await Layout();
                check(list.ItemCount == 0 && !Named<Button>(manager, "NameManagerEdit").IsEnabled,
                    "Manager deletes the selected local name and refreshes action availability");
                managerCompleted = true; Click(manager, "NameManagerClose");
            };
            await PerformAsync(NamesAsync);
            check(managerCompleted, "Name manager completes all dialog interactions without a swallowed handler failure");
            await History(false);
            check(Document.Snapshot.DefinedNames.Any(n => n.Name == "Allocation" && n.Scope == second), "One undo restores the deleted name with its scope");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = window => { window.Close(false); return Task.CompletedTask; };
            await CreateNameAsync(); check(Document.Snapshot.Revision == revision, "Cancelling name creation leaves the workbook unchanged");
            _dialogDriver = window =>
            {
                var editor = (FieldsDialog)window; Text(editor, "name", "StaleName");
                _workspace.Editor.SetCell(0, 9, 0, "newer"); editor.Close(true); return Task.CompletedTask;
            };
            try { await CreateNameAsync(); throw new Exception("Stale name creation was accepted."); }
            catch (InvalidOperationException) { check(!Document.Snapshot.DefinedNames.Any(n => n.Name == "StaleName"), "Name creation rejects a stale document revision"); }
        }
        finally { _dialogDriver = null; }
    }
}
