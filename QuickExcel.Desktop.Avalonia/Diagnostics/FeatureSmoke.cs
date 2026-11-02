using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunDataFeatureChecksAsync(Action<bool, string> check)
    {
        static void Click(Window window, string caption) => window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == caption).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static T Named<T>(Window window, string name) where T : Control => window.GetVisualDescendants().OfType<T>().Single(c => AutomationProperties.GetName(c) == name);
        _workspace.New(); ShowBook(); var sheet = Document.Snapshot.ActiveSheetId;
        _workspace.Editor.SetCells(0, GridClipboard.Prepare("品名\t数量\n甲\t2\n乙\t3\n丙\t1", new(0, 0, 0, 0), false)); Synchronize();
        _grid.SelectRange(new(0, 0, 3, 1), new(0, 0));
        _dialogDriver = async window =>
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            check(window is SortDialog, "Advanced sort opens native sort dialog"); Named<ComboBox>(window, "排序列").SelectedIndex = 1; Click(window, "排序");
        };
        await SortAsync(null);
        check(Document.GetCellValue(0, 0, 0) == "品名" && Document.GetCellValue(0, 1, 0) == "丙" && Document.GetCellValue(0, 3, 0) == "乙", "Native multi-level sort keeps header and moves complete records");
        await History(false); check(Document.GetCellValue(0, 1, 0) == "甲", "Native sort has one undo entry");
        await History(true); await CalculateAsync();
        await EstablishFilterAsync(); _grid.Select(new(0, 1));
        _dialogDriver = async window =>
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            check(window is FilterDialog, "Header filtering opens native candidates and conditions");
            Named<ComboBox>(window, "筛选方式").SelectedIndex = 1;
            Named<ComboBox>(window, "第一个条件").SelectedIndex = (int)FilterOperator.Greater;
            Named<TextBox>(window, "第一个条件值").Text = "1"; Click(window, "应用");
        };
        await FilterAsync(1);
        check(CurrentView.IsFilteredOut(Document.Snapshot, 0, 1) && !CurrentView.IsFilteredOut(Document.Snapshot, 0, 2), "Native custom filter persists conditions and applies row projection");
        check(_grid.GetCellBounds(new(1, 0)) is null, "Filtered rows disappear from grid geometry");
        _grid.SelectRange(new(0, 2, 3, 2), new(0, 2)); ((MemoryClipboard)_clipboard).Value = new("paste"); var before = Document.Snapshot;
        try { await PasteAsync("all"); throw new InvalidOperationException("Filtered paste succeeded"); }
        catch (InvalidOperationException ex) when (ex.Message != "Filtered paste succeeded") { check(ReferenceEquals(before, Document.Snapshot), "Filtered external paste rejects atomically without remapping"); }
        await ClearFiltersAsync(); check(!CurrentView.IsFilteredOut(Document.Snapshot, 0, 1), "Clear filter conditions restores visible rows"); await RemoveFilterAsync();
        _grid.SelectRange(new(0, 0, 3, 1), new(0, 0));
        _dialogDriver = async window =>
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var form = (FieldsDialog)window; ((TextBox)form.Input("name")).Text = "Sales"; Click(window, "确定");
        };
        await CreateTableAsync(); check(Document.Snapshot.Sheets[0].Data.Tables.Single().Name == "Sales", "Native table creation commits persisted table and style");
        await History(false); check(Document.Snapshot.Sheets[0].Data.Tables.IsEmpty, "Table creation is one undo entry");
        await History(true); var table = Document.Snapshot.Sheets[0].Data.Tables.Single();
        await ExecuteAsync(new SetTableCalculatedColumnCommand(sheet, table.Id, table.Columns[1].Id, "=LEN(A2)+2")); await CalculateAsync();
        check(Document.GetCellValue(0, 3, 1) == "3", "Calculated table column uses shared evaluator");
        await History(false);
        var formulas = WorkbookSearch.Find(Document.Snapshot, 0, new("甲", SearchScope.Selection), selected: a => a.SheetId == sheet);
        await ExecuteAsync(WorkbookSearch.Replace(formulas, new("甲", SearchScope.Selection), "替换"));
        check(Document.GetCellValue(0, 2, 0) == "替换", "Shared search applies one explicit replacement"); await History(false);
        _grid.Select(new(1, 0)); await ExecuteAsync(new SetCellProtectionCommand(sheet, new(1, 0, 1, 0), false));
        await ExecuteAsync(new SetProtectionCommand(sheet, new()));
        check(_grid.BeginEdit(), "Unlocked cell remains editable on protected sheet"); _grid.Cancel();
        _grid.Select(new(2, 0)); check(!_grid.BeginEdit(), "Locked protected cell refuses native editor");
        await ExecuteAsync(new SetProtectionCommand(sheet, null));
        _grid.Select(new(8, 0));
        _dialogDriver = async window =>
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var form = (FieldsDialog)window; ((ComboBox)form.Input("kind")).SelectedIndex = (int)ValidationKind.List;
            ((TextBox)form.Input("first")).Text = "=\"甲,乙\""; Click(window, "确定");
        };
        await EditValidationAsync();
        check(_validationPanel.IsVisible && _validationValues.IsVisible, "Native validation editor installs visible list and prompt");
        _validationValues.IsDropDownOpen = true; await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        _validationValues.SelectedIndex = 1; await _operation; _validationValues.IsDropDownOpen = false;
        check(Document.GetCellValue(0, 8, 0) == "乙", "Native validation list commits selected typed value");
        await History(false); check(Document.GetCellInput(0, 8, 0) == "", "Validation-list choice is one undo entry");
        var rule1 = new ConditionalFormatRule([new(1, 1, 3, 1)], ConditionalFormatKind.CellIs, 1, new(FillColor: "FFFF0000"), "=2", comparison: ConditionalComparison.GreaterThan);
        var rule2 = new ConditionalFormatRule([new(1, 1, 3, 1)], ConditionalFormatKind.UniqueValues, 2, new(Bold: true));
        await ExecuteAsync(new SetConditionalFormatsCommand(sheet, [rule1, rule2]));
        _dialogDriver = async window => { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); Click(window, "下移"); Click(window, "应用"); };
        await ConditionalFormatsAsync();
        check(Document.Snapshot.Sheets[0].Data.ConditionalFormats[0].Kind == ConditionalFormatKind.UniqueValues, "Native rule manager commits reordered priorities");
        await History(false); check(Document.Snapshot.Sheets[0].Data.ConditionalFormats[0].Kind == ConditionalFormatKind.CellIs, "Rule manager is one undo transaction");
        _grid.Select(new(12, 10)); _workspace.Editor.SetCell(0, 12, 10, "'1,234.5"); Synchronize();
        void CleanDriver(bool apply)
        {
            _dialogDriver = async window =>
            {
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                _dialogDriver = async preview => { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); Click(preview, apply ? "应用" : "取消"); };
                Click(window, "确定");
            };
        }
        CleanDriver(false); var revision = Document.Snapshot.Revision; await ConvertNumbersAsync();
        check(Document.Snapshot.Revision == revision && Document.Snapshot.Sheets[0].GetCell(12, 10)!.Value.Kind == CellValueKind.Text, "Cancel native cleaning preview preserves original input");
        CleanDriver(true); await ConvertNumbersAsync();
        check(Document.Snapshot.Sheets[0].GetCell(12, 10)!.Value == CellValue.FromNumber(1234.5), "Native cleaning preview commits explicit numeric conversion");
        await History(false); check(Document.Snapshot.Sheets[0].GetCell(12, 10)!.Value.Kind == CellValueKind.Text, "Cleaning batch has one undo entry");
        _grid.Select(new(4, 0));
        _dialogDriver = async window =>
        { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); ((TextBox)((FieldsDialog)window).Input("text")).Text = "原生批注"; Click(window, "确定"); };
        await NoteAsync(); check(Document.Snapshot.Sheets[0].Annotations.Notes[new(4, 0)].Text.Text == "原生批注", "Native note editor commits note metadata");
        _grid.SelectRange(new(0, 0, 3, 1), new(0, 0));
        _dialogDriver = async window =>
        { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); ((TextBox)((FieldsDialog)window).Input("title")).Text = "数量"; Click(window, "确定"); };
        await ChartsAsync(); check(Document.Snapshot.Sheets[0].Data.Charts.Single().Title == "数量", "Native chart editor creates shared chart model");
        _dialogDriver = async window =>
        {
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var form = (FieldsDialog)window; ((ComboBox)form.Input("orientation")).SelectedIndex = 1; ((TextBox)form.Input("areas")).Text = "A1:B4";
            ((TextBox)form.Input("rows")).Text = "1:1"; ((TextBox)form.Input("header")).Text = "&C报表 &P/&N"; Click(window, "确定");
        };
        await PageSettingsAsync(); var printing = Document.Snapshot.Sheets[0].Printing;
        check(printing.Setup?.Orientation == "landscape" && printing.Areas.Length == 1 && printing.Titles.Length == 1 && printing.HeaderFooter?.OddHeader == "&C报表 &P/&N", "Native page settings preserve layout, area, titles and header as one command");
        await History(false); check(Document.Snapshot.Sheets[0].Printing.IsEmpty, "Page settings have one undo entry"); await History(true);
        _grid.SelectRange(new(20, 0, 22, 0), new(20, 0));
        _dialogDriver = async window => { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); Click(window, "确定"); };
        await GroupAsync(false); check(WorksheetOutline.Groups(Document.Snapshot.Sheets[0], SheetAxis.Row).Count == 1, "Native grouping preserves row metadata");
        _dialogDriver = async window =>
        { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); ((ComboBox)((FieldsDialog)window).Input("state")).SelectedIndex = 1; Click(window, "确定"); };
        await OutlineAsync(); check(Document.Snapshot.Sheets[0].Rows[20].Hidden, "Native outline dialog collapses group through command"); await History(false);
    }
}
