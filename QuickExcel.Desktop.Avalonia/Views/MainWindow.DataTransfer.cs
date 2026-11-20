using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task DataTransferAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var request = await ShowOwnedAsync<DataTransferRequest?>(new DataTransferWindow(_book.IsVisible));
        if (request is null) return;
        switch (request.Format)
        {
            case DataTransferFormat.Text:
                if (request.Export) await ExportTextAsync(); else await ImportTextAsync(useTransferDestination: true);
                break;
            case DataTransferFormat.Xlsx:
                if (!request.Export) await ImportXlsxDataAsync(request.Mode == 1);
                else if (request.Mode < 2) await ExportAsync(request.Mode == 1);
                else await ExportXlsxDataAsync(request.Mode == 3);
                break;
            default:
                await TransferDatabaseAsync(request.Format == DataTransferFormat.Sqlite, request.Export);
                break;
        }
    }

    private async Task<T> RunDataTransferAsync<T>(string title, Func<CancellationToken, T> action)
    {
        using var cancellation = new CancellationTokenSource();
        using var feedback = BeginExecution(UiText.Get(title), cancellation);
        return await Task.Run(() => action(cancellation.Token), cancellation.Token);
    }

    private void RequireTransferDocument(WorkbookDocument document, long revision)
    {
        if (!ReferenceEquals(document, Document) || Document.Snapshot.Revision != revision)
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
    }

    private void RequireTransferPath(string path)
    {
        if (_workspace.FilePath is { } source && string.Equals(Path.GetFullPath(path), Path.GetFullPath(source), OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new InvalidOperationException(UiText.Get("导出需要使用新副本路径，不能覆盖当前打开的工作簿。"));
    }

    private static TextBox TransferPreview(WorkbookSnapshot workbook)
    {
        var lines = new List<string>();
        foreach (var sheet in workbook.Sheets.Take(8))
        {
            lines.Add(UiText.Format($"{sheet.Name} · {sheet.UsedRowCount} 行 × {sheet.UsedColumnCount} 列"));
            var used = WorkbookTextExchange.UsedRange(sheet);
            for (var row = used.Top; row <= Math.Min(used.Bottom, used.Top + 7); row++)
                lines.Add(string.Join(" | ", Enumerable.Range(used.Left, Math.Min(used.Columns, 8)).Select(column =>
                {
                    var value = CellDisplayFormatter.Format(sheet.GetCell(row, column)?.Value ?? CellValue.Blank, workbook.Styles[sheet.GetStyleId(row, column)], workbook.DateSystem);
                    return value.Length > 100 ? value[..100] + "…" : value;
                })));
            lines.Add("");
        }
        return new TextBox { Text = string.Join("\n", lines), IsReadOnly = true, AcceptsReturn = true, Height = 180, TextWrapping = TextWrapping.NoWrap };
    }

    private async Task ApplyDataImportAsync(WorkbookSnapshot values, IReadOnlyList<string> notices, WorkbookDocument original, long revision)
    {
        RequireTransferDocument(original, revision);
        var dialog = new FieldsDialog(UiText.Get("数据导入预览"))
            .Note(UiText.Get("导入为独立工作表。导入当前工作簿可一次撤销；原文件保持不变。"))
            .Choice("destination", UiText.Get("导入到"), [UiText.Get("新工作簿"), UiText.Get("当前工作簿（新增工作表）")])
            .Choice("conflict", UiText.Get("工作表名称冲突"), [UiText.Get("自动重命名"), UiText.Get("跳过同名工作表"), UiText.Get("冲突时停止")]);
        for (var i = 0; i < values.Sheets.Length; i++) dialog.Text("name" + i, UiText.Get("目标名称：") + values.Sheets[i].Name, values.Sheets[i].Name);
        if (notices.Count > 0) dialog.Note(string.Join("\n", notices.Select(UiText.Get)));
        dialog.Preview(TransferPreview(values));
        var mapping = new TextBlock { TextWrapping = TextWrapping.Wrap };
        dialog.Preview(mapping);
        WorkbookSnapshot NamedSource() => values.WithSheets(values.Sheets.Select((sheet, i) => sheet.WithName(dialog.String("name" + i).Trim())));
        void UpdatePlan()
        {
            try
            {
                var source = NamedSource();
                mapping.Text = dialog.Index("destination") == 0 ? string.Join("\n", source.Sheets.Select(s => s.Name)) :
                    string.Join("\n", WorkbookDataTransfer.PlanImport(original.Snapshot, source, (SheetNameConflictPolicy)dialog.Index("conflict"))
                        .Select(p => p.SourceName + " → " + (p.Skipped ? UiText.Get("跳过") : p.TargetName)));
                mapping.Foreground = Brushes.Gray;
            }
            catch (Exception ex) { mapping.Text = ex.Message; mapping.Foreground = Brushes.Firebrick; }
        }
        dialog.Validate = () =>
        {
            var named = NamedSource();
            if (dialog.Index("destination") == 1)
            {
                if (original.IsReadOnly) throw new InvalidOperationException(UiText.Get(DocumentSession.ReadOnlyMessage));
                var plan = WorkbookDataTransfer.PlanImport(original.Snapshot, named, (SheetNameConflictPolicy)dialog.Index("conflict"));
                if (plan.All(p => p.Skipped)) throw new InvalidOperationException(UiText.Get("所有工作表均被跳过，请修改名称或冲突策略。"));
            }
        };
        ((ComboBox)dialog.Input("destination")).SelectionChanged += (_, _) => UpdatePlan();
        ((ComboBox)dialog.Input("conflict")).SelectionChanged += (_, _) => UpdatePlan();
        for (var i = 0; i < values.Sheets.Length; i++) ((TextBox)dialog.Input("name" + i)).TextChanged += (_, _) => UpdatePlan();
        UpdatePlan();
        if (!await dialog.ShowAsync(this)) return;
        dialog.Validate(); RequireTransferDocument(original, revision);
        var imported = NamedSource();
        if (dialog.Index("destination") == 0)
        {
            if (!await ConfirmDiscardAsync()) return;
            RequireTransferDocument(original, revision);
            _workspace.Import(imported); ShowBook();
        }
        else
        {
            if (!await ExecuteAsync(new ImportWorksheetDataCommand(imported, (SheetNameConflictPolicy)dialog.Index("conflict")), revision)) return;
            ShowBook();
        }
        _status.Text = UiText.Get("数据已导入，请保存工作簿。");
    }

    private async Task ImportXlsxDataAsync(bool fullWorkbook)
    {
        if (!await CommitEditorsAsync()) return;
        var original = Document; var revision = original.Snapshot.Revision;
        var file = await _files.OpenAsync(new(UiText.Get("导入 Excel 工作簿"), [new("Excel", ["*.xlsx"])]));
        if (file is null) return;
        var source = await RunDataTransferAsync("正在读取工作簿…", token => new TypedWorkbookStorage().Open(file.Path, token));
        RequireTransferDocument(original, revision);
        if (fullWorkbook)
        {
            var preview = new FieldsDialog(UiText.Get("完整工作簿副本"))
                .Note(UiText.Get("打开为未绑定保存路径的完整副本。保留公式、格式与受支持对象，以及原文件的只读 / 兼容性限制。"))
                .Preview(TransferPreview(source));
            if (source.FileState?.CompatibilityIssues.Length > 0) preview.Note(string.Join("\n", source.FileState.CompatibilityIssues));
            if (!await preview.ShowAsync(this)) return;
            RequireTransferDocument(original, revision);
            if (!await ConfirmDiscardAsync()) return;
            RequireTransferDocument(original, revision);
            _workspace.Import(source); ShowBook(); _status.Text = UiText.Get("工作簿副本已导入，请另存。"); return;
        }
        var selection = new FieldsDialog(UiText.Get("选择导入工作表"))
            .Note(UiText.Get("只导入结果值与数字格式。公式、图片、图表、批注、超链接、富文本、合并单元格和其他布局不进入数据副本。"));
        for (var i = 0; i < source.Sheets.Length; i++) selection.Check("sheet" + i, source.Sheets[i].Name, true);
        selection.Check("stale", UiText.Get("允许使用文件中的已有公式缓存（可能过期）"));
        selection.Validate = () => { if (!Enumerable.Range(0, source.Sheets.Length).Any(i => selection.Boolean("sheet" + i))) throw new ArgumentException(UiText.Get("请至少选择一个工作表。")); };
        if (!await selection.ShowAsync(this)) return;
        selection.Validate();
        var sheets = source.Sheets.Where((_, i) => selection.Boolean("sheet" + i)).Select(sheet => new DataTransferSheetSelection(sheet.Id)).ToArray();
        var allowStale = selection.Boolean("stale");
        var extracted = await RunDataTransferAsync("正在准备数据预览…", token => WorkbookDataTransfer.ExtractValues(source, sheets, allowStale, token));
        await ApplyDataImportAsync(extracted.Workbook, extracted.Notices, original, revision);
    }

    private async Task ExportXlsxDataAsync(bool selectionOnly)
    {
        if (!await FlushEditsAsync()) return;
        var document = Document; var snapshot = document.Snapshot;
        var range = selectionOnly ? Rectangle(SelectedRange()) : (CellRectangle?)null;
        var data = await RunDataTransferAsync("正在准备数据预览…", token => WorkbookDataTransfer.ExtractValues(snapshot, [new(snapshot.ActiveSheetId, Range: range)], cancellation: token));
        var preview = new FieldsDialog(UiText.Get("导出 Excel 数据"))
            .Note(UiText.Get("导出结果值与数字格式，包含隐藏行列；公式和其他对象不进入数据副本。"))
            .Note(string.Join("\n", data.Notices.Select(UiText.Get))).Preview(TransferPreview(data.Workbook));
        if (!await preview.ShowAsync(this)) return;
        var file = await _files.SaveAsync(new(UiText.Get("导出 Excel 数据"), [new("Excel", ["*.xlsx"])], snapshot.GetSheet(snapshot.ActiveSheetId).Name + ".xlsx", ".xlsx"));
        if (file is null) return;
        RequireTransferDocument(document, snapshot.Revision); RequireTransferPath(file.Path);
        await RunDataTransferAsync("正在导出工作簿…", token =>
        {
            token.ThrowIfCancellationRequested();
            return new TypedWorkbookStorage().ExportDataValuesChecked(file.Path, data.Workbook, FileVersions.Read(file.Path), token);
        });
        _status.Text = UiText.Get("已导出 ") + file.Path;
    }

    private async Task<DatabaseConnectionOptions?> DatabaseConnectionAsync(bool sqlite, bool export)
    {
        if (sqlite)
        {
            var file = export ? await _files.SaveAsync(new(UiText.Get("选择 SQLite 导出文件"), [new("SQLite", ["*.sqlite", "*.sqlite3", "*.db"])], "export.sqlite", ".sqlite")) :
                await _files.OpenAsync(new(UiText.Get("打开 SQLite 文件"), [new("SQLite", ["*.sqlite", "*.sqlite3", "*.db"]), new(UiText.Get("所有文件"), ["*"])]));
            if (file is null) return null;
            RequireTransferPath(file.Path);
            return new(DatabaseProvider.Sqlite, file.Path);
        }
        var dialog = new FieldsDialog(UiText.Get("连接 SQL 数据库"))
            .Note(UiText.Get("使用本机已安装的 ODBC 驱动。例如 DSN=MyDatabase;UID=user;PWD=… 。连接信息不会写入设置或工作簿。"))
            .Text("connection", UiText.Get("ODBC 连接字符串"))
            .Choice("dialect", UiText.Get("SQL 方言"), [UiText.Get("标准 SQL"), "SQL Server", "PostgreSQL", "MySQL"])
            .Text("timeout", UiText.Get("命令超时（秒）"), "30");
        ((TextBox)dialog.Input("connection")).PasswordChar = '●';
        dialog.Validate = () =>
        {
            if (string.IsNullOrWhiteSpace(dialog.String("connection"))) throw new ArgumentException(UiText.Get("请输入连接字符串。"));
            _ = dialog.Integer("timeout", 1, 120);
        };
        if (!await dialog.ShowAsync(this)) return null;
        dialog.Validate();
        return new(DatabaseProvider.Odbc, dialog.String("connection"), (DatabaseSqlDialect)dialog.Index("dialect"), dialog.Integer("timeout", 1, 120));
    }

    private async Task TransferDatabaseAsync(bool sqlite, bool export)
    {
        if (export ? !await FlushEditsAsync() : !await CommitEditorsAsync()) return;
        var document = Document; var snapshot = document.Snapshot;
        var connection = await DatabaseConnectionAsync(sqlite, export);
        if (connection is null) return;
        if (!export)
        {
            var tables = await RunDataTransferAsync("正在读取数据库表…", token => DatabaseTransfer.ListTables(connection, token));
            if (tables.Count == 0) throw new InvalidOperationException(UiText.Get("数据库中没有可导入的表或视图。"));
            var choice = new FieldsDialog(UiText.Get("选择数据库表"))
                .Choice("table", UiText.Get("表 / 视图"), tables.Select(t => t.DisplayName).ToArray())
                .Check("headers", UiText.Get("包含列名称标题行"), true);
            if (!await choice.ShowAsync(this)) return;
            var table = tables[choice.Index("table")]; var headers = choice.Boolean("headers");
            var imported = await RunDataTransferAsync("正在读取数据库数据…", token => DatabaseTransfer.ImportTable(connection, table, headers, token));
            await ApplyDataImportAsync(imported.Workbook, imported.Notices, document, snapshot.Revision); return;
        }
        var sheet = snapshot.GetSheet(snapshot.ActiveSheetId); var selected = Rectangle(SelectedRange());
        var options = new FieldsDialog(UiText.Get("导出到数据库"))
            .Note(UiText.Get("导出结果值，包含隐藏行列。首行可用作数据库列名；公式必须具有最新结果。追加按列名匹配，在一个事务中提交。"))
            .Note(UiText.Get("日期格式的数字转换为数据库日期；混合类型列转换为文本；空白单元格写为 NULL。"))
            .Choice("scope", UiText.Get("导出范围"), [UiText.Get("当前工作表"), UiText.Get("当前选区")])
            .Text("table", UiText.Get("目标表名称"), sheet.Name)
            .Text("schema", UiText.Get("Schema（可留空）"))
            .Choice("mode", UiText.Get("写入方式"), [UiText.Get("新建表（同名时停止）"), UiText.Get("追加到已有表")])
            .Check("headers", UiText.Get("首行作为列名称"), true);
        if (sqlite) options.Input("schema").IsEnabled = false;
        options.Validate = () => { if (string.IsNullOrWhiteSpace(options.String("table"))) throw new ArgumentException(UiText.Get("请输入目标表名称。")); };
        if (!await options.ShowAsync(this)) return;
        options.Validate();
        var range = options.Index("scope") == 0 ? WorkbookTextExchange.UsedRange(sheet) : selected;
        var tableName = new DatabaseTable(options.String("table").Trim(), string.IsNullOrWhiteSpace(options.String("schema")) ? null : options.String("schema").Trim());
        var mode = (DatabaseWriteMode)options.Index("mode"); var firstHeaders = options.Boolean("headers");
        var data = await RunDataTransferAsync("正在准备数据预览…", token => WorkbookDataTransfer.ExtractValues(snapshot, [new(sheet.Id, Range: range)], cancellation: token));
        var confirm = new FieldsDialog(UiText.Get("确认数据库写入"))
            .Note((mode == DatabaseWriteMode.Create ? UiText.Get("新建表：") : UiText.Get("追加到表：")) + tableName.DisplayName)
            .Note(UiText.Format($"将写入 {Math.Max(0, range.Rows - (firstHeaders ? 1 : 0))} 行、{range.Columns} 列。"))
            .Note(UiText.Get("不会删除或替换已有表。追加操作提交后不能通过工作簿撤销。"))
            .Preview(TransferPreview(data.Workbook));
        if (!sqlite)
            confirm.Note(UiText.Get("目标数据库按其列类型和约束接收数据，可能进行类型转换。请确认列定义后写入。"));
        if (!sqlite && mode == DatabaseWriteMode.Create)
            confirm.Note(UiText.Get("部分 ODBC 数据库会自动提交建表操作；写入失败时可能保留空表。"));
        if (!await confirm.ShowAsync(this)) return;
        RequireTransferDocument(document, snapshot.Revision);
        var rows = await RunDataTransferAsync("正在写入数据库…", token => DatabaseTransfer.ExportTable(connection, tableName, snapshot, sheet.Id, range, firstHeaders, mode, token));
        _status.Text = UiText.Format($"已向 {tableName.DisplayName} 写入 {rows} 行。");
    }
}
