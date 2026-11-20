using System.Data;
using System.Data.Common;
using System.Text;
using Microsoft.Data.Sqlite;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public static partial class DatabaseTransfer
{
    private enum ColumnKind { Empty, Text, Number, Boolean, Date, Mixed }

    /// <summary>Creates a new table or appends to an existing table. No existing rows are deleted.
    /// Mixed-type columns are exported as invariant text; dates use the workbook's date system.
    /// SQLite DDL and rows commit together. ODBC DDL can leave an empty table after a failed row transaction.</summary>
    public static int ExportTable(DatabaseConnectionOptions options, DatabaseTable table, WorkbookSnapshot workbook,
        SheetId sheetId, CellRectangle range, bool firstRowHeaders = true,
        DatabaseWriteMode mode = DatabaseWriteMode.Create, CancellationToken cancellation = default)
    {
        ValidateOptions(options); ValidateTable(table, options); ArgumentNullException.ThrowIfNull(workbook);
        if (workbook.FileState?.FormulaReadOnly == true)
            throw new InvalidOperationException(MessageText.Get("只读公式工作簿不能直接导出数据库；请先明确导出值快照。"));
        if (!Enum.IsDefined(mode)) throw new ArgumentException(nameof(mode));
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Rows <= 0 || range.Columns <= 0 || range.Count > WorkbookLimits.MaxCells)
            throw new ArgumentException(MessageText.Get("数据库导出范围必须为有效矩形，最多 1,000,000 格。"));
        cancellation.ThrowIfCancellationRequested();
        var sheet = workbook.GetSheet(sheetId);
        var kinds = new ColumnKind[range.Columns];
        var names = new string[range.Columns];
        var rows = new List<object?[]>();
        long textBytes = 0;
        for (var col = 0; col < names.Length; col++)
        {
            names[col] = firstRowHeaders ? Header(Value(range.Top, range.Left + col), col) : "Column" + (col + 1);
            ValidateIdentifier(names[col], options);
        }
        if (names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Length)
            throw new ArgumentException(MessageText.Get("数据库列名重复；请修改首行标题后重试。"));
        for (var row = range.Top + (firstRowHeaders ? 1 : 0); row <= range.Bottom; row++)
        {
            cancellation.ThrowIfCancellationRequested();
            var values = new object?[range.Columns];
            for (var col = 0; col < values.Length; col++)
            {
                var value = Value(row, range.Left + col); values[col] = value;
                var kind = value switch { null => ColumnKind.Empty, string => ColumnKind.Text, double => ColumnKind.Number,
                    bool => ColumnKind.Boolean, DateTime => ColumnKind.Date, _ => throw new InvalidOperationException() };
                if (kind != ColumnKind.Empty)
                    kinds[col] = kinds[col] == ColumnKind.Empty ? kind : kinds[col] == kind ? kind : ColumnKind.Mixed;
            }
            rows.Add(values);
        }
        // Reject oversized data and stale results before opening/creating the destination.
        cancellation.ThrowIfCancellationRequested();
        var createdOdbcTable = false;
        try
        {
            using var connection = Open(options, write: true);
            cancellation.ThrowIfCancellationRequested();
            var exists = TableExists(connection, options, table, cancellation);
            if (mode == DatabaseWriteMode.Create && exists)
                throw new InvalidOperationException(MessageText.Get("数据库表已存在；请改名或明确选择追加。"));
            if (mode == DatabaseWriteMode.Append && !exists)
                throw new InvalidOperationException(MessageText.Get("追加目标表不存在；请先选择创建新表。"));
            if (mode == DatabaseWriteMode.Append)
            {
                RequireTransactionalMySql(connection, options, table, cancellation);
                using var schemaCommand = Command(connection, options, "SELECT * FROM " + QualifiedName(table, options) + " WHERE 1=0");
                using var schemaCancel = RegisterCancellation(schemaCommand, cancellation);
                using var reader = schemaCommand.ExecuteReader(CommandBehavior.SchemaOnly);
                var existingNames = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
                if (existingNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() != existingNames.Length)
                    throw new InvalidOperationException(MessageText.Get("目标表列名存在大小写歧义，无法安全追加。"));
                if (!firstRowHeaders)
                {
                    if (existingNames.Length != names.Length) throw new InvalidOperationException(MessageText.Get("按列顺序追加时，源数据与目标表列数必须一致。"));
                    names = existingNames;
                }
                else
                {
                    for (var col = 0; col < names.Length; col++)
                        names[col] = existingNames.SingleOrDefault(n => n.Equals(names[col], StringComparison.OrdinalIgnoreCase))
                            ?? throw new InvalidOperationException(MessageText.Get("目标表缺少源数据标题对应的列，未写入任何行。"));
                }
                if (connection is SqliteConnection)
                    for (var col = 0; col < names.Length; col++)
                    {
                        var ordinal = Array.FindIndex(existingNames, n => n.Equals(names[col], StringComparison.OrdinalIgnoreCase));
                        var declared = reader.GetDataTypeName(ordinal).ToUpperInvariant();
                        // SQLite checks INT before text markers (CHARINT therefore has integer affinity).
                        var textAffinity = !declared.Contains("INT", StringComparison.Ordinal) && (declared.Contains("CHAR", StringComparison.Ordinal)
                            || declared.Contains("CLOB", StringComparison.Ordinal) || declared.Contains("TEXT", StringComparison.Ordinal)
                            || declared.Contains("BLOB", StringComparison.Ordinal) || declared.Length == 0);
                        if (kinds[col] is ColumnKind.Text or ColumnKind.Mixed && !textAffinity)
                            throw new InvalidOperationException(MessageText.Get("文本列不能追加到具有数值或日期转换规则的 SQLite 列；请调整目标列类型，避免丢失前导零或原文。"));
                    }
            }
            else if (connection is not SqliteConnection)
            {
                // Some ODBC engines implicitly commit DDL. Keep it outside the row transaction and
                // report the possible empty table explicitly if a subsequent operation fails.
                using var create = Command(connection, options, CreateSql());
                using var createCancel = RegisterCancellation(create, cancellation);
                create.ExecuteNonQuery(); createdOdbcTable = true;
            }
            cancellation.ThrowIfCancellationRequested();
            using var transaction = connection.BeginTransaction();
            if (mode == DatabaseWriteMode.Create && connection is SqliteConnection)
            {
                using var create = Command(connection, options, CreateSql(), transaction);
                using var createCancel = RegisterCancellation(create, cancellation);
                create.ExecuteNonQuery();
            }
            using var insert = Command(connection, options,
                "INSERT INTO " + QualifiedName(table, options) + " (" + string.Join(", ", names.Select(n => QuoteIdentifier(n, options))) + ") VALUES ("
                + string.Join(", ", Enumerable.Range(0, names.Length).Select(i => options.Provider == DatabaseProvider.Sqlite ? "@p" + i : "?")) + ")", transaction);
            for (var col = 0; col < names.Length; col++)
            {
                var parameter = insert.CreateParameter(); parameter.ParameterName = "@p" + col;
                parameter.DbType = kinds[col] switch { ColumnKind.Number => DbType.Double, ColumnKind.Boolean => DbType.Boolean,
                    ColumnKind.Date => DbType.DateTime, _ => DbType.String };
                if (parameter.DbType == DbType.String) parameter.Size = WorkbookLimits.MaxCellTextLength;
                insert.Parameters.Add(parameter);
            }
            using var insertCancel = RegisterCancellation(insert, cancellation);
            foreach (var values in rows)
            {
                cancellation.ThrowIfCancellationRequested();
                for (var col = 0; col < names.Length; col++)
                    insert.Parameters[col].Value = values[col] is null ? DBNull.Value : kinds[col] == ColumnKind.Mixed ? InvariantText(values[col]!) : values[col]!;
                insert.ExecuteNonQuery();
            }
            cancellation.ThrowIfCancellationRequested(); transaction.Commit();
            return rows.Count;

            string CreateSql() => "CREATE TABLE " + QualifiedName(table, options) + " (" + string.Join(", ", names.Select((name, col) =>
                QuoteIdentifier(name, options) + " " + SqlType(kinds[col], options))) + ")"
                + (options.Provider == DatabaseProvider.Odbc && options.Dialect == DatabaseSqlDialect.MySql ? " ENGINE=InnoDB" : "");
        }
        catch (OperationCanceledException) when (createdOdbcTable)
        {
            throw new InvalidOperationException(MessageText.Get("数据库导出已取消，但新建的表可能保留，请检查数据库。"));
        }
        catch (DbException)
        {
            if (createdOdbcTable && cancellation.IsCancellationRequested)
                throw new InvalidOperationException(MessageText.Get("数据库导出已取消，但新建的表可能保留，请检查数据库。"));
            cancellation.ThrowIfCancellationRequested();
            if (createdOdbcTable) throw new InvalidOperationException(MessageText.Get("数据库行写入失败，行事务已结束；新建的空表可能保留，请检查数据库。连接信息未保存。"));
            throw DatabaseFailure();
        }
        catch (ArgumentException) when (options.Provider == DatabaseProvider.Odbc) { throw DatabaseFailure(); }

        object? Value(int row, int col)
        {
            var cell = sheet.GetCell(row, col);
            if (cell is null) return null;
            if ((cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null)
                && (cell.Calculation?.Value is null || cell.Calculation.InputRevision != workbook.Revision
                    || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException(MessageText.Get("数据库导出需要最新公式结果；请先重算。"));
            var value = cell.Value;
            switch (value.Kind)
            {
                case CellValueKind.Blank: return null;
                case CellValueKind.Text:
                    textBytes += Encoding.UTF8.GetByteCount(value.TextValue);
                    if (textBytes > MaximumTextBytes) throw new InvalidDataException(MessageText.Get("数据库导出文本超过 64 MiB 限制。"));
                    return value.TextValue;
                case CellValueKind.Number:
                    if (CellDisplayFormatter.IsDateFormat(workbook.Styles[sheet.GetStyleId(row, col)], value.NumberValue))
                    {
                        var date = new ExcelSerialDate(value.NumberValue, workbook.DateSystem);
                        if (!date.TryToDateTime(out var dateTime)) throw new InvalidOperationException(MessageText.Get("日期无法表示为数据库日期，请先转换为文本。"));
                        return dateTime;
                    }
                    return value.NumberValue;
                case CellValueKind.Boolean: return value.BooleanValue;
                default: throw new InvalidOperationException(MessageText.Get("数据库导出范围含错误值，请先修正或显式转换为文本。"));
            }
        }
    }

    private static string Header(object? value, int column) => value is null || value is string text && string.IsNullOrWhiteSpace(text)
        ? "Column" + (column + 1) : InvariantText(value);

    private static string InvariantText(object value) => value switch
    {
        string text => text, double number => number.ToString("R", Invariant), bool boolean => boolean ? "TRUE" : "FALSE",
        DateTime date => date.ToString("yyyy-MM-dd HH:mm:ss.fffffff", Invariant), _ => throw new InvalidOperationException()
    };

    private static string SqlType(ColumnKind kind, DatabaseConnectionOptions options) => kind switch
    {
        ColumnKind.Number => options.Provider == DatabaseProvider.Sqlite ? "REAL" : options.Dialect == DatabaseSqlDialect.SqlServer ? "FLOAT" : "DOUBLE PRECISION",
        ColumnKind.Boolean => options.Provider == DatabaseProvider.Odbc && options.Dialect == DatabaseSqlDialect.SqlServer ? "BIT" : "BOOLEAN",
        ColumnKind.Date => options.Provider == DatabaseProvider.Sqlite || options.Dialect == DatabaseSqlDialect.MySql ? "DATETIME"
            : options.Dialect == DatabaseSqlDialect.SqlServer ? "DATETIME2" : "TIMESTAMP",
        _ => options.Provider == DatabaseProvider.Sqlite ? "TEXT" : options.Dialect switch
        {
            DatabaseSqlDialect.SqlServer => "NVARCHAR(MAX)", DatabaseSqlDialect.PostgreSql => "TEXT",
            DatabaseSqlDialect.MySql => "LONGTEXT CHARACTER SET utf8mb4", _ => "VARCHAR(32767)"
        }
    };

    private static bool TableExists(DbConnection connection, DatabaseConnectionOptions options, DatabaseTable table, CancellationToken cancellation)
    {
        if (connection is SqliteConnection)
        {
            using var command = Command(connection, options, "SELECT COUNT(*) FROM sqlite_schema WHERE name = @name COLLATE NOCASE AND type IN ('table','view')");
            var parameter = command.CreateParameter(); parameter.ParameterName = "@name"; parameter.Value = table.Name; command.Parameters.Add(parameter);
            using var cancel = RegisterCancellation(command, cancellation);
            return Convert.ToInt64(command.ExecuteScalar(), Invariant) > 0;
        }
        using var schema = connection.GetSchema("Tables");
        foreach (DataRow row in schema.Rows)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!string.Equals(Convert.ToString(row["TABLE_NAME"], Invariant), table.Name, StringComparison.OrdinalIgnoreCase)) continue;
            if (string.IsNullOrEmpty(table.Schema) || schema.Columns.Contains("TABLE_SCHEM")
                && string.Equals(Convert.ToString(row["TABLE_SCHEM"], Invariant), table.Schema, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static void RequireTransactionalMySql(DbConnection connection, DatabaseConnectionOptions options, DatabaseTable table, CancellationToken cancellation)
    {
        if (options.Provider != DatabaseProvider.Odbc || options.Dialect != DatabaseSqlDialect.MySql) return;
        using var command = Command(connection, options, "SELECT ENGINE FROM information_schema.TABLES WHERE TABLE_SCHEMA = "
            + (string.IsNullOrEmpty(table.Schema) ? "DATABASE()" : "?") + " AND TABLE_NAME = ?");
        if (!string.IsNullOrEmpty(table.Schema)) Add(table.Schema);
        Add(table.Name);
        using var cancel = RegisterCancellation(command, cancellation);
        if (!string.Equals(Convert.ToString(command.ExecuteScalar(), Invariant), "InnoDB", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(MessageText.Get("MySQL 追加只支持 InnoDB 事务表。"));
        void Add(string value) { var parameter = command.CreateParameter(); parameter.Value = value; command.Parameters.Add(parameter); }
    }
}
