using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Globalization;
using System.Text;
using System.Xml;
using Microsoft.Data.Sqlite;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public enum DatabaseProvider { Sqlite, Odbc }
public enum DatabaseSqlDialect { Standard, SqlServer, PostgreSql, MySql }
public enum DatabaseWriteMode { Create, Append }

/// <summary>SQLite uses a local file path; ODBC uses an ephemeral connection string. Never persist or log this record.</summary>
public sealed record DatabaseConnectionOptions(DatabaseProvider Provider, string ConnectionString,
    DatabaseSqlDialect Dialect = DatabaseSqlDialect.Standard, int CommandTimeoutSeconds = 30)
{
    public override string ToString() => $"{Provider} ({Dialect})";
}

public sealed record DatabaseTable(string Name, string? Schema = null)
{
    public string DisplayName => string.IsNullOrEmpty(Schema) ? Name : Schema + "." + Name;
    public override string ToString() => DisplayName;
}

public sealed record DatabaseImportResult(WorkbookSnapshot Workbook, int Rows, int Columns, IReadOnlyList<string> Notices);

/// <summary>Host-owned, explicit data exchange. No connection is shared with a script worker.</summary>
public static partial class DatabaseTransfer
{
    public const int MaximumTables = 4096;
    public const int MaximumTextBytes = 64 * 1024 * 1024;
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static IReadOnlyList<DatabaseTable> ListTables(DatabaseConnectionOptions options, CancellationToken cancellation = default)
    {
        ValidateOptions(options); cancellation.ThrowIfCancellationRequested();
        try
        {
            using var connection = Open(options, write: false);
            cancellation.ThrowIfCancellationRequested();
            var tables = new List<DatabaseTable>();
            if (connection is SqliteConnection)
            {
                using var command = Command(connection, options, "SELECT name FROM sqlite_schema WHERE type IN ('table','view') AND substr(name,1,7) <> 'sqlite_' COLLATE NOCASE ORDER BY name");
                using var cancel = RegisterCancellation(command, cancellation);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    cancellation.ThrowIfCancellationRequested();
                    Add(new(reader.GetString(0)));
                }
            }
            else
            {
                using var schema = connection.GetSchema("Tables");
                foreach (DataRow row in schema.Rows)
                {
                    cancellation.ThrowIfCancellationRequested();
                    var type = Convert.ToString(row["TABLE_TYPE"], Invariant);
                    if (type is not ("TABLE" or "VIEW" or "BASE TABLE")) continue;
                    var name = Convert.ToString(row["TABLE_NAME"], Invariant)!;
                    var owner = schema.Columns.Contains("TABLE_SCHEM") ? Convert.ToString(row["TABLE_SCHEM"], Invariant) : null;
                    Add(new(name, string.IsNullOrEmpty(owner) ? null : owner));
                }
            }
            return tables.OrderBy(t => t.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray();
            void Add(DatabaseTable table)
            {
                ValidateTable(table, options);
                if (tables.Count == MaximumTables) throw new InvalidDataException(MessageText.Get("数据库表数量超过 4096，请使用范围更小的数据库连接。"));
                tables.Add(table);
            }
        }
        catch (DbException) { cancellation.ThrowIfCancellationRequested(); throw DatabaseFailure(); }
        catch (ArgumentException) when (options.Provider == DatabaseProvider.Odbc) { throw DatabaseFailure(); }
    }

    public static DatabaseImportResult ImportTable(DatabaseConnectionOptions options, DatabaseTable table,
        bool includeHeaders = true, CancellationToken cancellation = default)
    {
        ValidateOptions(options); ValidateTable(table, options); cancellation.ThrowIfCancellationRequested();
        try
        {
            using var connection = Open(options, write: false);
            using var command = Command(connection, options, "SELECT * FROM " + QualifiedName(table, options));
            using var cancel = RegisterCancellation(command, cancellation);
            using var reader = command.ExecuteReader(CommandBehavior.SingleResult | CommandBehavior.SequentialAccess);
            return ReadResult(reader, table.Name, includeHeaders, cancellation);
        }
        catch (DbException) { cancellation.ThrowIfCancellationRequested(); throw DatabaseFailure(); }
        catch (ArgumentException) when (options.Provider == DatabaseProvider.Odbc) { throw DatabaseFailure(); }
    }

    internal static DatabaseImportResult ReadResult(DbDataReader reader, string sheetName, bool headers, CancellationToken cancellation)
    {
        var columns = reader.FieldCount;
        if (columns is <= 0 or > WorkbookLimits.MaxColumns) throw new InvalidDataException(MessageText.Get("数据库结果列数超出工作表限制。"));
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var notices = new HashSet<string>();
        var row = headers ? 1 : 0;
        long textBytes = 0;
        if (headers)
            for (var col = 0; col < columns; col++) Add(0, col, CellValue.FromText(reader.GetName(col)), 0);
        while (reader.Read())
        {
            cancellation.ThrowIfCancellationRequested();
            if (row >= WorkbookLimits.MaxRows || (long)(row + 1) * columns > WorkbookLimits.MaxCells)
                throw new InvalidDataException(MessageText.Get("数据库结果超过工作表行数或 1,000,000 格限制，请缩小源表范围。"));
            for (var col = 0; col < columns; col++)
            {
                cancellation.ThrowIfCancellationRequested();
                if (reader.IsDBNull(col)) { Add(row, col, CellValue.Blank, 0); continue; }
                var fieldType = reader.GetFieldType(col);
                if (fieldType == typeof(byte[])) throw new InvalidDataException(MessageText.Get("数据库结果含二进制字段；请先在数据库中转换为明确的文本表示。"));
                if (fieldType == typeof(string) && reader.GetChars(col, 0, null, 0, 0) > WorkbookLimits.MaxCellTextLength)
                    throw new InvalidDataException(MessageText.Get("数据库文本字段超过 32767 字符限制。"));
                object value = reader.GetValue(col);
                // SQLite stores booleans and timestamps as integer/text. Only an explicit declared
                // type is interpreted; arbitrary numeric-looking or ISO-looking TEXT stays text.
                if (reader is SqliteDataReader sqlite)
                {
                    var declared = sqlite.GetDataTypeName(col).ToUpperInvariant();
                    if (declared is "BOOL" or "BOOLEAN" && value is long integer && integer is 0 or 1) value = integer == 1;
                    else if (declared is "DATETIME" or "TIMESTAMP" && value is string text)
                    {
                        if (DateTime.TryParseExact(text, ["yyyy-MM-dd HH:mm:ss.FFFFFFF", "yyyy-MM-ddTHH:mm:ss.FFFFFFF", "yyyy-MM-dd"],
                            Invariant, DateTimeStyles.None, out var date)) value = date;
                        else notices.Add("无法无损转换的数据库日期（包括时区）按原始文本保留。");
                    }
                }
                var (converted, style) = ConvertValue(value, notices);
                Add(row, col, converted, style);
            }
            row++;
        }
        cancellation.ThrowIfCancellationRequested();
        var safeName = new string(sheetName.Select(c => char.IsControl(c) || ":\\/?*[]".Contains(c) ? '_' : c).ToArray()).Trim().Trim('\'');
        if (safeName.Length > WorkbookLimits.MaxSheetNameLength)
        {
            var end = WorkbookLimits.MaxSheetNameLength;
            if (char.IsHighSurrogate(safeName[end - 1])) end--;
            safeName = safeName[..end].TrimEnd('\'');
        }
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "Data";
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), safeName, cells)], styles: [new(), new() { NumberFormatId = 164, NumberFormatCode = "yyyy-mm-dd hh:mm:ss" }],
            scriptModules: ScriptModuleSet.Single("", ScriptModuleSet.CurrentApiVersion));
        return new(book, row - (headers ? 1 : 0), columns, notices.ToArray());

        void Add(int r, int c, CellValue value, int style)
        {
            if (value.Kind == CellValueKind.Text)
            {
                WorkbookLimits.ValidateText(value.TextValue); XmlConvert.VerifyXmlChars(value.TextValue);
                textBytes += Encoding.UTF8.GetByteCount(value.TextValue);
                if (textBytes > MaximumTextBytes) throw new InvalidDataException(MessageText.Get("数据库结果文本超过 64 MiB 限制。"));
            }
            cells.Add(new(r, c), new(CellContent.FromLiteral(value, style)));
        }
    }

    private static (CellValue Value, int Style) ConvertValue(object value, HashSet<string> notices)
    {
        switch (value)
        {
            case string text: return (CellValue.FromText(text), 0);
            case char character: return (CellValue.FromText(character.ToString()), 0);
            case bool boolean: return (CellValue.FromBoolean(boolean), 0);
            case byte or sbyte or short or ushort or int or uint: return (CellValue.FromNumber(Convert.ToDouble(value, Invariant)), 0);
            case long integer when integer is >= -999_999_999_999_999 and <= 999_999_999_999_999: return (CellValue.FromNumber(integer), 0);
            case ulong integer when integer <= 999_999_999_999_999: return (CellValue.FromNumber(integer), 0);
            case long or ulong:
                notices.Add("超过 15 位的整数按文本保留，避免精度损失。");
                return (CellValue.FromText(Convert.ToString(value, Invariant)!), 0);
            case decimal number:
                var numeric = (double)number;
                if (decimal.TryParse(numeric.ToString("G15", Invariant), NumberStyles.Float, Invariant, out var restored) && restored == number)
                    return (CellValue.FromNumber(numeric), 0);
                notices.Add("高精度十进制数按文本保留，避免精度损失。");
                return (CellValue.FromText(number.ToString(Invariant)), 0);
            case double number when double.IsFinite(number): return (CellValue.FromNumber(number), 0);
            case float number when float.IsFinite(number): return (CellValue.FromNumber(number), 0);
            case DateTime date when date.Kind == DateTimeKind.Unspecified:
                return (CellValue.FromNumber(ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Windows1900).Serial), 1);
            case DateTime date:
                notices.Add("带时区的日期按 ISO 文本保留。"); return (CellValue.FromText(date.ToString("O", Invariant)), 0);
            case DateTimeOffset date:
                notices.Add("带时区的日期按 ISO 文本保留。"); return (CellValue.FromText(date.ToString("O", Invariant)), 0);
            case DateOnly date: return (CellValue.FromNumber(ExcelSerialDate.FromDateTime(date.ToDateTime(TimeOnly.MinValue), ExcelDateSystem.Windows1900).Serial), 1);
            case TimeSpan duration: notices.Add("时间间隔按固定格式文本保留。"); return (CellValue.FromText(duration.ToString("c", Invariant)), 0);
            case TimeOnly time: notices.Add("独立时间按固定格式文本保留。"); return (CellValue.FromText(time.ToString("O", Invariant)), 0);
            case Guid id: return (CellValue.FromText(id.ToString("D")), 0);
            default: throw new InvalidDataException(MessageText.Get("数据库结果含不支持的数据类型，请在数据库中显式转换后再导入。"));
        }
    }

    private static DbConnection Open(DatabaseConnectionOptions options, bool write)
    {
        DbConnection connection;
        if (options.Provider == DatabaseProvider.Sqlite)
        {
            var path = Path.GetFullPath(options.ConnectionString);
            if (!write && !File.Exists(path)) throw new FileNotFoundException(MessageText.Get("SQLite 文件不存在。"));
            connection = new SqliteConnection(new SqliteConnectionStringBuilder
            {
                DataSource = path, Mode = write ? SqliteOpenMode.ReadWriteCreate : SqliteOpenMode.ReadOnly,
                Pooling = false, ForeignKeys = true, DefaultTimeout = options.CommandTimeoutSeconds
            }.ConnectionString);
        }
        else connection = new OdbcConnection(options.ConnectionString);
        try
        {
            connection.Open();
            if (connection is SqliteConnection)
            {
                using var defensive = Command(connection, options, "PRAGMA trusted_schema=OFF"); defensive.ExecuteNonQuery();
            }
            return connection;
        }
        catch { connection.Dispose(); throw; }
    }

    private static DbCommand Command(DbConnection connection, DatabaseConnectionOptions options, string sql, DbTransaction? transaction = null)
    {
        var command = connection.CreateCommand(); command.CommandText = sql;
        command.CommandTimeout = options.CommandTimeoutSeconds; command.Transaction = transaction; return command;
    }

    private static CancellationTokenRegistration RegisterCancellation(DbCommand command, CancellationToken cancellation) => cancellation.Register(() =>
    {
        try { command.Cancel(); } catch (Exception ex) when (ex is DbException or InvalidOperationException or NotSupportedException) { }
    });

    private static InvalidOperationException DatabaseFailure() => new(MessageText.Get("数据库操作失败。请检查连接、权限、驱动、表结构与约束；连接信息未保存。"));

    private static void ValidateOptions(DatabaseConnectionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (!Enum.IsDefined(options.Provider) || !Enum.IsDefined(options.Dialect) || options.CommandTimeoutSeconds is < 1 or > 120)
            throw new ArgumentException(MessageText.Get("数据库连接选项无效；超时必须为 1–120 秒。"));
        if (string.IsNullOrWhiteSpace(options.ConnectionString)) throw new ArgumentException(MessageText.Get("请输入数据库连接或 SQLite 文件路径。"));
    }

    private static void ValidateTable(DatabaseTable table, DatabaseConnectionOptions options)
    {
        ArgumentNullException.ThrowIfNull(table); ValidateIdentifier(table.Name, options);
        if (!string.IsNullOrEmpty(table.Schema)) ValidateIdentifier(table.Schema, options);
        if (options.Provider == DatabaseProvider.Sqlite && !string.IsNullOrEmpty(table.Schema))
            throw new ArgumentException(MessageText.Get("SQLite 文件不支持跨数据库表名。"));
    }

    private static void ValidateIdentifier(string name, DatabaseConnectionOptions options)
    {
        var maximum = options.Provider == DatabaseProvider.Odbc && options.Dialect == DatabaseSqlDialect.MySql ? 64 : 128;
        if (string.IsNullOrWhiteSpace(name) || name.Length > maximum || name.Any(char.IsControl))
            throw new ArgumentException(MessageText.Format($"数据库名称必须为 1–{maximum} 个字符，且不含控制字符。"));
    }

    internal static string QuoteIdentifier(string name, DatabaseConnectionOptions options)
    {
        ValidateIdentifier(name, options);
        return options.Provider == DatabaseProvider.Sqlite ? "\"" + name.Replace("\"", "\"\"") + "\"" : options.Dialect switch
        {
            DatabaseSqlDialect.SqlServer => "[" + name.Replace("]", "]]") + "]",
            DatabaseSqlDialect.MySql => "`" + name.Replace("`", "``") + "`",
            _ => "\"" + name.Replace("\"", "\"\"") + "\""
        };
    }

    private static string QualifiedName(DatabaseTable table, DatabaseConnectionOptions options) =>
        (string.IsNullOrEmpty(table.Schema) ? "" : QuoteIdentifier(table.Schema, options) + ".") + QuoteIdentifier(table.Name, options);
}
