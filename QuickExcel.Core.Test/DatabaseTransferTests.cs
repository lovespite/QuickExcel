using System.Data;
using Microsoft.Data.Sqlite;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class DatabaseTransferTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-database-" + Guid.NewGuid().ToString("N"));
    private string DatabasePath => Path.Combine(_directory, "exchange.sqlite");
    private DatabaseConnectionOptions Options => new(DatabaseProvider.Sqlite, DatabasePath);
    public DatabaseTransferTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    [Fact]
    public void SqliteRoundTripPreservesTypedValuesAndQuotedIdentifiers()
    {
        var book = Book([["ID", "Amount", "Enabled", "Text"], ["001", 12.5, true, "=1+1"], ["002", null, false, "中文\n文字"]]);
        var table = new DatabaseTable("sales\"; DROP TABLE x;--");
        Assert.Equal(2, DatabaseTransfer.ExportTable(Options, table, book, book.Sheets[0].Id, new(0, 0, 2, 3)));
        Assert.Equal(table, Assert.Single(DatabaseTransfer.ListTables(Options)));
        var imported = DatabaseTransfer.ImportTable(Options, table);
        Assert.Equal(2, imported.Rows); Assert.Equal(4, imported.Columns);
        var sheet = imported.Workbook.Sheets[0];
        Assert.Equal(CellValue.FromText("001"), sheet.GetCell(1, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(12.5), sheet.GetCell(1, 1)!.Value);
        Assert.Equal(CellValue.FromBoolean(true), sheet.GetCell(1, 2)!.Value);
        Assert.Equal(CellValue.FromText("=1+1"), sheet.GetCell(1, 3)!.Value);
        Assert.Equal(CellContentKind.Literal, sheet.GetCell(1, 3)!.Content.Kind);
        Assert.Equal(CellValue.Blank, sheet.GetCell(2, 1)!.Value);
        Assert.Equal(CellValue.FromBoolean(false), sheet.GetCell(2, 2)!.Value);
    }

    [Fact]
    public void AppendMapsHeadersAndConstraintFailureRollsBackEveryRow()
    {
        Execute("CREATE TABLE target (Id TEXT UNIQUE NOT NULL, Value REAL); INSERT INTO target VALUES ('old', 1)");
        var invalid = Book([["Value", "Id"], [2.0, "new"], [3.0, "old"]]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("target"), invalid, invalid.Sheets[0].Id,
            new(0, 0, 2, 1), mode: DatabaseWriteMode.Append));
        Assert.Equal(1L, Scalar("SELECT COUNT(*) FROM target"));
        var valid = Book([["Value", "Id"], [2.0, "new"]]);
        Assert.Equal(1, DatabaseTransfer.ExportTable(Options, new("target"), valid, valid.Sheets[0].Id,
            new(0, 0, 1, 1), mode: DatabaseWriteMode.Append));
        Assert.Equal(2.0, Scalar("SELECT Value FROM target WHERE Id='new'"));
    }

    [Fact]
    public void CreateConflictNeverReplacesExistingData()
    {
        Execute("CREATE TABLE target (Value TEXT); INSERT INTO target VALUES ('keep')");
        var book = Book([["Value"], ["replace"]]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("TARGET"), book, book.Sheets[0].Id, new(0, 0, 1, 0)));
        Assert.Equal("keep", Scalar("SELECT Value FROM target"));
    }

    [Fact]
    public void AppendEnforcesForeignKeysAndRollsBackAllRows()
    {
        Execute("CREATE TABLE parent (Id TEXT PRIMARY KEY); INSERT INTO parent VALUES ('valid'); CREATE TABLE child (ParentId TEXT REFERENCES parent(Id))");
        var book = Book([["ParentId"], ["valid"], ["missing"]]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("child"), book,
            book.Sheets[0].Id, new(0, 0, 2, 0), mode: DatabaseWriteMode.Append));
        Assert.Equal(0L, Scalar("SELECT COUNT(*) FROM child"));
    }

    [Fact]
    public void AppendWithoutHeadersUsesDestinationColumnOrder()
    {
        Execute("CREATE TABLE target (Code TEXT, Amount REAL)");
        var book = Book([["001", 12.0]]);
        Assert.Equal(1, DatabaseTransfer.ExportTable(Options, new("target"), book, book.Sheets[0].Id, new(0, 0, 0, 1), false, DatabaseWriteMode.Append));
        Assert.Equal("001", Scalar("SELECT Code FROM target"));
    }

    [Theory]
    [InlineData("INTEGER", "001")]
    [InlineData("REAL", "9223372036854775807")]
    [InlineData("NUMERIC", "12345678901234567890.123456789")]
    [InlineData("DATETIME", "2026-11-20T10:00:00+08:00")]
    [InlineData("BOOLEAN", "TRUE")]
    [InlineData("CHARINT", "001")]
    public void AppendRejectsImplicitSqliteTextToNumberCoercion(string targetType, string identifier)
    {
        Execute("CREATE TABLE target (Id " + targetType + "); INSERT INTO target VALUES (7)");
        var book = Book([["Id"], [identifier]]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("target"), book,
            book.Sheets[0].Id, new(0, 0, 1, 0), mode: DatabaseWriteMode.Append));
        Assert.Equal(1L, Scalar("SELECT COUNT(*) FROM target"));
        Assert.Equal(7.0, Convert.ToDouble(Scalar("SELECT Id FROM target")));
    }

    [Fact]
    public void AppendRejectsMixedColumnCoercionButTextTargetPreservesLargeIdentifiers()
    {
        Execute("CREATE TABLE target (Id NUMERIC); CREATE TABLE identifiers (Id TEXT)");
        var mixed = Book([["Id"], [1.0], ["001"]]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("target"), mixed,
            mixed.Sheets[0].Id, new(0, 0, 2, 0), mode: DatabaseWriteMode.Append));
        Assert.Equal(0L, Scalar("SELECT COUNT(*) FROM target"));
        var text = Book([["Id"], ["001"], ["9223372036854775807"]]);
        DatabaseTransfer.ExportTable(Options, new("identifiers"), text, text.Sheets[0].Id, new(0, 0, 2, 0), mode: DatabaseWriteMode.Append);
        var imported = DatabaseTransfer.ImportTable(Options, new("identifiers"), false);
        Assert.Equal("001", imported.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Equal("9223372036854775807", imported.Workbook.Sheets[0].GetCell(1, 0)!.Value.TextValue);
    }

    [Fact]
    public void DatesUseWorkbookDateSystemAndTimezonesRemainText()
    {
        var id = SheetId.New();
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904,
            [new(id, "Dates", new Dictionary<CellPosition, CellSnapshot> { [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(0), 1)) })],
            styles: [new(), new() { NumberFormatId = 164, NumberFormatCode = "yyyy-mm-dd hh:mm:ss" }]);
        DatabaseTransfer.ExportTable(Options, new("dates"), book, id, new(0, 0, 0, 0), false);
        var result = DatabaseTransfer.ImportTable(Options, new("dates"), false);
        Assert.Equal(1462, result.Workbook.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.True(CellDisplayFormatter.IsDateFormat(result.Workbook.Styles[1], 1462));
        Execute("CREATE TABLE zoned (Stamp TIMESTAMP); INSERT INTO zoned VALUES ('2026-11-20T10:00:00+08:00')");
        var zoned = DatabaseTransfer.ImportTable(Options, new("zoned"), false);
        Assert.Equal("2026-11-20T10:00:00+08:00", zoned.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.NotEmpty(zoned.Notices);
    }

    [Fact]
    public void HighPrecisionIdentifiersAreTextAndBlobIsRejected()
    {
        Execute("CREATE TABLE precise (Id INTEGER, Literal TEXT); INSERT INTO precise VALUES (9223372036854775807, '00123')");
        var result = DatabaseTransfer.ImportTable(Options, new("precise"), false);
        Assert.Equal("9223372036854775807", result.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Equal("00123", result.Workbook.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        Assert.Single(result.Notices);
        Execute("CREATE TABLE binary (Payload BLOB); INSERT INTO binary VALUES (x'0001')");
        Assert.Throws<InvalidDataException>(() => DatabaseTransfer.ImportTable(Options, new("binary")));
    }

    [Fact]
    public void TypedReaderPreservesDecimalsDatesAndRejectsUnsupportedValues()
    {
        using var table = new DataTable(); table.Columns.Add("Decimal", typeof(decimal)); table.Columns.Add("Date", typeof(DateTime));
        table.Rows.Add(1234567890123456.123456m, new DateTime(2026, 11, 20));
        using var reader = table.CreateDataReader();
        var result = DatabaseTransfer.ReadResult(reader, "Data", false, default);
        Assert.Equal("1234567890123456.123456", result.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Equal(ExcelSerialDate.FromDateTime(new(2026, 11, 20), ExcelDateSystem.Windows1900).Serial,
            result.Workbook.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
    }

    [Fact]
    public void CancellationAndStaleFormulaValidationDoNotCreateDestination()
    {
        var book = Book([["x"]]);
        Assert.Throws<OperationCanceledException>(() => DatabaseTransfer.ExportTable(Options, new("data"), book,
            book.Sheets[0].Id, new(0, 0, 0, 0), false, cancellation: new(true)));
        Assert.False(File.Exists(DatabasePath));
        var id = SheetId.New();
        var formula = new WorkbookSnapshot(Guid.NewGuid(), 1, ExcelDateSystem.Windows1900,
            [new(id, "Sheet1", new Dictionary<CellPosition, CellSnapshot> { [new(0, 0)] = new(CellContent.FromFormula("=1+1"), new(CellValue.FromNumber(2), CalculationState.Fresh, 0, 0, 0)) })]);
        Assert.Throws<InvalidOperationException>(() => DatabaseTransfer.ExportTable(Options, new("data"), formula, id, new(0, 0, 0, 0), false));
        Assert.False(File.Exists(DatabasePath));
        Assert.Throws<OperationCanceledException>(() => DatabaseTransfer.ListTables(Options, new(true)));
    }

    [Fact]
    public void DuplicateHeadersAndOutOfBudgetRangeRejectBeforeFileCreation()
    {
        var book = Book([["x", "X"], [1.0, 2.0]]);
        Assert.Throws<ArgumentException>(() => DatabaseTransfer.ExportTable(Options, new("data"), book, book.Sheets[0].Id, new(0, 0, 1, 1)));
        Assert.Throws<ArgumentException>(() => DatabaseTransfer.ExportTable(Options, new("data"), book, book.Sheets[0].Id, new(0, 0, 1000, 1000), false));
        Assert.False(File.Exists(DatabasePath));
    }

    [Fact]
    public void OversizedSqliteTextFailsAndEmptyTableRetainsHeaders()
    {
        Execute("CREATE TABLE oversized (TextValue TEXT); INSERT INTO oversized VALUES (printf('%.*c', 32768, 'x'))");
        Assert.Throws<InvalidDataException>(() => DatabaseTransfer.ImportTable(Options, new("oversized")));
        Execute("CREATE TABLE empty (Id TEXT, Amount REAL)");
        var result = DatabaseTransfer.ImportTable(Options, new("empty"));
        Assert.Equal(0, result.Rows); Assert.Equal(2, result.Columns);
        Assert.Equal("Amount", result.Workbook.Sheets[0].GetCell(0, 1)!.Value.TextValue);
    }

    [Fact]
    public void MixedColumnsHaveExplicitInvariantTextRepresentation()
    {
        var book = Book([["Value"], [1.5], [true], ["001"]]);
        DatabaseTransfer.ExportTable(Options, new("mixed"), book, book.Sheets[0].Id, new(0, 0, 3, 0));
        var result = DatabaseTransfer.ImportTable(Options, new("mixed"), false);
        Assert.Equal("1.5", result.Workbook.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Equal("TRUE", result.Workbook.Sheets[0].GetCell(1, 0)!.Value.TextValue);
        Assert.Equal("001", result.Workbook.Sheets[0].GetCell(2, 0)!.Value.TextValue);
    }

    [Theory]
    [InlineData(DatabaseSqlDialect.Standard, "a\"b", "\"a\"\"b\"")]
    [InlineData(DatabaseSqlDialect.PostgreSql, "a\"b", "\"a\"\"b\"")]
    [InlineData(DatabaseSqlDialect.SqlServer, "a]b", "[a]]b]")]
    [InlineData(DatabaseSqlDialect.MySql, "a`b", "`a``b`")]
    public void OdbcIdentifiersAreQuotedAndCredentialsNeverAppearInOptionsText(DatabaseSqlDialect dialect, string identifier, string quoted)
    {
        var options = new DatabaseConnectionOptions(DatabaseProvider.Odbc, "DSN=test;PWD=private-secret;", dialect);
        Assert.Equal(quoted, DatabaseTransfer.QuoteIdentifier(identifier, options));
        Assert.DoesNotContain("private-secret", options.ToString());
        Assert.Throws<ArgumentException>(() => DatabaseTransfer.QuoteIdentifier("bad\0name", options));
    }

    private static WorkbookSnapshot Book(object?[][] rows)
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var r = 0; r < rows.Length; r++)
        for (var c = 0; c < rows[r].Length; c++)
        {
            var value = rows[r][c] switch { null => CellValue.Blank, string text => CellValue.FromText(text),
                bool boolean => CellValue.FromBoolean(boolean), double number => CellValue.FromNumber(number), _ => throw new InvalidOperationException() };
            cells.Add(new(r, c), new(CellContent.FromLiteral(value)));
        }
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "Data", cells)]);
    }

    private void Execute(string sql)
    {
        using var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = DatabasePath, Pooling = false }.ConnectionString);
        connection.Open(); using var command = connection.CreateCommand(); command.CommandText = sql; command.ExecuteNonQuery();
    }

    private object? Scalar(string sql)
    {
        using var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = DatabasePath, Pooling = false }.ConnectionString);
        connection.Open(); using var command = connection.CreateCommand(); command.CommandText = sql; return command.ExecuteScalar();
    }
}
