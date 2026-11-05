using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter(ScriptActionRequest request, CancellationToken cancellation)
{
    private const int MaxReadCells = 100_000;
    private const int MaxWriteCells = 20_000;
    private int _readCells;
    private int _writeCells;

    internal ModelReply Call(ModelCall call)
    {
        var model = request.Model ?? throw new ProtocolException("No live model for this request.");
        string? method = null;
        try
        {
            cancellation.ThrowIfCancellationRequested(); request.CancellationToken.ThrowIfCancellationRequested();
            if (call.OperationId != request.OperationId || call.Request.ValueKind != JsonValueKind.Object || call.Request.EnumerateObject().Count() != 2)
                throw new ProtocolException("Invalid model request.");
            method = call.Request.GetProperty("method").GetString();
            var arguments = call.Request.GetProperty("args");
            if (method is null || arguments.ValueKind != JsonValueKind.Array || arguments.GetArrayLength() > 3)
                throw new ProtocolException("Invalid model arguments.");
            var args = arguments.EnumerateArray().ToArray();
            // Files, storage and error utilities are internal host extensions with explicit
            // arities; Office object-model calls use the generated capability whitelist.
            if (!(method == "files" && args.Length == 0 || method == "errorCode" && args.Length == 2 || StorageCall(method, args.Length)) && !ScriptCapabilities.Current.AllowsHost(method, args.Length))
                throw new ModelFailure("InvalidArgument", "未知的工作簿模型方法或参数个数。");
            var view = model.View;
            object? value;
            if (StorageCall(method, args.Length)) value = CallStorage(method, args);
            else if (method.StartsWith("getApi:", StringComparison.Ordinal) || method.StartsWith("setApi:", StringComparison.Ordinal))
                value = CallApi(method, args);
            else if (method == "active")
            {
                Spend(1, write: false);
                var sheet = view.GetSheet(view.ActiveSheetId); var cell = sheet.View.ActiveCell;
                value = new { sheetId = sheet.Id.ToString(), cell = new CellRectangle(cell.Row, cell.Column, cell.Row, cell.Column).ToString(), ranges = model.Selection.Select(r => r.ToString()).ToArray() };
            }
            else if (method == "sheets")
            {
                Spend(1, write: false);
                value = view.Sheets.Select(s => new { id = s.Id.ToString(), name = s.Name }).ToArray();
            }
            else if (method == "sheet")
            {
                Spend(1, write: false);
                var name = args[0].GetString() ?? throw new ModelFailure("InvalidArgument", "缺少工作表名称或 ID。");
                var sheet = view.Sheets.FirstOrDefault(s => s.Id.ToString() == name || string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                value = sheet is null ? null : new { id = sheet.Id.ToString(), name = sheet.Name };
            }
            else if (method == "files")
            {
                Spend(1, write: false);
                value = request.Files.Entries.Values.Select(e => new VirtualEntry(e.Path,
                    e.IsDirectory ? null : Convert.ToBase64String(e.Data.AsSpan()), e.IsDirectory)).ToArray();
            }
            else
            {
                var sheetId = Guid.TryParseExact(args[0].GetString(), "D", out var id) && id != Guid.Empty
                    ? new SheetId(id) : throw new ModelFailure("InvalidArgument", "无效的工作表 ID。");
                var sheet = view.Sheets.FirstOrDefault(s => s.Id == sheetId)
                    ?? throw new ModelFailure("ItemNotFound", "找不到工作表。");
                var address = args[1].GetString();
                if (address is null || address.Length > 40) throw new ModelFailure("InvalidArgument", "无效的模型区域。");
                var range = CellRectangle.Parse(address);
                WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
                var rows = range.Bottom - range.Top + 1; var columns = range.Right - range.Left + 1;
                var cells = rows * columns;
                if (rows <= 0 || columns <= 0 || (long)cells > (method.StartsWith("set", StringComparison.Ordinal) ? MaxWriteCells : MaxReadCells))
                    throw new ModelFailure("BudgetExceeded", "单次模型请求超过单元格预算。");
                switch (method)
                {
                    case "getFormula":
                        Spend(cells, write: false);
                        var metadata = model.ReadCell(sheetId, new(range.Top, range.Left));
                        value = metadata?.Content.Kind == CellContentKind.Formula ? metadata.Content.FormulaText : ""; break;
                    case "getCell": case "errorCode":
                        Spend(1, write: false);
                        model.PrepareValueRead(sheetId, new(range.Top, range.Left, range.Top, range.Left), cancellation);
                        var cell = model.ReadCell(sheetId, new(range.Top, range.Left), false, calculated: true);
                        EnsureReadable(cell);
                        value = method == "errorCode" ? cell?.Value.Kind == CellValueKind.Error ? DisplayError(cell.Value.ErrorCode) : null : Scalar(cell?.Value ?? CellValue.Blank); break;
                    case "getValues":
                        Spend(cells, write: false);
                        model.PrepareValueRead(sheetId, range, cancellation);
                        var values = new object?[rows][];
                        for (var row = 0; row < rows; row++)
                        {
                            cancellation.ThrowIfCancellationRequested(); values[row] = new object?[columns];
                            for (var col = 0; col < columns; col++)
                            {
                                var item = model.ReadCell(sheetId, new(range.Top + row, range.Left + col), false, calculated: true);
                                EnsureReadable(item);
                                values[row][col] = Scalar(item?.Value ?? CellValue.Blank);
                            }
                        }
                        value = values; break;
                    case "getFill":
                        value = Uniform(sheet, range, style => Rgb(style.FillColor)); break;
                    case "clear":
                        Spend(cells, write: true);
                        model.SetValues(sheetId, range, Enumerable.Range(0, rows).Select(_ => (IReadOnlyList<CellValue>)Enumerable.Repeat(CellValue.Blank, columns).ToArray()).ToArray()); value = null; break;
                    case "setFill":
                        Spend(cells, write: true);
                        model.SetFill(sheetId, range, args[2].ValueKind == JsonValueKind.Null ? null : args[2].GetString()); value = null; break;
                    case "setValue":
                        Spend(cells, write: true);
                        var scalar = ReadScalar(args[2]);
                        model.SetValues(sheetId, range, Enumerable.Range(0, rows).Select(_ => (IReadOnlyList<CellValue>)Enumerable.Repeat(scalar, columns).ToArray()).ToArray()); value = null; break;
                    case "setValues":
                        Spend(cells, write: true);
                        var matrix = Matrix(args[2], rows, columns);
                        model.SetValues(sheetId, range, matrix.Select(r => (IReadOnlyList<CellValue>)r.Select(ReadScalar).ToArray()).ToArray()); value = null; break;
                    case "setFormula":
                        Spend(cells, write: true);
                        var formula = args[2].GetString() ?? throw new ModelFailure("InvalidArgument", "公式必须是字符串。");
                        model.SetFormulas(sheetId, range, Enumerable.Range(0, rows).Select(_ => (IReadOnlyList<string>)Enumerable.Repeat(formula, columns).ToArray()).ToArray()); value = null; break;
                    case "setFormulas":
                        Spend(cells, write: true);
                        var formulas = Matrix(args[2], rows, columns);
                        model.SetFormulas(sheetId, range, formulas.Select(r => (IReadOnlyList<string>)r.Select(v => v.GetString() ?? throw new ModelFailure("InvalidArgument", "公式必须是字符串。")).ToArray()).ToArray()); value = null; break;
                    default: throw new ModelFailure("InvalidArgument", "未知的工作簿模型方法。");
                }
            }
            cancellation.ThrowIfCancellationRequested();
            return new(call.OperationId, call.Sequence, true, WorkerProtocol.EncodePayload(value), null);
        }
        catch (ModelFailure failure)
        {
            if (IsWrite(method)) model.RejectWrite();
            return new(call.OperationId, call.Sequence, false, JsonSerializer.SerializeToElement<object?>(null), Bounded(failure.Message)) { Code = failure.Code };
        }
        catch (Exception error) when (error is ArgumentException or InvalidOperationException or IOException or FormatException or KeyNotFoundException or OverflowException or NotSupportedException)
        {
            if (IsWrite(method)) model.RejectWrite();
            var code = error switch
            {
                ScriptStorageQuotaException => "BudgetExceeded",
                KeyNotFoundException => "ItemNotFound",
                NotSupportedException => "UnsupportedApi",
                ArgumentException or FormatException or OverflowException => "InvalidArgument",
                _ => "InvalidReference"
            };
            return new(call.OperationId, call.Sequence, false, JsonSerializer.SerializeToElement<object?>(null), Bounded(error.Message)) { Code = code };
        }
    }

    private static bool IsWrite(string? method) =>
        method is not null && (method.StartsWith("set", StringComparison.Ordinal) || method == "clear");

    private void Spend(int cells, bool write)
    {
        if (write)
        {
            _writeCells = checked(_writeCells + cells);
            if (_writeCells > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "一次操作最多写入 20000 个单元格。");
        }
        else
        {
            _readCells = checked(_readCells + cells);
            if (_readCells > MaxReadCells) throw new ModelFailure("BudgetExceeded", "一次操作最多读取 100000 个单元格。");
        }
    }

    private static JsonElement[][] Matrix(JsonElement value, int rows, int columns)
    {
        if (value.ValueKind != JsonValueKind.Array) throw new ModelFailure("InvalidArgument", "需要二维数组。");
        var matrix = value.EnumerateArray().ToArray();
        if (matrix.Length != rows) throw new ModelFailure("InvalidArgument", "数据行数与目标区域不一致。");
        var result = new JsonElement[rows][];
        for (var row = 0; row < rows; row++)
        {
            if (matrix[row].ValueKind != JsonValueKind.Array || matrix[row].GetArrayLength() != columns)
                throw new ModelFailure("InvalidArgument", "数据尺寸与目标区域不一致。");
            result[row] = matrix[row].EnumerateArray().ToArray();
        }
        return result;
    }

    private static void EnsureReadable(CellSnapshot? cell)
    {
        var isFormula = cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null;
        if (isFormula && (cell!.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.Value is null))
            throw new ModelFailure("InvalidReference", "公式无法得到当前结果；可能包含不可用函数或受阻的依赖。");
    }

    private static string Bounded(string text) => text.Length > 1000 ? "工作簿模型请求失败。" : text;

    private static object? Scalar(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => "", CellValueKind.Number => value.NumberValue, CellValueKind.Text => value.TextValue,
        CellValueKind.Boolean => value.BooleanValue, CellValueKind.Error => DisplayError(value.ErrorCode),
        _ => throw new InvalidOperationException("Unsupported cell value.")
    };
    private static string DisplayError(CellErrorCode code) => code switch
    {
        CellErrorCode.Null => "#NULL!",
        CellErrorCode.DivisionByZero => "#DIV/0!",
        CellErrorCode.Value => "#VALUE!",
        CellErrorCode.Reference => "#REF!",
        CellErrorCode.Name => "#NAME?",
        CellErrorCode.Number => "#NUM!",
        CellErrorCode.NotAvailable => "#N/A",
        CellErrorCode.Cycle or CellErrorCode.Calculation => "#NUM!",
        CellErrorCode.Spill => "#SPILL!",
        _ => "#VALUE!"
    };
    private static CellValue ReadScalar(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String when value.GetString()!.Length <= WorkbookLimits.MaxCellTextLength => CellValue.FromText(value.GetString()!),
        JsonValueKind.Number when double.IsFinite(value.GetDouble()) => CellValue.FromNumber(value.GetDouble()),
        JsonValueKind.True => CellValue.FromBoolean(true), JsonValueKind.False => CellValue.FromBoolean(false),
        _ => throw new ModelFailure("InvalidArgument", "单元格值必须是有限数值、文本或布尔值。")
    };

    private sealed class ModelFailure(string code, string message) : Exception(message)
    {
        public string Code { get; } = code;
    }
}
