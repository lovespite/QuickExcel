using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

internal sealed class ScriptModelAdapter(ScriptActionRequest request, CancellationToken cancellation)
{
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
            if (arguments.ValueKind != JsonValueKind.Array || arguments.GetArrayLength() > 3) throw new ProtocolException("Invalid model arguments.");
            var args = arguments.EnumerateArray().ToArray();
            int expected = method switch { "active" or "sheets" or "files" => 0, "sheet" => 1,
                "getCell" or "getFormula" or "snapshotCell" or "getValues" or "snapshotValues" or "getFill" => 2,
                "setFill" or "setValue" or "setValues" or "setFormula" or "setFormulas" => 3, _ => throw new ProtocolException("Unknown model method.") };
            if (args.Length != expected) throw new ProtocolException("Wrong model argument count.");
            var view = model.View;
            object? value;
            if (method == "active")
            {
                var sheet = view.GetSheet(view.ActiveSheetId); var cell = sheet.View.ActiveCell;
                value = new { sheetId = sheet.Id.ToString(), cell = new CellRectangle(cell.Row, cell.Column, cell.Row, cell.Column).ToString(), ranges = model.Selection.Select(r => r.ToString()).ToArray() };
            }
            else if (method == "sheets") value = view.Sheets.Select(s => new { id = s.Id.ToString(), name = s.Name }).ToArray();
            else if (method == "sheet")
            {
                var name = args[0].GetString() ?? throw new ProtocolException("Missing worksheet name or id.");
                var sheet = view.Sheets.FirstOrDefault(s => s.Id.ToString() == name || string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                value = sheet is null ? null : new { id = sheet.Id.ToString(), name = sheet.Name };
            }
            else if (method == "files") value = request.Files.Entries.Values.Select(e => new VirtualEntry(e.Path,
                e.IsDirectory ? null : Convert.ToBase64String(e.Data.AsSpan()), e.IsDirectory)).ToArray();
            else
            {
                var sheetId = new SheetId(Guid.ParseExact(args[0].GetString()!, "D"));
                var sheet = view.GetSheet(sheetId);
                var address = args[1].GetString();
                if (address is null || address.Length > 40) throw new ProtocolException("Invalid model range.");
                var range = CellRectangle.Parse(address);
                WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
                var rows = range.Bottom - range.Top + 1; var columns = range.Right - range.Left + 1;
                if (rows <= 0 || columns <= 0 || (long)rows * columns > (method.StartsWith("set", StringComparison.Ordinal) ? 20_000 : 100_000))
                    throw new ProtocolException("Model range exceeds its cell budget.");
                switch (method)
                {
                    case "getFormula":
                        var metadata = model.ReadCell(sheetId, new(range.Top, range.Left));
                        value = metadata?.Content.Kind == CellContentKind.Formula ? metadata.Content.FormulaText : ""; break;
                    case "getCell": case "snapshotCell":
                        if (method == "getCell") model.PrepareValueRead(sheetId, new(range.Top, range.Left, range.Top, range.Left), cancellation);
                        var cell = model.ReadCell(sheetId, new(range.Top, range.Left), method == "snapshotCell", calculated: method == "getCell");
                        var isFormula = cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null;
                        value = new { value = Scalar(cell?.Value ?? CellValue.Blank), formula = cell?.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText : null,
                            state = isFormula ? cell?.Calculation?.State.ToString() ?? "Dirty" : "literal",
                            hasValue = !isFormula || cell?.Calculation?.Value is not null }; break;
                    case "getValues": case "snapshotValues":
                        if (method == "getValues") model.PrepareValueRead(sheetId, range, cancellation);
                        var values = new object?[rows][];
                        for (var row = 0; row < rows; row++)
                        {
                            cancellation.ThrowIfCancellationRequested(); values[row] = new object?[columns];
                            for (var col = 0; col < columns; col++)
                            {
                                var item = model.ReadCell(sheetId, new(range.Top + row, range.Left + col), method == "snapshotValues", calculated: method == "getValues");
                                if (method == "getValues" && (item?.Content.Kind == CellContentKind.Formula || item?.SpillOwner is not null) &&
                                    (item.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || item.Calculation.Value is null))
                                    throw new InvalidOperationException("公式无法得到当前结果；可能包含不可用函数或受阻的依赖。");
                                values[row][col] = Scalar(item?.Value ?? CellValue.Blank);
                            }
                        }
                        value = values; break;
                    case "getFill":
                        var color = view.Styles[sheet.GetStyleId(range.Top, range.Left)].FillColor;
                        value = color is { Length: 8 } ? "#" + color[2..] : ""; break;
                    case "setFill": model.SetFill(sheetId, range, args[2].ValueKind == JsonValueKind.Null ? null : args[2].GetString()); value = null; break;
                    case "setValue":
                        var scalar = ReadScalar(args[2]);
                        model.SetValues(sheetId, range, Enumerable.Range(0, rows).Select(_ => (IReadOnlyList<CellValue>)Enumerable.Repeat(scalar, columns).ToArray()).ToArray()); value = null; break;
                    case "setValues":
                        model.SetValues(sheetId, range, args[2].EnumerateArray().Select(r => (IReadOnlyList<CellValue>)r.EnumerateArray().Select(ReadScalar).ToArray()).ToArray()); value = null; break;
                    case "setFormula":
                        var formula = args[2].GetString()!;
                        model.SetFormulas(sheetId, range, Enumerable.Range(0, rows).Select(_ => (IReadOnlyList<string>)Enumerable.Repeat(formula, columns).ToArray()).ToArray()); value = null; break;
                    case "setFormulas":
                        model.SetFormulas(sheetId, range, args[2].EnumerateArray().Select(r => (IReadOnlyList<string>)r.EnumerateArray().Select(v => v.GetString()!).ToArray()).ToArray()); value = null; break;
                    default: throw new ProtocolException("Unknown model method.");
                }
            }
            cancellation.ThrowIfCancellationRequested();
            return new(call.OperationId, call.Sequence, true, WorkerProtocol.EncodePayload(value), null);
        }
        catch (Exception error) when (error is ArgumentException or InvalidOperationException or IOException or FormatException or KeyNotFoundException or OverflowException)
        {
            if (method is null || method.StartsWith("set", StringComparison.Ordinal)) model.RejectWrite();
            return new(call.OperationId, call.Sequence, false, JsonSerializer.SerializeToElement<object?>(null),
                error.Message.Length > 1000 ? "工作簿模型请求失败。" : error.Message);
        }
    }

    private static object? Scalar(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => null, CellValueKind.Number => value.NumberValue, CellValueKind.Text => value.TextValue,
        CellValueKind.Boolean => value.BooleanValue, CellValueKind.Error => new { kind = "error", code = value.ErrorCode.ToString() },
        _ => throw new InvalidOperationException("Unsupported cell value.")
    };
    private static CellValue ReadScalar(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => CellValue.Blank,
        JsonValueKind.String when value.GetString()!.Length <= WorkbookLimits.MaxCellTextLength => CellValue.FromText(value.GetString()!),
        JsonValueKind.Number when double.IsFinite(value.GetDouble()) => CellValue.FromNumber(value.GetDouble()),
        JsonValueKind.True => CellValue.FromBoolean(true), JsonValueKind.False => CellValue.FromBoolean(false),
        JsonValueKind.Object when value.EnumerateObject().Count() == 2 && value.GetProperty("kind").GetString() == "error" &&
            Enum.TryParse<CellErrorCode>(value.GetProperty("code").GetString(), out var code) && Enum.IsDefined(code) => CellValue.FromError(code),
        _ => throw new ArgumentException("单元格值必须是有限数值、文本、布尔值、null 或有效的错误值。")
    };
}
