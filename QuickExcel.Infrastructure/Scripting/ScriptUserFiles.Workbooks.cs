using System.Collections.Immutable;
using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.ScriptContracts;
using QuickExcel.Storage;
using static QuickExcel.Runtime.SelectedFileSystem;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptUserFiles
{
    private readonly FileReadBudget _workbookReadBudget = new(cancellation);
    private long _readCells, _writtenCells;

    private WorkbookDocument LoadWorkbook(OpenFile file, bool keep)
    {
        var path = file.Entry.Target.Path;
        file.Data.Position = 0;
        var snapshot = keep ? TypedWorkbookStorage.ReadSelectedWorkbook(file.Data, path, _workbookReadBudget)
            : new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new WorksheetSnapshot(SheetId.New(), "Sheet1")], textCompatibilityVersion: 2,
                scriptModules: TypedWorkbookStorage.IsXlsq(path) ? ScriptModuleSet.Single("", ScriptModuleSet.CurrentApiVersion) : null);
        // This private document never enables scripts. Its calculator can evaluate standard formulas.
        return new WorkbookDocument(snapshot, new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory())));
    }

    private MemoryStream EncodeWorkbook(WorkbookSnapshot snapshot, Entry target, string? copyType = null)
    {
        var bytes = TypedWorkbookStorage.WriteSelectedWorkbook(snapshot, target.Target.Path, cancellation, copyType);
        if (Reserved() + bytes.Length > DraftLimit) throw UserFileError.Limit();
        return new MemoryStream(bytes, writable: false);
    }

    private readonly Dictionary<string, Func<ImmutableArray<VirtualFileEdit>, WorkbookSnapshot>> _pendingCopies = [];

    private string BeginSaveAs()
    {
        if (_pendingCopies.Count >= 16) throw UserFileError.Limit();
        var id = Id();
        _pendingCopies.Add(id, request.BeginSaveAs(cancellation));
        return id;
    }

    private void SaveAs(string targetId, JsonElement args)
    {
        if (args.GetArrayLength() != 4 || args[1].GetString() is not ("xlsx" or "xlsq") ||
            args[3].ValueKind != JsonValueKind.Array) throw UserFileError.Invalid();
        var sourceId = args[2].GetString() ?? throw UserFileError.Invalid();
        _pendingCopies.Remove(sourceId, out var capture);
        if (!_entries.TryGetValue(targetId, out var target) || target.Directory)
            throw new UserFileError("InvalidReference", "Expected a current file handle.");
        if (!target.CanWrite) throw UserFileError.Denied();
        using var chain = new Chain(target.Target.Path, target.Target.Parents);
        Verify(target, chain.Observe());
        RequireWriterAvailable(target);
        WorkbookSnapshot snapshot;
        if (capture is not null)
        {
            if (args[3].GetArrayLength() > VirtualFileProtocol.MaxEntries) throw UserFileError.Limit();
            var edits = args[3].Deserialize(WorkerProtocol.TypeInfo<VirtualEdit[]>()) ?? throw UserFileError.Invalid();
            VirtualFileProtocol.ValidateEdits(edits);
            if (edits.Any(e => !e.Path.StartsWith("//xlsq/", StringComparison.Ordinal))) throw UserFileError.Invalid();
            snapshot = capture(ScriptWorkerActionExecutor.DecodeFileEdits(edits));
        }
        else
        {
            if (args[3].GetArrayLength() != 0) throw UserFileError.Invalid();
            if (!_streams.TryGetValue(sourceId, out var source) || source.Workbook is null)
                throw new UserFileError("InvalidReference", "Workbook handle expired.");
            if (source.Failed) throw new UserFileError("TransactionAborted", "Workbook draft was discarded.");
            if (SameTarget(source.Entry, target)) throw UserFileError.Conflict();
            snapshot = source.Workbook.Snapshot;
        }
        using var encoded = EncodeWorkbook(snapshot, target, args[1].GetString());
        PublishDraft(target, chain, encoded);
    }

    private JsonNode? WorkbookCall(string id, JsonElement args)
    {
        if (!_streams.TryGetValue(id, out var file) || file.Workbook is not { } book)
            throw new UserFileError("InvalidReference", "Workbook handle expired.");
        if (file.Failed) throw new UserFileError("TransactionAborted", "Workbook draft was discarded.");
        var method = args[1].GetString();
        var write = method is "setValues" or "setFormulas" or "addWorksheet";
        if (write ? !file.Entry.CanWrite : !file.Entry.CanRead) throw UserFileError.Denied();
        try
        {
            if (method == "getWorksheetNames")
            {
                var names = new JsonArray();
                foreach (var item in book.Snapshot.Sheets) names.Add(item.Name);
                return names;
            }
            var name = args[2].GetString() ?? throw UserFileError.Invalid();
            if (method == "addWorksheet")
            {
                book.Session.Execute(new AddSheetCommand(name));
                file.Dirty = true;
                return null;
            }
            var sheet = book.Snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                ?? throw new UserFileError("ItemNotFound", "Worksheet does not exist.");
            if (method == "getUsedRange")
            {
                if (sheet.Cells.IsEmpty) return null;
                SpendWorkbookCells(sheet.Cells.Count, false);
                var cells = sheet.Cells.Keys;
                return new CellRectangle(cells.Min(p => p.Row), cells.Min(p => p.Column),
                    cells.Max(p => p.Row), cells.Max(p => p.Column)).ToString();
            }
            if (args[3].GetString() is not { Length: > 0 and <= 40 } address) throw UserFileError.Invalid();
            var range = CellRectangle.Parse(address);
            SpendWorkbookCells(range.Count, write);
            if (method is "getValues" or "getFormulas")
            {
                if (method == "getValues" && file.CalculatedRevision != book.Snapshot.Revision)
                {
                    if (!book.Calculate(cancellation)) throw new UserFileError("InvalidReference", "Calculation was superseded.");
                    file.CalculatedRevision = book.Snapshot.Revision;
                    sheet = book.Snapshot.GetSheet(sheet.Id);
                }
                return ReadRange(sheet, range, method == "getFormulas");
            }
            if (method is not ("setValues" or "setFormulas")) throw UserFileError.Invalid();
            var command = RangeEdit(sheet, range, args[4], method == "setFormulas");
            file.Dirty |= book.Session.Execute(command).Changed;
            return null;
        }
        catch { if (write) file.Failed = true; throw; }
    }

    private void SpendWorkbookCells(long count, bool write)
    {
        if (write ? (_writtenCells += count) > 20_000 : (_readCells += count) > 100_000)
            throw UserFileError.Limit();
    }

    private JsonArray ReadRange(WorksheetSnapshot sheet, CellRectangle range, bool formulas)
    {
        var rows = new JsonArray();
        long size = 2;
        for (var row = range.Top; row <= range.Bottom; row++)
        {
            cancellation.ThrowIfCancellationRequested();
            var items = new JsonArray();
            size += 3;
            for (var column = range.Left; column <= range.Right; column++)
            {
                var cell = sheet.GetCell(row, column);
                if (!formulas && (cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) &&
                    (cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.Value is null))
                    throw new UserFileError("InvalidReference", "A formula has no current result with workbook scripts disabled.");
                var value = cell?.Value ?? CellValue.Blank;
                JsonNode? item = formulas ? (JsonNode?)(cell?.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText : null)
                    : value.Kind switch
                    {
                        CellValueKind.Blank => null,
                        CellValueKind.Boolean => value.BooleanValue,
                        CellValueKind.Number => value.NumberValue,
                        _ => CellInput.DisplayText(value)
                    };
                size += item is JsonValue scalar && scalar.TryGetValue<string>(out var text)
                    ? JsonEncodedText.Encode(text).EncodedUtf8Bytes.Length + 3 : 32;
                if (size > WorkerProtocol.MaxFrameBytes) throw UserFileError.Limit();
                items.Add(item);
            }
            rows.Add(items);
        }
        return rows;
    }

    private SetScriptCellsCommand RangeEdit(WorksheetSnapshot sheet, CellRectangle range, JsonElement matrix, bool formulas)
    {
        if (matrix.GetArrayLength() != range.Rows) throw UserFileError.Invalid();
        var edits = new List<CellEdit>((int)range.Count);
        var row = range.Top;
        foreach (var values in matrix.EnumerateArray())
        {
            cancellation.ThrowIfCancellationRequested();
            if (values.GetArrayLength() != range.Columns) throw UserFileError.Invalid();
            var column = range.Left;
            foreach (var item in values.EnumerateArray())
            {
                var style = sheet.GetStyleId(row, column);
                CellContent content;
                if (formulas)
                {
                    if (item.GetString() is not { Length: >= 2 and <= 8192 } formula || formula[0] != '=') throw UserFileError.Invalid();
                    _ = FormulaSyntax.Parse(formula);
                    content = CellContent.FromFormula(formula, style);
                }
                else
                {
                    var value = item.ValueKind switch
                    {
                        JsonValueKind.Null => CellValue.Blank,
                        JsonValueKind.True or JsonValueKind.False => CellValue.FromBoolean(item.GetBoolean()),
                        JsonValueKind.Number when double.IsFinite(item.GetDouble()) => CellValue.FromNumber(item.GetDouble()),
                        JsonValueKind.String => CellValue.FromText(item.GetString()!),
                        _ => throw UserFileError.Invalid()
                    };
                    content = CellContent.FromLiteral(value, style);
                }
                edits.Add(new(new(sheet.Id, row, column++), content));
            }
            row++;
        }
        return new SetScriptCellsCommand(edits);
    }
}
