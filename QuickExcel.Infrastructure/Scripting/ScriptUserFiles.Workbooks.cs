using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;
using QuickExcel.Storage;
using static QuickExcel.Runtime.SelectedFileSystem;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptUserFiles
{
    private readonly FileReadBudget _workbookReadBudget = new(cancellation);
    private readonly ScriptModelAdapter.CellBudget _workbookCellBudget = new();

    private WorkbookDocument LoadWorkbook(OpenFile file, bool keep)
    {
        var path = file.Entry.Target.Path;
        file.Data.Position = 0;
        var snapshot = keep ? TypedWorkbookStorage.ReadSelectedWorkbook(file.Data, path, _workbookReadBudget)
            : new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new WorksheetSnapshot(SheetId.New(), "Sheet1")], textCompatibilityVersion: 2,
                scriptModules: TypedWorkbookStorage.IsXlsq(path) ? ScriptModuleSet.Single("", ScriptModuleSet.CurrentApiVersion) : null);
        // Even a read grant needs private calculation commits. The broker gates every write
        // using the selected entry, and this private document never enables workbook scripts.
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

    private ModelReply WorkbookModelCall(ModelCall call, JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Array || args.GetArrayLength() != 2) throw UserFileError.Invalid();
        var id = args[0].GetString() ?? throw UserFileError.Invalid();
        if (!_streams.TryGetValue(id, out var file) || file.WorkbookModel is null || file.Workbook is null)
            throw new UserFileError("InvalidReference", "Workbook handle expired.");
        if (file.Failed) throw new UserFileError("TransactionAborted", "Workbook draft was discarded.");
        var operation = args[1];
        if (operation.ValueKind != JsonValueKind.Object || operation.EnumerateObject().Count() != 2 ||
            !operation.TryGetProperty("method", out var name) || name.ValueKind != JsonValueKind.String ||
            !operation.TryGetProperty("args", out var arguments) || arguments.ValueKind != JsonValueKind.Array)
            throw UserFileError.Invalid();
        var method = name.GetString()!;
        if (!ScriptCapabilities.Current.AllowsSelectedHost(method, arguments.GetArrayLength()))
            throw new UserFileError("UnsupportedApi", "This operation is unavailable on a selected workbook.");
        var write = ScriptModelAdapter.IsWrite(method);
        try
        {
            var readsData = write ? method == "setApi:Range.copyFrom" : !IsWriteOnlyNavigation(method);
            if (write && !file.Entry.CanWrite || readsData && !file.Entry.CanRead)
                throw UserFileError.Denied();
            if (method.StartsWith("getApi:", StringComparison.Ordinal) || method.StartsWith("setApi:", StringComparison.Ordinal))
                Keys(arguments[0], "sheetId", "address", "side");
            var operandIndex = method switch { "setApi:Range.copyFrom" => 0, "getApi:Range.getRangeEdge" => 1, _ => -1 };
            if (operandIndex >= 0 && arguments[1].GetArrayLength() > operandIndex &&
                arguments[1][operandIndex] is { ValueKind: JsonValueKind.Object } operand &&
                (!operand.TryGetProperty("selectedWorkbook", out var owner) || owner.ValueKind != JsonValueKind.String || owner.GetString() != id))
                throw new UserFileError("InvalidReference", "Range belongs to another workbook.");
            var revision = file.Workbook.Session.Revision;
            var reply = file.WorkbookModel.Call(call with { Request = operation });
            if (!reply.Ok && write) file.Failed = true;
            file.Dirty |= file.Workbook.Session.Revision != revision;
            return reply;
        }
        catch
        {
            if (write) file.Failed = true;
            throw;
        }
    }

    // Write-only opens start with a new document. Return only navigation and format
    // descriptors needed to reach write methods, never a previously stored cell or style.
    private static bool IsWriteOnlyNavigation(string method) => method is
        "sheet" or "sheets" or "active" or "getApi:Workbook.getFirstWorksheet" or "getApi:Workbook.getLastWorksheet" or
        "getApi:Worksheet.getNext" or "getApi:Worksheet.getPrevious" or
        "getApi:RangeFormat.getFont" or "getApi:RangeFormat.getProtection" or
        "getApi:RangeFormat.getBorders" or "getApi:RangeFormat.getRangeBorder";
}
