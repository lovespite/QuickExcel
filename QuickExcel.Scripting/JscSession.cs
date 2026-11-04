using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace QuickExcel.Scripting;

/// <summary>A serialized JSC context. Production callers must run in the isolated worker.</summary>
public sealed partial class JscSession : IDisposable
{
    private const string NativeLibrary = "JavaScriptCore";
    public const string EngineIdentity = "JavaScriptCore/2.54.2/b95c91b1/qx-sync-2";
    public const int MaxMessageBytes = 16 * 1024 * 1024;
    private readonly object _gate = new();
    private IntPtr _session;
    private readonly HostCallback _hostCallback;
    private IntPtr _hostResponse;
    public Func<string, string>? HostRequest { get; set; }

    // The bridge object is held by native code, not exposed on globalThis. Workbook
    // code receives only registration and console APIs, with no host CLR objects.
    private static readonly string Bootstrap = """
        (() => {
            const nativeModelCall = globalThis.__qxModelCall;
            delete globalThis.__qxModelCall;
            const registrationRole = globalThis.__qxRegistrationRole;
            delete globalThis.__qxRegistrationRole;
            /*VIRTUAL_FILES*/
            const registry = Object.create(null);
            const descriptors = Object.create(null);
            const ownKeys = Object.keys;
            const freeze = Object.freeze;
            const makeFunction = Function;
            const upper = Function.prototype.call.bind(String.prototype.toUpperCase);
            const validName = Function.prototype.call.bind(RegExp.prototype.test, /^[A-Z]+$/);
            const validV2Name = Function.prototype.call.bind(RegExp.prototype.test, /^JS\.[A-Z_][A-Z0-9_]*(\.[A-Z_][A-Z0-9_]*)*$/);
            const isArray = Array.isArray;
            const deepFreeze = value => {
                if (value && typeof value === 'object') { for (const key of ownKeys(value)) deepFreeze(value[key]); freeze(value); }
                return value;
            };
            const isFiniteNumber = Number.isFinite;
            const TypeErrorCtor = TypeError;
            /*WORKBOOK_MODEL*/
            let loaded = false;
            let actionMode = false;
            let sharedMode = false;
            let actionInvoked = false;
            const actions = Object.create(null);
            const actionInfo = Object.create(null);
            const eventHandlers = Object.create(null);
            /*EVENT_API*/
            const logs = [];
            function display(value, depth = 0, seen = new Set()) {
                if (typeof value === 'string') return value.slice(0, 4000);
                if (typeof value === 'bigint') return String(value) + 'n';
                if (value === null || typeof value !== 'object') return String(value);
                if (seen.has(value)) return '[Circular]';
                if (depth >= 3) return isArray(value) ? '[Array]' : '[Object]';
                seen.add(value);
                try {
                    if (value instanceof Error) return String(value.stack || value);
                    const parts = ownKeys(value).slice(0, 30).map(key => {
                        const descriptor = Object.getOwnPropertyDescriptor(value, key);
                        const text = descriptor && 'value' in descriptor ? display(descriptor.value, depth + 1, seen) : '[Getter]';
                        return isArray(value) ? text : key + ': ' + text;
                    });
                    return ((isArray(value) ? '[' : '{') + parts.join(', ') + (ownKeys(value).length > 30 ? ', …' : '') + (isArray(value) ? ']' : '}')).slice(0, 4000);
                } catch { return '[Uninspectable]'; }
                finally { seen.delete(value); }
            }
            const logger = (prefix, values) => {
                if (logs.length < 100) logs.push((prefix + values.slice(0, 30).map(value => display(value)).join(' ')).slice(0, 4000));
            };
            const console = freeze({ log: (...v) => logger('', v), info: (...v) => logger('[info] ', v),
                warn: (...v) => logger('[warn] ', v), error: (...v) => logger('[error] ', v), debug: (...v) => logger('[debug] ', v) });
            const quickExcel = freeze({ ...workbookApi, functions: freeze({
                register(name, fn) {
                    if (actionMode || sharedMode && !(registrationRole() & 1)) return;
                    if (loaded) throw new TypeErrorCtor('Functions can only be registered while loading the script.');
                    let descriptor = { apiVersion: 1, minArguments: 0, maxArguments: 255, returnType: 'any', profile: 'conservative', parameters: [] };
                    if (typeof name === 'object' && name !== null) {
                        const input = name;
                        if (input.apiVersion !== 2 || typeof input.name !== 'string') throw new TypeErrorCtor('A v2 descriptor requires apiVersion: 2 and a name.');
                        name = upper(input.name);
                        if (!validV2Name(name)) throw new TypeErrorCtor('V2 functions must use the JS namespace.');
                        const parameters = input.parameters ?? [];
                        if (!isArray(parameters) || parameters.length > 255) throw new TypeErrorCtor('Invalid parameter descriptors.');
                        const types = ['any', 'number', 'text', 'boolean', 'blank', 'range'];
                        for (const p of parameters)
                            if (!p || typeof p.name !== 'string' || !types.includes(p.type ?? 'any')) throw new TypeErrorCtor('Invalid parameter type.');
                        descriptor = { name, apiVersion: 2, minArguments: parameters.length, maxArguments: parameters.length,
                            returnType: input.returns ?? 'any', profile: input.profile ?? 'conservative',
                            parameters: parameters.map(p => ({ name: p.name, type: p.type ?? 'any' })) };
                        if (!['any', 'number', 'text', 'boolean', 'blank'].includes(descriptor.returnType) || descriptor.profile !== 'conservative')
                            throw new TypeErrorCtor('M3 supports scalar returns and conservative execution only.');
                    } else {
                        if (typeof name !== 'string') throw new TypeErrorCtor('Function name must be a string.');
                        name = upper(name);
                        if (!validName(name)) throw new TypeErrorCtor('V1 function names must contain letters A-Z only.');
                    }
                    descriptor.name = name;

                    if (typeof fn !== 'function') throw new TypeErrorCtor('A registered formula must be a function.');
                    if (registry[name]) throw new TypeErrorCtor('Duplicate function: ' + name);
                    registry[name] = fn;
                    descriptors[name] = deepFreeze(descriptor);
                }
            }), actions: freeze({
                register(input, fn) {
                    if (!actionMode || sharedMode && !(registrationRole() & 2)) return;
                    if (loaded) throw new TypeErrorCtor('Actions can only be registered while loading.');
                    const name = typeof input === 'string' ? input : input?.name;
                    const description = typeof input === 'string' ? '' : (input?.description ?? '');
                    if (typeof name !== 'string' || !name.trim() || name.length > 128 || /[\x00-\x1f\x7f-\x9f]/.test(name) ||
                        name.startsWith('@event/') || typeof description !== 'string' || description.length > 1000 || typeof fn !== 'function' || actions[name] || ownKeys(actions).length >= 256)
                        throw new TypeErrorCtor('Invalid or duplicate action registration.');
                    actions[name] = fn;
                    actionInfo[name] = freeze({ name, description });
                }
            }), events: createEvents() });
            // One frozen API object and one registry. A global alias also lets
            // legacy function bodies keep their own local `const $` bindings.
            Object.defineProperty(globalThis, '$', { value: quickExcel });
            function runAction(name, request) {
                if (!actionMode || !loaded || actionInvoked || !actions[name]) throw new TypeErrorCtor('Action is unavailable.');
                actionInvoked = true;
                activeModel = request.liveModel === true;
                modelWriteFailed = false;
                const data = deepFreeze(request.snapshot);
                const sheets = new Map();
                for (const sheet of data.sheets) {
                    const cells = new Map();
                    for (const c of sheet.cells) cells.set(c.row + ':' + c.column, c);
                    sheets.set(sheet.id, cells);
                }
                let reads = 0, active = true, invalidPatch = false;
                const edits = [], targets = new Set();
                let files = null;
                function fileContext() {
                    if(!active) throw new TypeErrorCtor('Expired file context.');
                    return files ??= createFs(activeModel ? model('files') : (request.files ?? []),
                        ()=>{if(!active) throw new TypeErrorCtor('Expired file context.');}, ()=>{invalidPatch=true;});
                }
                function range(id, address) {
                    if (!active || !sheets.has(id) || typeof address !== 'string' || address.length > 40) throw new TypeErrorCtor('Invalid sheet or range.');
                    const parts = address.split(':');
                    if (parts.length > 2) throw new TypeErrorCtor('Expected A1 or rectangular A1:B2 range.');
                    function point(text) {
                        const m = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(text);
                        if (!m) throw new TypeErrorCtor('Invalid A1 address.');
                        let col=0; for (const c of m[1].toUpperCase()) col=col*26+c.charCodeAt(0)-64;
                        const row=Number(m[2]);
                        if (!Number.isSafeInteger(row) || row>100000 || col>1024) throw new TypeErrorCtor('Address out of bounds.');
                        return [row-1,col-1];
                    }
                    const a=point(parts[0]), b=point(parts[1] ?? parts[0]);
                    if (b[0]<a[0] || b[1]<a[1]) throw new TypeErrorCtor('Reversed range.');
                    return {row:a[0],column:a[1],rows:b[0]-a[0]+1,columns:b[1]-a[1]+1};
                }
                function value(v) { return v.kind === 'error' ? freeze({kind:'error',code:v.value}) : v.value; }
                function cell(id,row,col) {
                    const c=sheets.get(id).get(row+':'+col);
                    return c ? freeze({value:value(c.value),formula:c.formula,state:c.state,hasValue:c.hasValue})
                        : freeze({value:null,formula:null,state:'literal',hasValue:true});
                }
                function readBudget(count) { reads+=count; if(reads>100000) throw new TypeErrorCtor('Snapshot read budget exceeded.'); }
                const snapshot=freeze({baseRevision:data.baseRevision,dateSystem:data.dateSystem,
                    sheets:freeze(data.sheets.map(s=>freeze({id:s.id,name:s.name}))),
                    getCell(id,address) { const r=range(id,address); if(r.rows*r.columns!==1) throw new TypeErrorCtor('Expected one cell.'); readBudget(1); return activeModel ? deepFreeze(model('snapshotCell',id,address)) : cell(id,r.row,r.column); },
                    getRange(id,address) { const r=range(id,address); readBudget(r.rows*r.columns); return activeModel ? deepFreeze(model('snapshotValues',id,address)) : freeze(Array.from({length:r.rows},(_,i)=>freeze(Array.from({length:r.columns},(_,j)=>cell(id,r.row+i,r.column+j).value)))); }
                });
                function scalar(v) {
                    if(v===null) return {kind:'blank',value:null};
                    if(typeof v==='string' && v.length<=32767) return {kind:'text',value:v};
                    if(typeof v==='number' && isFiniteNumber(v)) return {kind:'number',value:v};
                    if(typeof v==='boolean') return {kind:'boolean',value:v};
                    if(v && typeof v==='object' && v.kind==='error' && typeof v.code==='string') return {kind:'error',value:v.code};
                    throw new TypeErrorCtor('Unsupported patch value.');
                }
                function write(id,address,values,formulas) {
                    try { writeChecked(id,address,values,formulas); }
                    catch(error) { invalidPatch=true; throw error; }
                }
                function writeChecked(id,address,values,formulas) {
                    const r=range(id,address);
                    if(r.rows*r.columns+edits.length>20000 || !isArray(values) || values.length!==r.rows) throw new TypeErrorCtor('Patch size or shape is invalid.');
                    for(let i=0;i<r.rows;i++) {
                        if(!isArray(values[i]) || values[i].length!==r.columns) throw new TypeErrorCtor('Patch shape is invalid.');
                        for(let j=0;j<r.columns;j++) {
                            const key=id+':'+(r.row+i)+':'+(r.column+j), v=values[i][j];
                            if(targets.has(key)) throw new TypeErrorCtor('Overlapping patch writes.');
                            if(formulas && (typeof v!=='string' || v.length<2 || v.length>8192 || v[0]!=='=')) throw new TypeErrorCtor('Invalid formula.');
                            targets.add(key);
                            edits.push({sheetId:id,row:r.row+i,column:r.column+j,value:formulas?null:scalar(v),formula:formulas?v:null});
                        }
                    }
                    if(activeModel) model(formulas?'setFormulas':'setValues',id,address,values);
                }
                const patch=freeze({setValues:(id,r,v)=>write(id,r,v,false),setFormulas:(id,r,v)=>write(id,r,v,true),
                    setValue:(id,a,v)=>write(id,a,[[v]],false),setFormula:(id,a,v)=>write(id,a,[[v]],true)});
                try {
                    const returned=actions[name](freeze({snapshot,patch,get fs(){return fileContext().fs;},workbook:quickExcel,event:request.event ? deepFreeze(request.event) : null}));
                    if(invalidPatch || modelWriteFailed) throw new TypeErrorCtor('An invalid write rejected the entire patch.');
                    if(returned!==undefined && returned!==null) throw new TypeErrorCtor('Actions must complete synchronously without a return value.');
                    return {operationId:request.operationId,baseRevision:data.baseRevision,name,edits:activeModel?[]:edits,fileEdits:files?.edits ?? []};
                } finally { active=false; activeModel=false; }
            }
            return freeze({
                load(source) {
                    if (loaded) throw new TypeErrorCtor('Script has already been loaded.');
                    makeFunction('quickExcel', 'console', '"use strict";\n' + source)(quickExcel, console);
                    loaded = true;
                    return ownKeys(registry);
                },
                prepareModules(useActions, share) {
                    if (loaded) throw new TypeErrorCtor('Script has already been loaded.');
                    actionMode = useActions === true;
                    sharedMode = share === true;
                    Object.defineProperty(globalThis, 'quickExcel', { value: quickExcel });
                    Object.defineProperty(globalThis, 'console', { value: console });
                },
                finishModules() {
                    loaded = true;
                    return ownKeys(registry);
                },
                invoke(name, args) {
                    const fn = registry[name];
                    if (!fn) throw new TypeErrorCtor('Unknown function: ' + name);
                    const descriptor = descriptors[name];
                    if (descriptor.apiVersion === 2) {
                        if (args.length !== descriptor.parameters.length) throw new TypeErrorCtor('Argument count disagrees with descriptor.');
                        for (let i = 0; i < args.length; i++) {
                            const value = args[i];
                            const type = value === null ? 'blank' : typeof value === 'string' ? 'text'
                                : typeof value === 'object' && value.kind === 'range' ? 'range' : typeof value;
                            if (descriptor.parameters[i].type !== 'any' && descriptor.parameters[i].type !== type)
                                throw new TypeErrorCtor('Argument type disagrees with descriptor: ' + descriptor.parameters[i].name);
                        }
                        deepFreeze(args);
                    }
                    const value = fn(...args);
                    if (value === null || value === undefined) return null;
                    if (typeof value === 'string' || typeof value === 'boolean') return value;
                    if (typeof value === 'number' && isFiniteNumber(value)) return value;
                    throw new TypeErrorCtor('Formula functions must return a finite number, string, boolean, or null; objects and promises are not supported.');
                },
                descriptors() { return ownKeys(descriptors).map(name => descriptors[name]); },
                actionDescriptors() { return ownKeys(actionInfo).map(name => actionInfo[name]); },
                beginRound() { logs.length=0; actionInvoked=false; activeModel=false; modelWriteFailed=false; },
                runAction,
                logs() { return logs.slice(); }
            });
        })()
        """.Replace("/*VIRTUAL_FILES*/", VirtualFileBootstrap, StringComparison.Ordinal)
            .Replace("/*EVENT_API*/", EventApiBootstrap, StringComparison.Ordinal)
            .Replace("/*WORKBOOK_MODEL*/", WorkbookModelBootstrap, StringComparison.Ordinal);

    public IReadOnlyList<string> FunctionNames { get; }
    public IReadOnlyList<JsFunctionDescriptor> FunctionDescriptors { get; }
    public IReadOnlyList<JsActionDescriptor> ActionDescriptors { get; }

    public bool LastOperationInterrupted
    {
        get { lock (_gate) { ObjectDisposedException.ThrowIf(_session == IntPtr.Zero, this); return NativeWasInterrupted(_session) != 0; } }
    }

    public bool LastOperationInvalidated
    {
        get
        {
            lock (_gate)
            {
                ObjectDisposedException.ThrowIf(_session == IntPtr.Zero, this);
                return NativeIsUsable(_session) == 0;
            }
        }
    }

    public JscSession(string source) : this(source, null, null, false, null) { }

    public JscSession(IReadOnlyDictionary<string, string> modules, IReadOnlyList<string> entries, bool actions = false, IReadOnlyDictionary<string, int>? roles = null)
        : this(null, modules, entries, actions, roles) { }

    private JscSession(string? source, IReadOnlyDictionary<string, string>? modules, IReadOnlyList<string>? entries, bool actions, IReadOnlyDictionary<string, int>? roles)
    {
        _hostCallback = HostCall;
        if (modules is null) ArgumentNullException.ThrowIfNull(source);
        else
        {
            if (modules.Count is < 1 or > 64 || entries is null || entries.Count is < 1 or > 64 ||
                entries.Distinct(StringComparer.Ordinal).Count() != entries.Count || entries.Any(id => id is null || !modules.ContainsKey(id)) ||
                modules.Any(m => string.IsNullOrEmpty(m.Key) || m.Key.Length > 128 || m.Key is "." or ".." ||
                    m.Key.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')) || m.Value is null) ||
                modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value)) > 1024 * 1024)
                throw new ArgumentException("Invalid package module mapping or source budget.");
        }
        if (!Environment.Is64BitProcess || !(OperatingSystem.IsWindows() || OperatingSystem.IsLinux()))
            throw new PlatformNotSupportedException("QuickExcel scripting requires Windows x64 or qualified Linux x64 isolation.");
        if (Marshal.PtrToStringUTF8(NativeEngineIdentity()) != EngineIdentity)
            throw new InvalidOperationException("Native JSC build identity does not match this worker.");
        _session = NativeCreate(Bootstrap, out var error);
        if (_session == IntPtr.Zero) throw ReadException(error);
        NativeSetHostCall(_session, _hostCallback);
        try
        {
            string result;
            if (modules is null) result = Call("load", new object?[] { source });
            else
            {
                Call("prepareModules", [actions, roles is not null]);
                var json = JsonSerializer.Serialize(new { modules = modules.Select(m => new { id = m.Key, source = m.Value, roles = roles is null ? 3 : roles.GetValueOrDefault(m.Key) }), entries });
                if (Encoding.UTF8.GetByteCount(json) > MaxMessageBytes) throw new ArgumentException("Module payload exceeds 16 MiB UTF-8.");
                if (NativeLoadModules(_session, json, out var moduleError) == 0) throw ReadException(moduleError);
                result = Call("finishModules", []);
            }
            FunctionNames = Array.AsReadOnly(JsonSerializer.Deserialize<string[]>(result) ?? []);
            FunctionDescriptors = Array.AsReadOnly(JsonSerializer.Deserialize<JsFunctionDescriptor[]>(Call("descriptors", []),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? []);
            ActionDescriptors = Array.AsReadOnly(JsonSerializer.Deserialize<JsActionDescriptor[]>(Call("actionDescriptors", []),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? []);
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    /// <summary>Calls a registered formula with scalar or array arguments.</summary>
    public object? Invoke(string functionName, IReadOnlyList<object> arguments)
    {
        ArgumentNullException.ThrowIfNull(functionName);
        ArgumentNullException.ThrowIfNull(arguments);
        lock (_gate)
        {
            string result = Call("invoke", new object?[] { functionName.ToUpperInvariant(), arguments });
            using var document = JsonDocument.Parse(result);
            var value = document.RootElement;
            return value.ValueKind switch
            {
                JsonValueKind.Null => null,
                JsonValueKind.String => value.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Number => value.GetDouble(),
                _ => throw new InvalidOperationException("JavaScript returned an unsupported formula value.")
            };
        }
    }

    public IReadOnlyList<string> GetLogs()
    {
        lock (_gate)
            return JsonSerializer.Deserialize<string[]>(Call("logs", [])) ?? [];
    }

    public void BeginRound() { lock (_gate) Call("beginRound", []); }

    public JsonElement RunAction(string name, object request)
    {
        lock (_gate)
        {
            using var json = JsonDocument.Parse(Call("runAction", [name, request]));
            return json.RootElement.Clone();
        }
    }

    private string Call(string method, object?[] arguments)
    {
        ObjectDisposedException.ThrowIf(_session == IntPtr.Zero, this);
        var json = JsonSerializer.Serialize(arguments);
        if (Encoding.UTF8.GetByteCount(json) > MaxMessageBytes)
            throw new InvalidOperationException("JavaScript arguments exceed the 16 MiB UTF-8 input limit.");
        IntPtr result; IntPtr error;
        try { result = NativeCall(_session, method, json, out error); }
        finally { if (_hostResponse != IntPtr.Zero) { Marshal.FreeCoTaskMem(_hostResponse); _hostResponse = IntPtr.Zero; } }
        if (result == IntPtr.Zero) throw ReadException(error);
        try { return Marshal.PtrToStringUTF8(result) ?? "null"; }
        finally { NativeFreeString(result); }
    }

    private IntPtr HostCall(string request)
    {
        // Exceptions must never cross the reverse P/Invoke boundary.
        string response;
        try { response = HostRequest?.Invoke(request) ?? throw new InvalidOperationException("此执行阶段没有工作簿模型。"); }
        catch (Exception error) { response = JsonSerializer.Serialize(new { ok = false, error = error.Message.Length <= 4000 ? error.Message : "模型请求失败。" }); }
        if (Encoding.UTF8.GetByteCount(response) > MaxMessageBytes) response = "{\"ok\":false,\"error\":\"模型响应超过 16 MiB。\"}";
        if (_hostResponse != IntPtr.Zero) Marshal.FreeCoTaskMem(_hostResponse);
        return _hostResponse = Marshal.StringToCoTaskMemUTF8(response);
    }

    private static InvalidOperationException ReadException(IntPtr error)
    {
        try { return new InvalidOperationException(Marshal.PtrToStringUTF8(error) ?? "JavaScript execution failed."); }
        finally { if (error != IntPtr.Zero) NativeFreeString(error); }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_session == IntPtr.Zero) return;
            NativeDestroy(_session);
            _session = IntPtr.Zero;
        }
        GC.SuppressFinalize(this);
    }

    ~JscSession() => Dispose();

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate IntPtr HostCallback([MarshalAs(UnmanagedType.LPUTF8Str)] string request);
    [DllImport(NativeLibrary, EntryPoint = "qx_set_host_call", CallingConvention = CallingConvention.Cdecl)]
    private static extern void NativeSetHostCall(IntPtr session, HostCallback callback);

    [DllImport(NativeLibrary, EntryPoint = "qx_create", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr NativeCreate([MarshalAs(UnmanagedType.LPUTF8Str)] string bootstrap, out IntPtr error);
    [DllImport(NativeLibrary, EntryPoint = "qx_call", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr NativeCall(IntPtr session, [MarshalAs(UnmanagedType.LPUTF8Str)] string method, [MarshalAs(UnmanagedType.LPUTF8Str)] string argumentsJson, out IntPtr error);
    [DllImport(NativeLibrary, EntryPoint = "qx_load_modules", CallingConvention = CallingConvention.Cdecl)]
    private static extern int NativeLoadModules(IntPtr session, [MarshalAs(UnmanagedType.LPUTF8Str)] string modulesJson, out IntPtr error);
    [DllImport(NativeLibrary, EntryPoint = "qx_destroy", CallingConvention = CallingConvention.Cdecl)]
    private static extern void NativeDestroy(IntPtr session);
    [DllImport(NativeLibrary, EntryPoint = "qx_free_string", CallingConvention = CallingConvention.Cdecl)]
    private static extern void NativeFreeString(IntPtr text);
    [DllImport(NativeLibrary, EntryPoint = "qx_was_interrupted", CallingConvention = CallingConvention.Cdecl)]
    private static extern int NativeWasInterrupted(IntPtr session);
    [DllImport(NativeLibrary, EntryPoint = "qx_is_usable", CallingConvention = CallingConvention.Cdecl)]
    private static extern int NativeIsUsable(IntPtr session);
    [DllImport(NativeLibrary, EntryPoint = "qx_engine_identity", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr NativeEngineIdentity();
}
