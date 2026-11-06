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

    // The bridge object is held by native code, not exposed on globalThis. Module
    // code sees only the host extension object and a bounded console.
    private static readonly string Bootstrap = """
        (() => {
            const nativeModelCall = globalThis.__qxModelCall;
            delete globalThis.__qxModelCall;
            const registrationRole = globalThis.__qxRegistrationRole;
            delete globalThis.__qxRegistrationRole;
            /*VIRTUAL_FILES*/
            const registry = Object.create(null);
            const descriptors = Object.create(null);
            const mains = Object.create(null);
            let actionDescriptions = Object.create(null);
            const eventHandlers = Object.create(null);
            const ownKeys = Object.keys;
            const freeze = Object.freeze;
            const makeFunction = Function;
            const upper = Function.prototype.call.bind(String.prototype.toUpperCase);
            const validName = Function.prototype.call.bind(RegExp.prototype.test, /^[A-Z]+$/);
            const validV2Name = Function.prototype.call.bind(RegExp.prototype.test, /^JS\.[A-Z_][A-Z0-9_]*(\.[A-Z_][A-Z0-9_]*)*$/);
            const isArray = Array.isArray;
            const isFiniteNumber = Number.isFinite;
            const TypeErrorCtor = TypeError;
            const capabilities = JSON.parse(/*CAPABILITIES*/);
            let loaded = false;
            let actionMode = false;
            let sharedMode = false;
            let actionInvoked = false;
            let activeModel = false;
            let hostWriteFailed = false;
            let hostWrote = false;
            let automatic = false;
            let initialized = true;
            const initializers = [];
            class ScriptError extends Error {
                constructor(code, message, member) {
                    super(message);
                    if (code === 'InvalidArgument' && hostWrote || code === 'BudgetExceeded' || code === 'Locked') hostWriteFailed = true;
                    Object.defineProperty(this, 'name', { value: 'ScriptError', configurable: true });
                    Object.defineProperty(this, 'code', { value: code, enumerable: true });
                    if (member !== undefined) Object.defineProperty(this, 'member', { value: member, enumerable: true });
                }
            }
            const deepFreeze = value => {
                if (value && typeof value === 'object') { for (const key of ownKeys(value)) deepFreeze(value[key]); freeze(value); }
                return value;
            };
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
            const console = freeze({ log: (...v) => log('', v), info: (...v) => log('[info] ', v),
                warn: (...v) => log('[warn] ', v), error: (...v) => log('[error] ', v), debug: (...v) => log('[debug] ', v) });
            function log(prefix, values) {
                if (logs.length < 100) logs.push((prefix + values.slice(0, 30).map(value => display(value)).join(' ')).slice(0, 4000));
            }
            /*EVENT_API*/
            /*WORKBOOK_MODEL*/
            /*STORAGE*/
            /*MODULE_CONVENTIONS*/
            return freeze({
                load(source) {
                    if (loaded) throw new ScriptError('InvalidArgument', '脚本已经加载。');
                    makeFunction('host', 'console', '"use strict";\n' + source)(host, console);
                    loaded = true;
                    return ownKeys(registry);
                },
                prepareModules(useActions, options) {
                    if (loaded) throw new ScriptError('InvalidArgument', '脚本已经加载。');
                    actionMode = useActions === true;
                    sharedMode = options.share === true;
                    automatic = options.automatic === true;
                    if (automatic) host = freeze({ localStorage, sessionStorage, values: host.values, get files() { return fileContext(); } });
                    Object.defineProperty(globalThis, 'host', { value: host });
                    Object.defineProperty(globalThis, 'console', { value: console });
                    Object.defineProperty(globalThis, '__qxRegisterMain', {
                        value: (id, entry) => { if (typeof id === 'string' && typeof entry === 'function') mains[id] = entry; },
                        configurable: true
                    });
                    if (automatic) Object.defineProperty(globalThis, '__qxBindExports', { value: bindExports, configurable: true });
                },
                finishModules(entries, descriptions) {
                    delete globalThis.__qxRegisterMain;
                    delete globalThis.__qxBindExports;
                    loaded = true;
                    const roots = isArray(entries) ? entries.filter(id => typeof id === 'string') : [];
                    actionEntries.length = 0;
                    for (const id of roots) actionEntries.push(id);
                    actionDescriptions = Object.create(null);
                    for (const id of ownKeys(descriptions || {}))
                        if (typeof descriptions[id] === 'string') actionDescriptions[id] = descriptions[id].slice(0, 200);
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
                actionDescriptors() {
                    // Event bindings stay visible to Core as @event/ pseudo entries so
                    // notifications are gated by what the module actually registered;
                    // the action picker filters them out.
                    const result = actionEntries.filter(name => typeof mains[name] === 'function')
                        .map(name => ({ name, description: actionDescriptions[name] || '' }));
                    if (initializers.length) result.push({ name: '@initialize', description: '' });
                    for (const name of ownKeys(eventHandlers)) result.push({ name: '@event/' + name, description: '' });
                    return result;
                },
                beginRound() { logs.length = 0; actionInvoked = false; activeModel = false; hostWriteFailed = false; hostWrote = false; },
                runAction,
                logs() { return logs.slice(); }
            });
        })()
        """.Replace("/*VIRTUAL_FILES*/", VirtualFileBootstrap, StringComparison.Ordinal)
            .Replace("/*EVENT_API*/", EventApiBootstrap, StringComparison.Ordinal)
            .Replace("/*WORKBOOK_MODEL*/", WorkbookModelBootstrap, StringComparison.Ordinal)
            .Replace("/*STORAGE*/", StorageBootstrap, StringComparison.Ordinal)
            .Replace("/*MODULE_CONVENTIONS*/", ModuleConventionsBootstrap, StringComparison.Ordinal)
            .Replace("/*CAPABILITIES*/", JsonSerializer.Serialize(ScriptCapabilities.Current.Json), StringComparison.Ordinal);

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

    public JscSession(string source) : this(source, null, null, false, null, null, false, null, null) { }

    public JscSession(IReadOnlyDictionary<string, string> modules, IReadOnlyList<string> entries, bool actions = false, IReadOnlyDictionary<string, int>? roles = null,
        IReadOnlyDictionary<string, string>? descriptions = null, bool automatic = false,
        IReadOnlyList<string>? eventEntries = null, IReadOnlyList<string>? initializationEntries = null)
        : this(null, modules, entries, actions, roles, descriptions, automatic, eventEntries, initializationEntries) { }

    private JscSession(string? source, IReadOnlyDictionary<string, string>? modules, IReadOnlyList<string>? entries, bool actions, IReadOnlyDictionary<string, int>? roles,
        IReadOnlyDictionary<string, string>? descriptions, bool automatic, IReadOnlyList<string>? eventEntries, IReadOnlyList<string>? initializationEntries)
    {
        _hostCallback = HostCall;
        if (modules is null) ArgumentNullException.ThrowIfNull(source);
        else
        {
            if (modules.Count is < 1 or > 64 || entries is null || entries.Count is < 1 or > 64 ||
                entries.Distinct(StringComparer.Ordinal).Count() != entries.Count || entries.Any(id => id is null || !modules.ContainsKey(id)) ||
                modules.Any(m => string.IsNullOrEmpty(m.Key) || m.Key.Length > 128 || m.Key is "." or ".." ||
                    m.Key.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')) || m.Value is null) ||
                descriptions?.Any(d => !modules.ContainsKey(d.Key) || d.Value is null || d.Value.Length > 200 || d.Value.Any(char.IsControl)) == true ||
                modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value)) > 1024 * 1024)
                throw new ArgumentException("Invalid package module mapping or source budget.");
        }
        if (!Environment.Is64BitProcess || (!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux() &&
            !(OperatingSystem.IsMacOS() && System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture == System.Runtime.InteropServices.Architecture.Arm64)))
            throw new PlatformNotSupportedException("QuickExcel scripting requires Windows x64, Linux x64 or qualified macOS arm64 isolation.");
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
                Call("prepareModules", [actions, new { share = roles is not null, automatic }]);
                // Each entry module registers its own main(workbook, host) binding
                // during evaluation; the native loader does not expose exports.
                var rooted = new Dictionary<string, string>(modules, StringComparer.Ordinal);
                foreach (var entry in automatic || actions ? entries! : [])
                {
                    if (automatic)
                    {
                        var local = "__qxExports";
                        while (rooted[entry].Contains(local, StringComparison.Ordinal)) local += "_";
                        var kind = !actions ? "udf" : initializationEntries?.Contains(entry) == true ? "initialization" : eventEntries?.Contains(entry) == true ? "events" : "actions";
                        rooted[entry] += "\n;import * as " + local + " from " + JsonSerializer.Serialize(entry) + ";\nglobalThis.__qxBindExports(" +
                            JsonSerializer.Serialize(entry) + "," + JsonSerializer.Serialize(kind) + "," + local + "," + JsonSerializer.Serialize(FunctionPrefix(entry)) + ");";
                    }
                    else rooted[entry] += "\n;globalThis.__qxRegisterMain && globalThis.__qxRegisterMain(" +
                        JsonSerializer.Serialize(entry) + ", typeof main === 'function' ? main : undefined);";
                }
                var actionEntries = actions
                    ? entries!.Where(id => (roles is null || (roles.GetValueOrDefault(id) & 2) != 0) &&
                        initializationEntries?.Contains(id) != true && eventEntries?.Contains(id) != true).ToArray()
                    : [];
                var json = JsonSerializer.Serialize(new { modules = rooted.Select(m => new { id = m.Key, source = m.Value, roles = roles is null ? 3 : roles.GetValueOrDefault(m.Key) }), entries });
                if (Encoding.UTF8.GetByteCount(json) > MaxMessageBytes) throw new ArgumentException("Module payload exceeds 16 MiB UTF-8.");
                if (NativeLoadModules(_session, json, out var moduleError) == 0) throw ReadException(moduleError);
                result = Call("finishModules", [actionEntries, descriptions is null ? new Dictionary<string, string>() : new Dictionary<string, string>(descriptions, StringComparer.Ordinal)]);
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
        catch (Exception error) { response = JsonSerializer.Serialize(new { ok = false, error = error.Message.Length <= 4000 ? error.Message : "模型请求失败。", code = "TransactionAborted" }); }
        if (Encoding.UTF8.GetByteCount(response) > MaxMessageBytes) response = "{\"ok\":false,\"error\":\"模型响应超过 16 MiB。\",\"code\":\"BudgetExceeded\"}";
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
