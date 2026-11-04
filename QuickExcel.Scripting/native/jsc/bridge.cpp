// QuickExcel's pinned JSC embedding. Compiled into JavaScriptCore so private C++
// interfaces never cross the C ABI. This file is not an upstream source copy.
#include "config.h"
#include "APICast.h"
#include "Completion.h"
#include "TopExceptionScope.h"
#include "GlobalObjectMethodTable.h"
#include "Error.h"
#include "ErrorInstance.h"
#include "Exception.h"
#include "InitializeThreading.h"
#include "Interpreter.h"
#include "JSCInlines.h"
#include "JSGlobalObjectInlines.h"
#include "JSArray.h"
#include "JSModuleLoader.h"
#include "JSModuleRecord.h"
#include "JSPromise.h"
#include "JSSourceCode.h"
#include "MicrotaskQueue.h"
#include "MicrotaskQueueInlines.h"
#include "Options.h"
#include "SourceCode.h"
#include "StackFrame.h"
#include "Strong.h"
#include <JavaScriptCore/JavaScript.h>
#include <atomic>
#include <algorithm>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#if OS(WINDOWS)
#define QX_EXPORT extern "C" __declspec(dllexport)
#else
#define QX_EXPORT extern "C" __attribute__((visibility("default")))
#endif

using namespace JSC;
using namespace WTF;
using Clock = std::chrono::steady_clock;
constexpr size_t messageLimit = 16 * 1024 * 1024;
constexpr size_t diagnosticLimit = 16384;

struct Module {
    String id;
    String source;
    Strong<JSModuleRecord> record;
    unsigned declaredRoles { 3 };
    unsigned roles { 0 };
};

struct qx_session {
    JSContextGroupRef group { nullptr };
    JSGlobalContextRef context { nullptr };
    JSObjectRef bridge { nullptr };
    GlobalObjectMethodTable methods { };
    std::unique_ptr<VM::DrainMicrotaskDelayScope> delay;
    std::vector<Module> modules;
    bool modulesAttempted { false };
    bool compiler { false };
    bool poisoned { false };
    bool asyncRejected { false };
    const char* (*hostCall)(const char*) { nullptr };
    std::atomic<bool> interrupted { false };
    std::mutex timerMutex;
    std::condition_variable timerChanged;
    bool stopped { false };
    bool armed { false };
    uint64_t generation { 0 };
    Clock::time_point deadline;
    std::thread timer;

    VM& vm() { return *toJS(group); }
    JSGlobalObject* global() { return toJS(context); }
    void startTimer()
    {
        timer = std::thread([this] {
            std::unique_lock lock(timerMutex);
            while (!stopped) {
                timerChanged.wait(lock, [this] { return stopped || armed; });
                if (stopped) break;
                const auto current = generation;
                if (timerChanged.wait_until(lock, deadline, [this, current] { return stopped || !armed || generation != current; }))
                    continue;
                // This is the engine's concurrent-safe uncatchable termination
                // trap, driven by wall time rather than per-thread CPU time.
                interrupted.store(true);
                vm().notifyNeedTermination();
                armed = false;
            }
        });
    }
};

static thread_local qx_session* activeSession = nullptr;

static Module* findModule(qx_session&, const String&);

static JSValueRef registrationRole(JSContextRef context, JSObjectRef, JSObjectRef, size_t, const JSValueRef[], JSValueRef*)
{
    if (!activeSession) return JSValueMakeNumber(context, 0);
    Vector<StackFrame> frames;
    activeSession->vm().interpreter.getStackTrace(activeSession->global(), frames, 0, 64);
    for (const auto& frame : frames) {
        auto* module = findModule(*activeSession, frame.sourceURL(activeSession->vm(), AllowURLOverride::No));
        if (module) return JSValueMakeNumber(context, module->roles);
    }
    return JSValueMakeNumber(context, 0);
}

static char* copyText(std::string_view text)
{
    auto* result = static_cast<char*>(malloc(text.size() + 1));
    if (result) { memcpy(result, text.data(), text.size()); result[text.size()] = 0; }
    return result;
}

static void setError(char** error, std::string_view message)
{
    if (!error) return;
    size_t length = std::min(message.size(), diagnosticLimit);
    while (length && length < message.size() && (static_cast<unsigned char>(message[length]) & 0xc0) == 0x80) --length;
    *error = copyText(message.substr(0, length));
}

// A bounded synchronous message to the worker's managed IPC client. No OS or
// CLR object is exposed to JavaScript. The managed reply lives until this call returns.
static JSValueRef hostModelCall(JSContextRef context, JSObjectRef, JSObjectRef, size_t count, const JSValueRef values[], JSValueRef* exception)
{
    auto fail = [&](const char* message) -> JSValueRef {
        auto text = JSStringCreateWithUTF8CString(message);
        auto value = JSValueMakeString(context, text); JSStringRelease(text);
        *exception = JSObjectMakeError(context, 1, &value, nullptr);
        return JSValueMakeUndefined(context);
    };
    if (!activeSession || !activeSession->hostCall || count != 1 || !JSValueIsString(context, values[0]))
        return fail("Workbook model is unavailable outside an action.");
    auto text = JSValueToStringCopy(context, values[0], exception);
    if (!text) return JSValueMakeUndefined(context);
    if (JSStringGetLength(text) > messageLimit) { JSStringRelease(text); return fail("Model request exceeds its budget."); }
    std::string request(JSStringGetMaximumUTF8CStringSize(text), '\0');
    auto size = JSStringGetUTF8CString(text, request.data(), request.size()); JSStringRelease(text);
    if (!size || size - 1 > messageLimit) return fail("Model request exceeds its budget.");
    auto* reply = activeSession->hostCall(request.c_str());
    if (!reply || strnlen(reply, messageLimit + 1) > messageLimit) return fail("Invalid model response.");
    auto json = JSStringCreateWithUTF8CString(reply);
    auto value = JSValueMakeFromJSONString(context, json); JSStringRelease(json);
    return value ? value : fail("Invalid model response JSON.");
}

class Operation {
public:
    explicit Operation(qx_session& session) : s(session)
    {
        activeSession = &s;
        s.asyncRejected = false;
        s.interrupted.store(false);
        std::lock_guard lock(s.timerMutex);
        s.deadline = Clock::now() + std::chrono::seconds(s.compiler ? 5 : 1);
        ++s.generation;
        s.armed = true;
        s.timerChanged.notify_one();
    }
    ~Operation()
    {
        std::lock_guard lock(s.timerMutex);
        s.armed = false;
        ++s.generation;
        s.timerChanged.notify_one();
        activeSession = nullptr;
    }
    bool complete(char** error, bool final = true)
    {
        // Serialize completion against the watchdog: no late trap may reach the
        // next call, and a result finishing after its deadline is never accepted.
        std::lock_guard lock(s.timerMutex);
        if (final) {
            s.armed = false;
            ++s.generation;
            s.timerChanged.notify_one();
        }
        if (Clock::now() >= s.deadline) s.interrupted.store(true);
        if (s.interrupted.load()) {
            s.poisoned = true;
            setError(error, s.compiler ? "TypeScript compilation exceeded the 5 second time limit." : "JavaScript execution exceeded the 1 second time limit.");
            return false;
        }
        // Legacy v1 never pumps queued jobs; preserve that inert behavior.
        // Package modules require the stronger synchronous contract.
        if ((s.compiler || s.modulesAttempted) && (s.asyncRejected || !s.global()->microtaskQueue().isEmpty())) {
            s.poisoned = true;
            s.global()->microtaskQueue().clear();
            setError(error, "Package modules require synchronous execution; dynamic import and pending jobs are unsupported.");
            return false;
        }
        return true;
    }
private:
    qx_session& s;
};

static JSPromise* rejectDynamicImport(JSGlobalObject* global, JSModuleLoader*, JSString*, RefPtr<ScriptFetchParameters>, const SourceOrigin&, bool)
{
    if (activeSession) activeSession->asyncRejected = true;
    auto scope = DECLARE_THROW_SCOPE(global->vm());
    throwTypeError(global, scope, "Dynamic import is not supported by QuickExcel"_s);
    return nullptr;
}

static JSGlobalObject* rejectRealm(JSGlobalObject* global)
{
    auto scope = DECLARE_THROW_SCOPE(global->vm());
    throwTypeError(global, scope, "Additional execution realms are not supported"_s);
    return nullptr;
}

static bool validId(const String& id)
{
    if (id.isEmpty() || id.length() > 128 || id == "."_s || id == ".."_s) return false;
    for (unsigned i = 0; i < id.length(); ++i) {
        auto c = id[i];
        if (!((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '_' || c == '-' || c == '.')) return false;
    }
    return true;
}

static std::string utf8(const String& value)
{
    auto text = value.utf8();
    return { text.data(), text.length() };
}

static void exceptionError(qx_session& s, JSValue value, char** error, const String& module = { }, bool parsing = false)
{
    if (s.interrupted.load()) {
        s.poisoned = true;
        setError(error, s.compiler ? "TypeScript compilation exceeded the 5 second time limit." : "JavaScript execution exceeded the 1 second time limit.");
        return;
    }
    std::string message = "JavaScript execution failed";
    if (auto* instance = dynamicDowncast<ErrorInstance>(value)) {
        // VMInquiry reads only an own data string. Never invoke message/name/
        // stack/toString accessors or a Proxy trap while reporting an exception.
        auto detail = instance->tryGetMessageForDebugging();
        message = utf8(String(errorTypeName(instance->errorType()))) + ": " + utf8(detail.left(diagnosticLimit));
        if (parsing && instance->isParseError() && !module.isEmpty()) {
            auto line = instance->getDirect(s.vm(), Identifier::fromString(s.vm(), "line"_s));
            if (line.isInt32()) message += "\n at " + utf8(module) + ":" + std::to_string(line.asInt32());
        }
        if (instance->isStackOverflowError()) { message = "JavaScript call stack size limit exceeded."; s.poisoned = true; }
        if (instance->isOutOfMemoryError()) { message = "JavaScript out of memory."; s.poisoned = true; }
        if (auto* frames = instance->stackTrace()) {
            unsigned count = 0;
            for (const auto& frame : *frames) {
                if (!frame.hasLineAndColumnInfo() || count++ >= 12) continue;
                auto id = frame.sourceURL(s.vm(), AllowURLOverride::No);
                if (!validId(id) && id != "quickexcel-host.js"_s) continue;
                auto position = frame.computeLineAndColumn();
                message += "\n at " + utf8(id) + ":" + std::to_string(position.line) + ":" + std::to_string(position.column);
                if (message.size() >= diagnosticLimit) break;
            }
        }
    } else if (value.isString()) {
        message = utf8(asString(value)->value(s.global()).data.left(diagnosticLimit));
    }
    if (!module.isEmpty()) message = utf8(module) + ": " + message;
    setError(error, message);
}

static Module* findModule(qx_session& s, const String& id)
{
    for (auto& module : s.modules) if (module.id == id) return &module;
    return nullptr;
}

static JSModuleRecord* prepareModule(qx_session& s, Module& module, char** error)
{
    if (module.record) return module.record.get();
    VM& vm = s.vm();
    auto scope = DECLARE_TOP_EXCEPTION_SCOPE(vm);
    auto source = makeSource(module.source, SourceOrigin { }, SourceTaintedOrigin::Untainted,
        module.id, TextPosition(), SourceProviderSourceType::Module);
    // makeModule only parses/analyzes and settles its internal promise. It does
    // not register continuations or execute any workbook code.
    auto* parsed = JSModuleLoader::makeModule(s.global(), Identifier::fromString(vm, module.id), JSSourceCode::create(vm, WTF::move(source)));
    if (scope.exception()) {
        exceptionError(s, scope.exception()->value(), error, module.id); scope.clearException(); return nullptr;
    }
    if (!parsed || parsed->status() != JSPromise::Status::Fulfilled) {
        if (parsed && parsed->status() == JSPromise::Status::Rejected) exceptionError(s, parsed->result(), error, module.id, true);
        else setError(error, "Module parsing did not complete synchronously.");
        return nullptr;
    }
    auto* record = dynamicDowncast<JSModuleRecord>(parsed->result());
    if (!record || record->hasTLA()) { setError(error, "Top-level await is not supported; modules must initialize synchronously."); return nullptr; }
    module.record.set(vm, record); // Root before recursive graph construction (including cycles).
    for (const auto& request : record->requestedModules()) {
        if (request.m_attributes || request.m_phase != AbstractModuleRecord::ModulePhase::Evaluation) {
            setError(error, "Import attributes and deferred imports are not supported."); return nullptr;
        }
        Module* dependency = findModule(s, request.m_specifier.string());
        if (!dependency) { setError(error, "Import must name a declared package module ID."); return nullptr; }
        auto* imported = prepareModule(s, *dependency, error);
        if (!imported) return nullptr;
        record->loadedModules().set(ModuleMapKey { request.m_specifier.impl(), request.type() },
            AbstractModuleRecord::LoadedModuleRequest { vm, request, imported, record });
    }
    record->setStatus(CyclicModuleRecord::Status::Unlinked);
    return record;
}

QX_EXPORT void qx_destroy(qx_session* s)
{
    if (!s) return;
    { std::lock_guard lock(s->timerMutex); s->stopped = true; s->armed = false; s->timerChanged.notify_one(); }
    if (s->timer.joinable()) s->timer.join();
    if (s->group) {
        JSLockHolder lock(s->vm());
        if (s->context) s->global()->microtaskQueue().clear();
        s->vm().defaultMicrotaskQueue().clear();
        s->modules.clear();
        if (s->bridge) JSValueUnprotect(s->context, s->bridge);
        // Both queues are empty before releasing the lifetime-long delay scope.
        s->delay.reset();
        if (s->context) JSGlobalContextRelease(s->context);
    }
    if (s->group) JSContextGroupRelease(s->group);
    delete s;
}

static qx_session* createSession(const char* bootstrap, char** error, bool compiler)
{
    if (error) *error = nullptr;
    if (!bootstrap || strlen(bootstrap) > messageLimit) { setError(error, "Invalid bootstrap."); return nullptr; }
    static std::once_flag initialized;
    static bool configured = false;
    static bool compilerProcess = false;
    std::call_once(initialized, [compiler] {
        JSC::initialize();
        compilerProcess = compiler;
        configured = Options::setOptions(compiler
            ? "maxPerThreadStackUsage=5242880 jitMemoryReservationSize=16777216 gcMaxHeapSize=33554432"
            : "maxPerThreadStackUsage=262144 jitMemoryReservationSize=16777216");
    });
    if (!configured || compilerProcess != compiler) { setError(error, "JSC role or policy configuration failed."); return nullptr; }
    auto* s = new qx_session;
    s->compiler = compiler;
    s->group = JSContextGroupCreate();
    {
        JSLockHolder lock(s->vm());
        s->delay = std::make_unique<VM::DrainMicrotaskDelayScope>(s->vm());
        auto* base = JSGlobalObject::create(s->vm(), JSGlobalObject::createStructure(s->vm(), jsNull()));
        s->methods = *base->globalObjectMethodTable();
        s->methods.moduleLoaderImportModule = rejectDynamicImport;
        s->methods.deriveShadowRealmGlobalObject = rejectRealm;
        auto* global = JSGlobalObject::createWithCustomMethodTable(s->vm(), JSGlobalObject::createStructure(s->vm(), jsNull()), &s->methods);
        s->context = JSGlobalContextRetain(toGlobalRef(global));
        if (!compiler) {
            auto name = JSStringCreateWithUTF8CString("__qxModelCall");
            auto function = JSObjectMakeFunctionWithCallback(s->context, name, hostModelCall);
            JSObjectSetProperty(s->context, JSContextGetGlobalObject(s->context), name, function, kJSPropertyAttributeDontEnum, nullptr);
            JSStringRelease(name);
            name = JSStringCreateWithUTF8CString("__qxRegistrationRole");
            function = JSObjectMakeFunctionWithCallback(s->context, name, registrationRole);
            JSObjectSetProperty(s->context, JSContextGetGlobalObject(s->context), name, function, kJSPropertyAttributeDontEnum, nullptr);
            JSStringRelease(name);
        }
        // notifyNeedTermination assumes this preallocated uncatchable exception
        // exists (the stock Watchdog creates it when initialized).
        s->vm().ensureTerminationException();
        s->vm().forbidExecutionOnTermination();
    }
    s->startTimer();
    bool valid = false;
    {
        Operation operation(*s);
        JSLockHolder lock(s->vm());
        auto script = JSStringCreateWithUTF8CString(bootstrap);
        auto origin = JSStringCreateWithUTF8CString("quickexcel-host.js");
        JSValueRef exception = nullptr;
        auto result = JSEvaluateScript(s->context, script, nullptr, origin, 1, &exception);
        JSStringRelease(script); JSStringRelease(origin);
        if (!exception && result && JSValueIsObject(s->context, result)) {
            s->bridge = JSValueToObject(s->context, result, nullptr);
            JSValueProtect(s->context, s->bridge);
            valid = operation.complete(error);
        } else exceptionError(*s, exception ? toJS(s->global(), exception) : jsUndefined(), error);
    }
    if (!valid) { qx_destroy(s); return nullptr; }
    return s;
}

QX_EXPORT qx_session* qx_create(const char* bootstrap, char** error) { return createSession(bootstrap, error, false); }
QX_EXPORT qx_session* qx_create_compiler(const char* bootstrap, char** error) { return createSession(bootstrap, error, true); }
QX_EXPORT void qx_set_host_call(qx_session* s, const char* (*callback)(const char*)) { if (s && !s->compiler) s->hostCall = callback; }

QX_EXPORT int qx_load_modules(qx_session* s, const char* json, char** error)
{
    if (error) *error = nullptr;
    if (!s || s->compiler || s->poisoned || s->modulesAttempted || !json || strlen(json) > messageLimit) {
        setError(error, "Module session is closed, invalid or already loaded."); return 0;
    }
    s->modulesAttempted = true;
    Operation operation(*s);
    JSLockHolder lock(s->vm());
    auto scope = DECLARE_TOP_EXCEPTION_SCOPE(s->vm());
    auto input = JSStringCreateWithUTF8CString(json);
    auto parsed = JSValueMakeFromJSONString(s->context, input);
    JSStringRelease(input);
    auto fail = [&] { s->poisoned = true; if (scope.exception()) scope.clearException(); return 0; };
    if (!parsed || !JSValueIsObject(s->context, parsed)) { setError(error, "Invalid module JSON."); return fail(); }
    auto object = asObject(toJS(s->global(), parsed));
    auto modules = object->get(s->global(), Identifier::fromString(s->vm(), "modules"_s));
    auto entries = object->get(s->global(), Identifier::fromString(s->vm(), "entries"_s));
    auto* array = dynamicDowncast<JSArray>(modules);
    auto* entryArray = dynamicDowncast<JSArray>(entries);
    if (scope.exception() || !array || !entryArray || !array->length() || array->length() > 64 || !entryArray->length() || entryArray->length() > 64) {
        setError(error, "Invalid module mapping."); return fail();
    }
    // The parsed package stays rooted across parser/GC calls.
    Strong<Unknown> package(s->vm(), toJS(s->global(), parsed));
    s->modules.reserve(array->length());
    size_t bytes = 0;
    for (unsigned i = 0; i < array->length(); ++i) {
        auto item = array->getIndex(s->global(), i);
        if (!item.isObject()) { setError(error, "Invalid module."); return fail(); }
        auto id = asObject(item)->get(s->global(), Identifier::fromString(s->vm(), "id"_s));
        auto source = asObject(item)->get(s->global(), Identifier::fromString(s->vm(), "source"_s));
        auto role = asObject(item)->get(s->global(), Identifier::fromString(s->vm(), "roles"_s));
        if (!role.isUndefined() && (!role.isInt32() || role.asInt32() < 0 || role.asInt32() > 3)) { setError(error, "Invalid module role."); return fail(); }
        if (!id.isString() || !source.isString()) { setError(error, "Invalid module source."); return fail(); }
        auto idText = asString(id)->value(s->global()).data;
        auto text = asString(source)->value(s->global()).data;
        if (scope.exception() || !validId(idText) || findModule(*s, idText) || (bytes += text.utf8().length()) > 1024 * 1024) {
            setError(error, "Invalid module mapping or source budget."); return fail();
        }
        s->modules.push_back({ idText, text, { }, role.isUndefined() ? 3u : static_cast<unsigned>(role.asInt32()), 0 });
    }
    std::vector<Module*> roots;
    for (unsigned i = 0; i < entryArray->length(); ++i) {
        auto id = entryArray->getIndex(s->global(), i);
        auto* root = id.isString() ? findModule(*s, asString(id)->value(s->global()).data) : nullptr;
        if (!root || std::find(roots.begin(), roots.end(), root) != roots.end()) { setError(error, "Invalid entry order."); return fail(); }
        roots.push_back(root);
        if (!prepareModule(*s, *root, error)) return fail();
    }
    // Native record operations preserve ES linking, cycles, TDZ and live bindings.
    std::function<void(Module&, unsigned)> bind = [&](Module& module, unsigned roles) {
        if ((module.roles & roles) == roles) return;
        module.roles |= roles;
        for (const auto& request : module.record->requestedModules())
            if (auto* dependency = findModule(*s, request.m_specifier.string())) bind(*dependency, roles);
    };
    for (auto* root : roots) bind(*root, root->declaredRoles);
    // No loader promise jobs or user jobs are advanced, even during cleanup.
    for (auto* root : roots) {
        root->record->link(s->global(), nullptr);
        if (scope.exception()) { exceptionError(*s, scope.exception()->value(), error, root->id); return fail(); }
    }
    for (auto* root : roots) {
        root->record->evaluateSync(s->global());
        if (scope.exception()) { exceptionError(*s, scope.exception()->value(), error, root->id); return fail(); }
        if (!operation.complete(error, false)) return fail();
        // Do not reset the deadline: all entries share one loading operation.
    }
    return operation.complete(error) ? 1 : fail();
}

QX_EXPORT char* qx_call(qx_session* s, const char* method, const char* json, char** error)
{
    if (error) *error = nullptr;
    if (!s || s->poisoned || !method || !json || strlen(json) > messageLimit) { setError(error, "JavaScript session is closed or invalidated."); return nullptr; }
    Operation operation(*s);
    JSLockHolder lock(s->vm());
    JSValueRef exception = nullptr;
    auto input = JSStringCreateWithUTF8CString(json);
    auto args = JSValueMakeFromJSONString(s->context, input);
    JSStringRelease(input);
    if (!args || !JSValueIsArray(s->context, args)) { setError(error, "Invalid call arguments."); return nullptr; }
    JSValueProtect(s->context, args);
    auto name = JSStringCreateWithUTF8CString(method);
    auto fnValue = JSObjectGetProperty(s->context, s->bridge, name, &exception);
    JSStringRelease(name);
    auto* array = JSValueToObject(s->context, args, &exception);
    JSValueRef argv[2] = { JSObjectGetPropertyAtIndex(s->context, array, 0, &exception), JSObjectGetPropertyAtIndex(s->context, array, 1, &exception) };
    auto fn = exception || !JSValueIsObject(s->context, fnValue) ? nullptr : JSValueToObject(s->context, fnValue, &exception);
    auto value = fn && JSObjectIsFunction(s->context, fn) ? JSObjectCallAsFunction(s->context, fn, s->bridge, 2, argv, &exception) : nullptr;
    JSValueUnprotect(s->context, args);
    if (exception || !value) {
        exceptionError(*s, exception ? toJS(s->global(), exception) : jsUndefined(), error);
        char* completionError = nullptr;
        if (!operation.complete(&completionError)) {
            if (error) { free(*error); *error = completionError; } else free(completionError);
        }
        return nullptr;
    }
    // Serialization may invoke user toJSON/getters; it remains inside the same
    // deadline and pending-job check, never after accepting an operation result.
    auto serialized = JSValueCreateJSONString(s->context, JSValueIsUndefined(s->context, value) ? JSValueMakeNull(s->context) : value, 0, &exception);
    char* result = nullptr;
    if (serialized && !exception) {
        auto length = JSStringGetLength(serialized);
        if (length <= messageLimit) {
            std::string text(JSStringGetMaximumUTF8CStringSize(serialized), '\0');
            auto count = JSStringGetUTF8CString(serialized, text.data(), text.size());
            if (count && count - 1 <= messageLimit) result = copyText(std::string_view(text.data(), count - 1));
        }
        if (!result) setError(error, "JavaScript result exceeds the 16 MiB result limit.");
    } else exceptionError(*s, exception ? toJS(s->global(), exception) : jsUndefined(), error);
    if (serialized) JSStringRelease(serialized);
    char* completionError = nullptr;
    if (!operation.complete(&completionError)) {
        free(result); result = nullptr;
        if (error) { free(*error); *error = completionError; } else free(completionError);
    }
    return result;
}

QX_EXPORT void qx_free_string(char* text) { free(text); }
QX_EXPORT int qx_was_interrupted(qx_session* s) { return s && s->interrupted.load(); }
QX_EXPORT int qx_is_usable(qx_session* s) { return s && !s->poisoned; }
QX_EXPORT const char* qx_engine_identity() { return "JavaScriptCore/2.54.2/b95c91b1/qx-sync-2"; }
