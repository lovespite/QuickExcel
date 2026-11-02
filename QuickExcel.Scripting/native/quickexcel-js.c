#ifdef _WIN32
#include <windows.h>
#else
#include <time.h>
#endif
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#include "quickjs.h"

#ifdef _WIN32
#define QX_EXPORT __declspec(dllexport)
#else
#define QX_EXPORT __attribute__((visibility("default")))
#endif

static uint64_t qx_monotonic_milliseconds(void) {
#ifdef _WIN32
    return GetTickCount64();
#else
    struct timespec now;
    if (clock_gettime(CLOCK_MONOTONIC, &now) != 0) abort();
    return (uint64_t)now.tv_sec * 1000 + (uint64_t)now.tv_nsec / 1000000;
#endif
}
#define QX_RESULT_LIMIT (16 * 1024 * 1024)

typedef struct {
    const char *id;
    const char *source;
    size_t source_length;
} qx_module;

typedef struct {
    JSRuntime *runtime;
    JSContext *context;
    JSValue bridge;
    JSValue error_stack_getter;
    uint64_t deadline;
    int interrupted;
    qx_module modules[64];
    uint32_t module_count;
    int modules_loading;
} qx_session;

static char *copy_text(const char *text, size_t length) {
    char *copy = malloc(length + 1);
    if (copy) { memcpy(copy, text, length); copy[length] = 0; }
    return copy;
}

static void set_error(char **error, const char *message) {
    if (error) *error = copy_text(message, strlen(message));
}

static int interrupt(JSRuntime *runtime, void *opaque) {
    (void)runtime;
    qx_session *session = opaque;
    if (qx_monotonic_milliseconds() >= session->deadline) {
        session->interrupted = 1;
        return 1;
    }
    return 0;
}

static void start_operation(qx_session *session) {
    JS_UpdateStackTop(session->runtime);
    session->interrupted = 0;
    session->deadline = qx_monotonic_milliseconds() + 1000;
}

static size_t diagnostic_length(const char *text, size_t length, size_t budget) {
    if (length <= budget) return length;
    length = budget;
    while (length && (((unsigned char)text[length] & 0xc0) == 0x80)) length--;
    return length;
}

static void exception_error(qx_session *session, char **error) {
    JSValue exception = JS_GetException(session->context);
    if (session->interrupted) {
        set_error(error, "JavaScript execution exceeded the 1 second time limit.");
    } else {
        size_t length = 0;
        const char *text = JS_ToCStringLen(session->context, &length, exception);
        if (text) {
            length = diagnostic_length(text, length, 16384);
            if (error) {
                *error = malloc(16385);
                if (*error) {
                    memcpy(*error, text, length);
                    /* Call the pinned engine getter directly. User-defined stack
                       getters and replaced Error prototypes are never consulted. */
                    if (length < 16383 && JS_IsError(exception) && !session->interrupted) {
                        JSValue trace = JS_Call(session->context, session->error_stack_getter, exception, 0, NULL);
                        if (JS_IsString(trace)) {
                            size_t stack_length = 0;
                            const char *stack = JS_ToCStringLen(session->context, &stack_length, trace);
                            if (stack) {
                                stack_length = diagnostic_length(stack, stack_length, 16384 - length - 1);
                                if (stack_length) { (*error)[length++] = '\n'; memcpy(*error + length, stack, stack_length); length += stack_length; }
                                JS_FreeCString(session->context, stack);
                            }
                        } else if (JS_IsException(trace)) {
                            JSValue secondary = JS_GetException(session->context);
                            JS_FreeValue(session->context, secondary);
                        }
                        JS_FreeValue(session->context, trace);
                    }
                    (*error)[length] = 0;
                }
            }
            JS_FreeCString(session->context, text);
        } else {
            set_error(error, "JavaScript failed (execution or memory limit reached).");
            JSValue secondary = JS_GetException(session->context);
            JS_FreeValue(session->context, secondary);
        }
    }
    JS_FreeValue(session->context, exception);
}

QX_EXPORT void qx_destroy(qx_session *session) {
    if (!session) return;
    if (session->context) {
        for (uint32_t i = 0; i < session->module_count; i++) {
            if (session->modules[i].id) JS_FreeCString(session->context, session->modules[i].id);
            if (session->modules[i].source) JS_FreeCString(session->context, session->modules[i].source);
        }
        JS_FreeValue(session->context, session->bridge);
        JS_FreeValue(session->context, session->error_stack_getter);
        JS_FreeContext(session->context);
    }
    if (session->runtime) JS_FreeRuntime(session->runtime);
    free(session);
}

QX_EXPORT qx_session *qx_create(const char *bootstrap, char **error) {
    if (error) *error = NULL;
    qx_session *session = calloc(1, sizeof(qx_session));
    if (!session) { set_error(error, "Unable to allocate JavaScript session."); return NULL; }
    session->bridge = JS_UNDEFINED;
    session->error_stack_getter = JS_UNDEFINED;
    session->runtime = JS_NewRuntime();
    if (!session->runtime) { set_error(error, "Unable to create QuickJS runtime."); qx_destroy(session); return NULL; }
    JS_SetMemoryLimit(session->runtime, 64 * 1024 * 1024);
    JS_SetMaxStackSize(session->runtime, 256 * 1024);
    JS_SetInterruptHandler(session->runtime, interrupt, session);
    session->context = JS_NewContext(session->runtime);
    if (!session->context) { set_error(error, "Unable to create QuickJS context."); qx_destroy(session); return NULL; }
    start_operation(session);
    JSValue sample_error = JS_NewError(session->context);
    JSValue error_prototype = JS_GetPrototype(session->context, sample_error);
    JSAtom stack_atom = JS_NewAtom(session->context, "stack");
    JSPropertyDescriptor stack_descriptor;
    int stack_found = JS_GetOwnProperty(session->context, &stack_descriptor, error_prototype, stack_atom);
    JS_FreeAtom(session->context, stack_atom);
    JS_FreeValue(session->context, error_prototype);
    JS_FreeValue(session->context, sample_error);
    if (stack_found > 0) {
        session->error_stack_getter = stack_descriptor.getter;
        JS_FreeValue(session->context, stack_descriptor.value);
        JS_FreeValue(session->context, stack_descriptor.setter);
    }
    if (!JS_IsFunction(session->context, session->error_stack_getter)) {
        set_error(error, "Unable to capture engine error diagnostics."); qx_destroy(session); return NULL;
    }
    session->bridge = JS_Eval(session->context, bootstrap, strlen(bootstrap), "quickexcel-host.js", JS_EVAL_TYPE_GLOBAL);
    if (JS_IsException(session->bridge)) {
        exception_error(session, error);
        session->bridge = JS_UNDEFINED;
        qx_destroy(session);
        return NULL;
    }
    return session;
}

static int valid_module_id(const char *id, size_t length) {
    if (length == 0 || length > 128 || !strcmp(id, ".") || !strcmp(id, "..")) return 0;
    for (size_t i = 0; i < length; i++) {
        char c = id[i];
        if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '-' || c == '.')) return 0;
    }
    return 1;
}

static qx_module *find_module(qx_session *session, const char *id) {
    for (uint32_t i = 0; i < session->module_count; i++)
        if (!strcmp(session->modules[i].id, id)) return &session->modules[i];
    return NULL;
}

/* Exact logical IDs only: never use QuickJS's filename normalizer or libc loader. */
static char *normalize_module(JSContext *ctx, const char *base, const char *name, void *opaque) {
    (void)base;
    qx_session *session = opaque;
    if (!session->modules_loading || !find_module(session, name)) {
        JS_ThrowReferenceError(ctx, "Import must name a declared package module ID.");
        return NULL;
    }
    return js_strdup(ctx, name);
}

static int check_import_attributes(JSContext *ctx, void *opaque, JSValueConst attributes) {
    (void)opaque;
    if (!JS_IsUndefined(attributes)) {
        JS_ThrowTypeError(ctx, "Import attributes are not supported.");
        return -1;
    }
    return 0;
}

static JSModuleDef *load_module(JSContext *ctx, const char *name, void *opaque, JSValueConst attributes) {
    (void)attributes;
    qx_session *session = opaque;
    qx_module *module = find_module(session, name);
    if (!session->modules_loading || !module) {
        JS_ThrowReferenceError(ctx, "Package module is unavailable.");
        return NULL;
    }
    JSValue compiled = JS_Eval(ctx, module->source, module->source_length, module->id,
        JS_EVAL_TYPE_MODULE | JS_EVAL_FLAG_COMPILE_ONLY);
    if (JS_IsException(compiled)) return NULL;
    JSModuleDef *definition = JS_VALUE_GET_PTR(compiled);
    JS_FreeValue(ctx, compiled);
    return definition;
}

QX_EXPORT int qx_load_modules(qx_session *session, const char *json, char **error) {
    if (error) *error = NULL;
    if (!session || session->module_count) { set_error(error, "Module session is closed or already loaded."); return 0; }
    start_operation(session);
    JSContext *ctx = session->context;
    JSValue package = JS_ParseJSON(ctx, json, strlen(json), "package.json");
    if (JS_IsException(package)) { exception_error(session, error); return 0; }
    JSValue modules = JS_GetPropertyStr(ctx, package, "modules");
    JSValue entries = JS_GetPropertyStr(ctx, package, "entries");
    JSValue count_value = JS_GetPropertyStr(ctx, modules, "length");
    uint32_t count = 0, entry_count = 0;
    int ok = 0;
    JS_ToUint32(ctx, &count, count_value); JS_FreeValue(ctx, count_value);
    count_value = JS_GetPropertyStr(ctx, entries, "length");
    JS_ToUint32(ctx, &entry_count, count_value); JS_FreeValue(ctx, count_value);
    if (count < 1 || count > 64 || entry_count < 1 || entry_count > 64) goto invalid;
    size_t source_bytes = 0;
    for (uint32_t i = 0; i < count; i++) {
        JSValue item = JS_GetPropertyUint32(ctx, modules, i);
        JSValue id = JS_GetPropertyStr(ctx, item, "id");
        JSValue source = JS_GetPropertyStr(ctx, item, "source");
        size_t id_length = 0;
        qx_module *module = &session->modules[i];
        session->module_count = i + 1;
        int strings = JS_IsString(id) && JS_IsString(source);
        if (strings) {
            module->id = JS_ToCStringLen(ctx, &id_length, id);
            module->source = JS_ToCStringLen(ctx, &module->source_length, source);
        }
        JS_FreeValue(ctx, source); JS_FreeValue(ctx, id); JS_FreeValue(ctx, item);
        if (!module->id || !module->source || !valid_module_id(module->id, id_length)) goto invalid;
        for (uint32_t j = 0; j < i; j++) if (!strcmp(session->modules[j].id, module->id)) goto invalid;
        source_bytes += module->source_length;
        if (source_bytes > 1024 * 1024) goto invalid;
    }
    /* A single synthetic root lets the engine own linking, cycles and once-only evaluation. */
    char root[64 * 144 + 1];
    size_t root_length = 0;
    for (uint32_t i = 0; i < entry_count; i++) {
        JSValue entry = JS_GetPropertyUint32(ctx, entries, i);
        size_t length = 0;
        const char *id = JS_IsString(entry) ? JS_ToCStringLen(ctx, &length, entry) : NULL;
        JS_FreeValue(ctx, entry);
        int valid = id && valid_module_id(id, length) && find_module(session, id);
        if (valid) {
            memcpy(root + root_length, "import '", 8); root_length += 8;
            memcpy(root + root_length, id, length); root_length += length;
            memcpy(root + root_length, "';\n", 3); root_length += 3;
        }
        if (id) JS_FreeCString(ctx, id);
        if (!valid) goto invalid;
    }
    root[root_length] = 0;
    session->modules_loading = 1;
    JS_SetModuleLoaderFunc2(session->runtime, normalize_module, load_module, check_import_attributes, session);
    JSValue result = JS_Eval(ctx, root, root_length, "@quickexcel/entries", JS_EVAL_TYPE_MODULE);
    session->modules_loading = 0;
    if (JS_IsException(result)) { exception_error(session, error); goto done; }
    JSPromiseStateEnum state = JS_PromiseState(ctx, result);
    if (state == JS_PROMISE_REJECTED) {
        JS_Throw(ctx, JS_PromiseResult(ctx, result));
        JS_FreeValue(ctx, result);
        exception_error(session, error);
        goto done;
    }
    JS_FreeValue(ctx, result);
    if (state == JS_PROMISE_PENDING || JS_IsJobPending(session->runtime)) {
        set_error(error, "Package modules require synchronous initialization; dynamic import and pending jobs are unsupported.");
        goto done;
    }
    ok = 1;
    goto done;
invalid:
    set_error(error, "Invalid module mapping or source budget.");
done:
    JS_FreeValue(ctx, entries); JS_FreeValue(ctx, modules); JS_FreeValue(ctx, package);
    return ok;
}

QX_EXPORT char *qx_call(qx_session *session, const char *method, const char *arguments_json, char **error) {
    if (error) *error = NULL;
    if (!session) { set_error(error, "JavaScript session is closed."); return NULL; }
    start_operation(session);
    JSContext *context = session->context;
    JSValue arguments = JS_ParseJSON(context, arguments_json, strlen(arguments_json), "arguments.json");
    if (JS_IsException(arguments)) { exception_error(session, error); return NULL; }
    JSValue function = JS_GetPropertyStr(context, session->bridge, method);
    JSValue argv[2] = { JS_GetPropertyUint32(context, arguments, 0), JS_GetPropertyUint32(context, arguments, 1) };
    JSValue result = JS_Call(context, function, session->bridge, 2, argv);
    JS_FreeValue(context, argv[0]);
    JS_FreeValue(context, argv[1]);
    JS_FreeValue(context, function);
    JS_FreeValue(context, arguments);
    if (JS_IsException(result)) { exception_error(session, error); return NULL; }
    if (session->module_count && JS_IsJobPending(session->runtime)) {
        JS_FreeValue(context, result);
        set_error(error, "Asynchronous package module jobs are unsupported.");
        return NULL;
    }
    JSValue json = JS_JSONStringify(context, result, JS_UNDEFINED, JS_UNDEFINED);
    JS_FreeValue(context, result);
    if (JS_IsException(json)) { exception_error(session, error); return NULL; }
    size_t length = 0;
    const char *text = JS_ToCStringLen(context, &length, json);
    char *output = NULL;
    if (text) {
        if (length <= QX_RESULT_LIMIT) {
            output = copy_text(text, length);
            if (!output) set_error(error, "Unable to allocate JavaScript result.");
        } else {
            set_error(error, "JavaScript result exceeds the 16 MiB result limit.");
        }
        JS_FreeCString(context, text);
    } else {
        exception_error(session, error);
    }
    JS_FreeValue(context, json);
    return output;
}

QX_EXPORT void qx_free_string(char *text) { free(text); }

/* Stable status, including an interrupt raised while formatting an exception. */
QX_EXPORT int qx_was_interrupted(qx_session *session) { return session ? session->interrupted : 0; }
