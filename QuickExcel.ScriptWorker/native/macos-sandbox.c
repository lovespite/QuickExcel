// QuickExcel macOS worker containment. Compiled twice: a launcher executable
// and the final guard library loaded by the managed worker. The profile is
// denied-by-default; the worker gets read access to its private payload and
// system libraries, no network, no Mach lookups, no child processes and no
// writable files. Memory and thread ceilings use a native watchdog because
// Darwin rejects RLIMIT_AS/DATA and task footprint limits require private
// entitlements; the limit is enforced, but it is a polling bound, not a
// cgroup-equivalent hard cap.
#include <errno.h>
#include <fcntl.h>
#include <libproc.h>
#include <limits.h>
#include <mach/mach.h>
#include <mach/task_info.h>
#include <mach-o/dyld.h>
#include <poll.h>
#include <pthread.h>
#include <sandbox.h>
#include <signal.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/event.h>
#include <sys/resource.h>
#include <sys/socket.h>
#include <sys/stat.h>
#include <sys/time.h>
#include <unistd.h>

#define WATCHDOG_INTERVAL_MS 25
#define DEFAULT_THREAD_LIMIT 32

// Private but exported libsandbox entry point; the public SDK header omits it.
extern int sandbox_check(pid_t pid, const char *operation, int type, ...);

#ifndef QX_SANDBOX_LAUNCHER
static int qx_denied_read(const char *path) {
    int fd = open(path, O_RDONLY | O_CLOEXEC);
    if (fd >= 0) { close(fd); return 0; }
    return errno == EPERM || errno == EACCES;
}

static int qx_denied_write(const char *directory) {
    char path[PATH_MAX];
    if (snprintf(path, sizeof(path), "%s/.qx-write-probe-%d", directory, getpid()) >= (int)sizeof(path)) return 0;
    int fd = open(path, O_CREAT | O_WRONLY | O_EXCL | O_CLOEXEC, 0600);
    if (fd >= 0) { close(fd); unlink(path); return 0; }
    return errno == EPERM || errno == EACCES || errno == EROFS;
}

static int qx_denied_network(void) {
    int socket_handle = socket(AF_INET, SOCK_STREAM, 0);
    if (socket_handle < 0) return errno == EPERM || errno == EACCES;
    struct sockaddr_in address;
    memset(&address, 0, sizeof(address));
    address.sin_family = AF_INET;
    address.sin_port = htons(9);
    address.sin_addr.s_addr = htonl(0x7f000001);
    int connected = connect(socket_handle, (struct sockaddr *)&address, sizeof(address));
    close(socket_handle);
    return connected < 0 && (errno == EPERM || errno == EACCES);
}

static void qx_exit_now(void) {
    kill(getpid(), SIGKILL);
    _exit(137);
}

typedef struct {
    pid_t parent;
    uint64_t memory_limit;
    int thread_limit;
} qx_watchdog;

static void *qx_watch(void *argument) {
    qx_watchdog *watch = (qx_watchdog *)argument;
    int queue = kqueue();
    if (queue < 0) return NULL;
    struct kevent registration;
    EV_SET(&registration, watch->parent, EVFILT_PROC, EV_ADD | EV_ENABLE, NOTE_EXIT, 0, NULL);
    if (kevent(queue, &registration, 1, NULL, 0, NULL) < 0) return NULL;
    for (;;) {
        struct timespec timeout = {0, WATCHDOG_INTERVAL_MS * 1000 * 1000};
        struct kevent event;
        int changed = kevent(queue, NULL, 0, &event, 1, &timeout);
        if (changed > 0) qx_exit_now();
        if (changed < 0 && errno != EINTR) { dprintf(2, "watchdog: kevent %d\n", errno); qx_exit_now(); }
        if (getppid() != watch->parent) { dprintf(2, "watchdog: parent changed getppid=%d watch=%d\n", getppid(), watch->parent); qx_exit_now(); }
        task_vm_info_data_t info;
        mach_msg_type_number_t count = TASK_VM_INFO_COUNT;
        if (task_info(mach_task_self(), TASK_VM_INFO, (task_info_t)&info, &count) != KERN_SUCCESS) { dprintf(2, "watchdog: task_info\n"); qx_exit_now(); }
        if ((uint64_t)info.phys_footprint > watch->memory_limit) {
            dprintf(2, "watchdog: footprint %llu > %llu\n", (unsigned long long)info.phys_footprint, (unsigned long long)watch->memory_limit);
            qx_exit_now();
        }
        struct proc_taskinfo task;
        if (proc_pidinfo(getpid(), PROC_PIDTASKINFO, 0, &task, sizeof(task)) != sizeof(task)) { dprintf(2, "watchdog: proc_pidinfo\n"); qx_exit_now(); }
        if (task.pti_threadnum > watch->thread_limit) {
            dprintf(2, "watchdog: threads %d > %d\n", task.pti_threadnum, watch->thread_limit);
            qx_exit_now();
        }
    }
}
#endif

#ifdef QX_SANDBOX_LAUNCHER
static int qx_append(char *buffer, size_t capacity, const char *text) {
    size_t used = strlen(buffer), add = strlen(text);
    if (used + add + 1 > capacity) return 0;
    memcpy(buffer + used, text, add + 1);
    return 1;
}

static int qx_append_quoted(char *buffer, size_t capacity, const char *path) {
    if (!qx_append(buffer, capacity, "(literal \"")) return 0;
    for (const char *cursor = path; *cursor; cursor++) {
        char piece[3] = {*cursor, 0, 0};
        if (*cursor == '"' || *cursor == '\\') { piece[0] = '\\'; piece[1] = *cursor; }
        if (!qx_append(buffer, capacity, piece)) return 0;
    }
    return qx_append(buffer, capacity, "\")");
}

static int qx_append_subpath(char *buffer, size_t capacity, const char *path) {
    if (!qx_append(buffer, capacity, "(subpath \"")) return 0;
    for (const char *cursor = path; *cursor; cursor++) {
        char piece[3] = {*cursor, 0, 0};
        if (*cursor == '"' || *cursor == '\\') { piece[0] = '\\'; piece[1] = *cursor; }
        if (!qx_append(buffer, capacity, piece)) return 0;
    }
    return qx_append(buffer, capacity, "\")");
}

static int qx_profile(char *buffer, size_t capacity, const char *payload, int final) {
    buffer[0] = 0;
    if (!qx_append(buffer, capacity,
        "(version 1)\n(deny default)\n(allow file-read-metadata)\n"
        "(allow file-read* (literal \"/\") ")) return 0;
    if (!qx_append_subpath(buffer, capacity, payload)) return 0;
    if (!qx_append(buffer, capacity,
        " (subpath \"/System\") (subpath \"/usr/lib\") (subpath \"/usr/share\") (subpath \"/Library/Apple\")"
        " (subpath \"/private/preboot\") (subpath \"/private/var/db/timezone\") (subpath \"/dev\"))\n"
        "(allow file-map-executable ")) return 0;
    if (!qx_append_subpath(buffer, capacity, payload)) return 0;
    if (!qx_append(buffer, capacity,
        " (subpath \"/System\") (subpath \"/usr/lib\") (subpath \"/Library/Apple\"))\n"
        "(allow file-write-data (literal \"/dev/null\"))\n"
        "(allow sysctl-read)\n(deny mach-lookup)\n(allow signal (target self))\n"
        "(deny network-inbound)\n(deny network-outbound)\n(deny network-bind)\n"
        "(deny process-fork)\n(allow dynamic-code-generation)\n")) return 0;
    if (!final) {
        char executable[PATH_MAX];
        if (snprintf(executable, sizeof(executable), "%s/QuickExcel.ScriptWorker", payload) >= (int)sizeof(executable)) return 0;
        if (!qx_append(buffer, capacity, "(allow process-exec* ")) return 0;
        if (!qx_append_quoted(buffer, capacity, executable)) return 0;
        if (!qx_append(buffer, capacity, ")\n")) return 0;
    } else if (!qx_append(buffer, capacity, "(deny process-exec*)\n")) return 0;
    return 1;
}

static void qx_fail(const char *operation) {
    dprintf(2, "QuickExcel sandbox: %s: %s\n", operation, strerror(errno));
    _exit(125);
}

static void qx_require(int result, const char *operation) {
    if (result < 0) qx_fail(operation);
}

int main(int argc, char **argv) {
    if ((argc != 3 && (argc != 4 || strcmp(argv[3], "compile"))) ||
        getuid() == 0 || getuid() != geteuid() || getgid() != getegid()) return 2;
    const int compilation = argc == 4;
    const pid_t parent = (pid_t)strtol(argv[2], NULL, 10);
    if (parent <= 1 || getppid() != parent) return 125;
    struct rlimit zero = {0, 0}, files = {128, 128}, cpu = {30, 31};
    qx_require(setrlimit(RLIMIT_CORE, &zero), "core limit");
    qx_require(setrlimit(RLIMIT_FSIZE, &zero), "file size limit");
    qx_require(setrlimit(RLIMIT_NOFILE, &files), "descriptor limit");
    qx_require(setrlimit(RLIMIT_CPU, &cpu), "cpu limit");
    for (int descriptor = 3; descriptor < getdtablesize(); descriptor++) close(descriptor);
    struct pollfd start = {.fd = 0, .events = POLLIN};
    char barrier = 0;
    if (poll(&start, 1, 5000) != 1 || read(0, &barrier, 1) != 1 || barrier != 'S') return 125;
    if (getppid() != parent) return 125;
    qx_require(chdir(argv[1]), "private payload");
    char profile[8192];
    if (!qx_profile(profile, sizeof(profile), argv[1], 0)) return 125;
    if (getenv("QX_SHOW_PROFILE")) dprintf(2, "%s\n", profile);
    char *error = NULL;
    if (sandbox_init(profile, 0, &error) != 0) {
        dprintf(2, "QuickExcel sandbox: profile rejected: %s\n", error ? error : "(unknown)");
        sandbox_free_error(error);
        return 125;
    }
    char *arguments[] = {"QuickExcel.ScriptWorker", compilation ? "--compile" : NULL, NULL};
    char *environment[] = {
        "DOTNET_EnableDiagnostics=0", "COMPlus_EnableDiagnostics=0",
        compilation ? "DOTNET_GCHeapHardLimit=10000000" : "DOTNET_GCHeapHardLimit=8000000",
        "DOTNET_GCConserveMemory=9", "DOTNET_PROCESSOR_COUNT=2",
        compilation ? "QX_PROCESS_LIMIT_MIB=1024" : "QX_PROCESS_LIMIT_MIB=256",
        "QX_THREAD_LIMIT=32", "LANG=C.UTF-8", "LC_ALL=C.UTF-8", NULL};
    execve("./QuickExcel.ScriptWorker", arguments, environment);
    qx_fail("exec worker");
}
#else
// The host queries a launched pid through these fixed-arity wrappers. Calling
// the variadic sandbox_check directly from P/Invoke passes path arguments under
// the wrong ABI on arm64.
__attribute__((visibility("default"))) int qx_sandbox_check_path(int pid, const char *operation, const char *path) {
    return sandbox_check((pid_t)pid, operation, 1, path);
}

__attribute__((visibility("default"))) int qx_sandbox_check(int pid, const char *operation) {
    return sandbox_check((pid_t)pid, operation, 0);
}

int qx_macos_lock_worker(void) {
    char executable[PATH_MAX];
    uint32_t size = sizeof(executable);
    if (_NSGetExecutablePath(executable, &size) != 0) return -1;
    char resolved[PATH_MAX];
    if (realpath(executable, resolved) == NULL) return -1;
    char *slash = strrchr(resolved, '/');
    if (slash == NULL) return -1;
    *slash = 0;
    if (!qx_denied_read("/etc/passwd")) return -2;
    if (!qx_denied_write(resolved)) return -3;
    if (!qx_denied_network()) return -4;
    // Darwin permits only one effective profile for this lifecycle; the launcher
    // applied the strict profile before exec. Verify it here instead of trying to
    // replace it, then keep the watchdog outside any managed runtime control.
    const char *limit = getenv("QX_PROCESS_LIMIT_MIB");
    const char *threads = getenv("QX_THREAD_LIMIT");
    static qx_watchdog watch;
    watch.parent = getppid();
    watch.memory_limit = (uint64_t)(limit ? strtoull(limit, NULL, 10) : 256) * 1024 * 1024;
    watch.thread_limit = threads ? atoi(threads) : DEFAULT_THREAD_LIMIT;
    if (watch.parent <= 1 || watch.memory_limit == 0 || watch.thread_limit <= 0) return -7;
    pthread_t thread;
    if (pthread_create(&thread, NULL, qx_watch, &watch) != 0) return -8;
    pthread_detach(thread);
    return 0;
}
#endif
