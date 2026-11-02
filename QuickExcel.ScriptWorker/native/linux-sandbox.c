#define _GNU_SOURCE
#include <errno.h>
#include <fcntl.h>
#include <linux/audit.h>
#include <linux/filter.h>
#include <linux/landlock.h>
#include <linux/seccomp.h>
#include <poll.h>
#include <sched.h>
#include <signal.h>
#include <stddef.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/prctl.h>
#include <sys/resource.h>
#include <sys/socket.h>
#include <sys/stat.h>
#include <sys/syscall.h>
#include <unistd.h>

#if !defined(__x86_64__)
#error Linux sandbox currently requires x86_64; other ABIs need independent policy verification.
#endif

#ifdef QX_SANDBOX_LAUNCHER
static void fail(const char *operation) { dprintf(2, "QuickExcel sandbox: %s: %s\n", operation, strerror(errno)); _exit(125); }
static void require(int result, const char *operation) { if (result < 0) fail(operation); }
#endif

#define ALLOW(call) BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_##call, 0, 1), BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW)
#define DENY BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ERRNO|EPERM)

// Default deny, including new syscalls and x32. Only threads in this process may
// be created; no child, namespace, network, IPC, tracing or mount operations.
static int filter(int final) {
    const unsigned thread_required = CLONE_VM|CLONE_SIGHAND|CLONE_THREAD;
    const unsigned thread_allowed = thread_required|CLONE_FS|CLONE_FILES|CLONE_SYSVSEM|CLONE_SETTLS|CLONE_PARENT_SETTID|CLONE_CHILD_SETTID|CLONE_CHILD_CLEARTID;
    struct sock_filter rules[] = {
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, arch)),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, AUDIT_ARCH_X86_64, 1, 0),
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_KILL_PROCESS),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, nr)),
        BPF_JUMP(BPF_JMP|BPF_JSET|BPF_K, 0x40000000, 0, 1),
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_KILL_PROCESS),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_clone3, 0, 1),
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ERRNO|ENOSYS),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_clone, 0, 10),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[0])+4),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, 0, 1, 0), DENY,
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[0])),
        BPF_JUMP(BPF_JMP|BPF_JSET|BPF_K, ~thread_allowed, 0, 1), DENY,
        BPF_STMT(BPF_ALU|BPF_AND|BPF_K, thread_required),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, thread_required, 1, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_tgkill, 0, 4),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[0])),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, (unsigned)getpid(), 1, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_prlimit64, 0, 6),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[2])),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, 0, 0, 2),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[2])+4),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, 0, 1, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_sched_setaffinity, 0, 4),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[0])),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, 0, 1, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        // prctl cannot change parent-death behavior or drop protections.
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_prctl, 0, 12),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[0])),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, PR_SET_NAME, 5, 0),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, PR_GET_NAME, 4, 0),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, PR_GET_NO_NEW_PRIVS, 3, 0),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, PR_GET_SECCOMP, 2, 0),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, PR_SET_DUMPABLE, 2, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        BPF_STMT(BPF_LD|BPF_W|BPF_ABS, offsetof(struct seccomp_data, args[1])),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, 0, 1, 0), DENY,
        BPF_STMT(BPF_RET|BPF_K, SECCOMP_RET_ALLOW),
        BPF_JUMP(BPF_JMP|BPF_JEQ|BPF_K, __NR_execve, 0, 1),
        BPF_STMT(BPF_RET|BPF_K, final ? SECCOMP_RET_ERRNO|EPERM : SECCOMP_RET_ALLOW),
        ALLOW(read), ALLOW(write), ALLOW(close), ALLOW(fstat), ALLOW(newfstatat), ALLOW(stat), ALLOW(lstat),
        ALLOW(lseek), ALLOW(pread64), ALLOW(pwrite64), ALLOW(readv), ALLOW(writev), ALLOW(openat), ALLOW(open),
        ALLOW(access), ALLOW(faccessat), ALLOW(faccessat2), ALLOW(readlink), ALLOW(readlinkat), ALLOW(getdents64),
        ALLOW(fcntl), ALLOW(fadvise64), ALLOW(statfs), ALLOW(fstatfs), ALLOW(mmap), ALLOW(mprotect), ALLOW(munmap),
        ALLOW(brk), ALLOW(mremap), ALLOW(madvise), ALLOW(msync), ALLOW(mincore),
        ALLOW(rt_sigaction), ALLOW(rt_sigprocmask), ALLOW(rt_sigreturn), ALLOW(sigaltstack), ALLOW(rt_sigsuspend), ALLOW(rt_sigtimedwait),
        ALLOW(getpid), ALLOW(getppid), ALLOW(gettid), ALLOW(getuid), ALLOW(geteuid), ALLOW(getgid), ALLOW(getegid),
        ALLOW(getgroups), ALLOW(getresuid), ALLOW(getresgid), ALLOW(uname), ALLOW(sysinfo), ALLOW(sched_yield),
        ALLOW(sched_get_priority_min), ALLOW(sched_get_priority_max), ALLOW(getcwd),
        ALLOW(sched_getaffinity), ALLOW(sched_getscheduler), ALLOW(sched_getparam), ALLOW(getcpu),
        ALLOW(futex), ALLOW(futex_waitv), ALLOW(set_tid_address), ALLOW(set_robust_list), ALLOW(rseq), ALLOW(membarrier),
        ALLOW(restart_syscall), ALLOW(clock_gettime), ALLOW(clock_getres), ALLOW(gettimeofday), ALLOW(nanosleep),
        ALLOW(clock_nanosleep), ALLOW(times), ALLOW(getrusage), ALLOW(getrlimit), ALLOW(getrandom), ALLOW(arch_prctl),
        ALLOW(eventfd2), ALLOW(epoll_create1), ALLOW(epoll_ctl), ALLOW(epoll_wait), ALLOW(epoll_pwait), ALLOW(epoll_pwait2),
        ALLOW(poll), ALLOW(ppoll), ALLOW(select), ALLOW(pselect6), ALLOW(pipe2), ALLOW(pipe), ALLOW(dup), ALLOW(dup2), ALLOW(dup3),
        ALLOW(close_range), ALLOW(seccomp), ALLOW(capget), ALLOW(exit), ALLOW(exit_group), DENY
    };
    struct sock_fprog program = { .len = sizeof(rules)/sizeof(rules[0]), .filter = rules };
    return syscall(SYS_seccomp, SECCOMP_SET_MODE_FILTER, final ? SECCOMP_FILTER_FLAG_TSYNC : 0, &program);
}

#ifdef QX_SANDBOX_LAUNCHER
static void read_path(int rules, const char *path, int executable, int optional) {
    int fd = open(path, O_PATH|O_CLOEXEC);
    if (fd < 0 && optional && errno == ENOENT) return;
    require(fd, path);
    struct stat state; require(fstat(fd, &state), "fstat");
    struct landlock_path_beneath_attr entry = { .parent_fd = fd, .allowed_access = LANDLOCK_ACCESS_FS_READ_FILE };
    if (S_ISDIR(state.st_mode)) entry.allowed_access |= LANDLOCK_ACCESS_FS_READ_DIR;
    if (executable) entry.allowed_access |= LANDLOCK_ACCESS_FS_EXECUTE;
    require(syscall(SYS_landlock_add_rule, rules, LANDLOCK_RULE_PATH_BENEATH, &entry, 0), path);
    close(fd);
}

int main(int argc, char **argv) {
    if (argc != 3 || getuid() == 0 || getuid() != geteuid() || getgid() != getegid()) return 2;
    const pid_t parent = (pid_t)strtol(argv[2], NULL, 10);
    require(prctl(PR_SET_PDEATHSIG, SIGKILL), "parent death");
    if (parent <= 1 || getppid() != parent) return 125;
    require(prctl(PR_SET_DUMPABLE, 0), "core dump protection");
    struct rlimit zero = {0,0}, files = {128,128};
    require(setrlimit(RLIMIT_CORE, &zero), "core limit");
    require(setrlimit(RLIMIT_FSIZE, &zero), "file size limit");
    require(setrlimit(RLIMIT_NOFILE, &files), "descriptor limit");
    require(syscall(SYS_close_range, 3u, ~0u, 0), "close inherited descriptors");
    // Host attaches this exact, still trusted PID to a bounded systemd scope,
    // verifies cgroup files, and only then supplies the one-byte start barrier.
    struct pollfd start = { .fd=0, .events=POLLIN }; char byte=0;
    if (poll(&start,1,5000) != 1 || read(0,&byte,1) != 1 || byte != 'S') return 125;
    require(chdir(argv[1]), "private payload");
    require(prctl(PR_SET_NO_NEW_PRIVS,1,0,0,0), "no_new_privs");
    int abi = syscall(SYS_landlock_create_ruleset,NULL,0,LANDLOCK_CREATE_RULESET_VERSION);
    if (abi < 3) { errno=ENOTSUP; fail("Landlock ABI >= 3 required"); }
    // Handle every filesystem right through ABI 3, including truncate and refer.
    struct landlock_ruleset_attr attributes = { .handled_access_fs = (1ULL << 15)-1 };
    int rules = syscall(SYS_landlock_create_ruleset,&attributes,sizeof(attributes),0);
    require(rules,"Landlock ruleset");
    read_path(rules,argv[1],0,0);
    read_path(rules,"QuickExcel.ScriptWorker",1,0);
    // execve also checks the ELF interpreter. Only this exact system loader
    // gains execute permission; the final synchronized filter denies all exec.
    read_path(rules,"/lib64/ld-linux-x86-64.so.2",1,0);
    read_path(rules,"/usr/lib",0,0); read_path(rules,"/lib",0,0); read_path(rules,"/lib64",0,1);
    read_path(rules,"/etc/ld.so.cache",0,1); read_path(rules,"/etc/localtime",0,1);
    read_path(rules,"/dev/urandom",0,0); read_path(rules,"/dev/random",0,0);
    read_path(rules,"/sys/devices/system/cpu",0,0);
    const char *proc[] = {"/proc/self/maps","/proc/self/status","/proc/self/stat","/proc/self/statm","/proc/self/mountinfo","/proc/self/cgroup","/proc/meminfo","/proc/cpuinfo",NULL};
    for(int i=0;proc[i];i++) read_path(rules,proc[i],0,0);
    require(syscall(SYS_landlock_restrict_self,rules,0),"Landlock enforcement"); close(rules);
    require(filter(0),"seccomp launch filter");
    char *arguments[] = {"QuickExcel.ScriptWorker",NULL};
    char *environment[] = {"DOTNET_EnableDiagnostics=0","COMPlus_EnableDiagnostics=0","DOTNET_GCHeapHardLimit=4000000","DOTNET_GCConserveMemory=9","DOTNET_PROCESSOR_COUNT=2","LANG=C.UTF-8","LC_ALL=C.UTF-8",NULL};
    execve("./QuickExcel.ScriptWorker",arguments,environment); fail("exec worker");
}
#else
__attribute__((visibility("default"))) int qx_linux_lock_worker(void) {
    // exec resets dumpability; restore it before any request is read. The
    // filter permits only setting zero, so native code cannot re-enable dumps.
    if (prctl(PR_SET_DUMPABLE,0) != 0) return -7;
    if (prctl(PR_GET_NO_NEW_PRIVS,0,0,0,0) != 1 || prctl(PR_GET_SECCOMP,0,0,0,0) != 2) return -1;
    int denied = open("/etc/passwd",O_RDONLY|O_CLOEXEC);
    if (denied >= 0) {close(denied);return -2;} if(errno != EACCES) return -3;
    int network = socket(AF_INET,SOCK_STREAM,0);
    if (network >= 0) {close(network);return -4;} if(errno != EPERM) return -5;
    // TSYNC must return exactly zero; a positive result is an unsynchronized TID.
    return filter(1) == 0 ? 0 : -6;
}
#endif
