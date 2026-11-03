#define _GNU_SOURCE
#include <errno.h>
#include <fcntl.h>
#include <linux/sched.h>
#include <pthread.h>
#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/mman.h>
#include <sys/prctl.h>
#include <sys/resource.h>
#include <sys/socket.h>
#include <sys/syscall.h>
#include <unistd.h>

extern int qx_linux_lock_worker(void);
static int count=0;
static void denied(long result, int expected, const char *name) {
    if(result != -1 || errno != expected) { dprintf(2,"FAIL %s result=%ld errno=%d\n",name,result,errno); exit(1); }
    count++;
}
static void *wait_thread(void *unused) { (void)unused; for(;;) sleep(1); return NULL; }
int main(int argc, char **argv) {
    int compilation=argc==2 && !strcmp(argv[1],"--compile");
    const char *managed_limit=getenv("DOTNET_GCHeapHardLimit");
    if(!managed_limit || strcmp(managed_limit,compilation ? "10000000" : "4000000")) return 3;
    if(qx_linux_lock_worker() != 0) return 125;
    char mode[32]={0},canary[4096]={0};
    FILE *configuration=fopen("probe.mode","r"); if(!configuration) return 2;
    if(!fgets(mode,sizeof(mode),configuration)) return 2;
    if(!fgets(canary,sizeof(canary),configuration)) return 2;
    fclose(configuration); mode[strcspn(mode,"\n")]=0; canary[strcspn(canary,"\n")]=0;
    printf("READY %s\n",mode); fflush(stdout);
    if(getchar()!='P') return 2;
    if(!strcmp(mode,"memory")) {
        size_t bytes=(compilation ? 1088UL : 320UL)*1024*1024;
        volatile char *memory=malloc(bytes); if(!memory) return 3;
        for(size_t i=0;i<bytes;i+=4096) memory[i]=42;
        return 4; // Surviving this allocation is a failed memory-containment test.
    }
    if(!strcmp(mode,"timeout") || !strcmp(mode,"parent")) {
        for(;;) sleep(1);
    }
    if(!strcmp(mode,"threads")) {
        pthread_t threads[64]; int made=0,result=0;
        while(made<64 && (result=pthread_create(&threads[made],NULL,wait_thread,NULL))==0) made++;
        if(result!=EAGAIN || made>31 || made<1) {dprintf(2,"thread bound %d result=%d\n",made,result);return 1;}
        printf("PASS threads %d\n",made+1);return 0;
    }
    denied(open(canary,O_RDONLY),EACCES,"outside read");
    denied(open(canary,O_WRONLY),EACCES,"outside write");
    denied(open("probe.mode",O_WRONLY),EACCES,"payload write");
    denied(open("/etc/passwd",O_RDONLY),EACCES,"system file outside allowlist");
    denied(open("/home",O_RDONLY|O_DIRECTORY),EACCES,"home enumeration");
    denied(open("/tmp/qx-sandbox-must-not-create",O_CREAT|O_WRONLY,0600),EACCES,"file creation");
    denied(socket(AF_INET,SOCK_STREAM,0),EPERM,"IPv4 TCP");
    denied(socket(AF_INET,SOCK_DGRAM,0),EPERM,"IPv4 UDP and DNS");
    denied(socket(AF_INET6,SOCK_STREAM,0),EPERM,"IPv6 TCP");
    denied(socket(AF_UNIX,SOCK_STREAM,0),EPERM,"Unix sockets");
    denied(syscall(SYS_fork),EPERM,"fork");
    denied(syscall(SYS_vfork),EPERM,"vfork");
    denied(syscall(SYS_clone,SIGCHLD,0,0,0,0),EPERM,"process clone");
    denied(syscall(SYS_clone,CLONE_VM|CLONE_SIGHAND|CLONE_THREAD|CLONE_NEWUSER,0,0,0,0),EPERM,"namespace clone");
    struct clone_args clone={0}; denied(syscall(SYS_clone3,&clone,sizeof(clone)),ENOSYS,"clone3");
    char *arguments[]={"QuickExcel.ScriptWorker",NULL}; char *environment[]={NULL};
    denied(execve("./QuickExcel.ScriptWorker",arguments,environment),EPERM,"repeat exec");
    denied(execve("/bin/sh",arguments,environment),EPERM,"system exec");
    denied(syscall(SYS_execveat,-1,"",arguments,environment,0),EPERM,"execveat");
    denied(syscall(SYS_ptrace,0,0,0,0),EPERM,"ptrace");
    denied(syscall(SYS_process_vm_readv,getppid(),NULL,0,NULL,0,0),EPERM,"process memory");
    denied(syscall(SYS_pidfd_getfd,-1,0,0),EPERM,"descriptor theft");
    denied(syscall(SYS_io_uring_setup,1,NULL),EPERM,"io_uring");
    denied(syscall(SYS_open_by_handle_at,-1,NULL,0),EPERM,"filesystem handle bypass");
    denied(syscall(SYS_keyctl,0,0,0,0,0),EPERM,"keyring");
    denied(syscall(SYS_bpf,0,NULL,0),EPERM,"bpf");
    denied(syscall(SYS_unshare,CLONE_NEWUSER),EPERM,"unshare");
    denied(prctl(PR_SET_PDEATHSIG,0),EPERM,"parent death override");
    denied(prctl(PR_SET_DUMPABLE,1),EPERM,"core dump override");
    struct rlimit limit={1024,1024}; denied(setrlimit(RLIMIT_NOFILE,&limit),EPERM,"rlimit override");
    denied(syscall(SYS_prlimit64,0,RLIMIT_NOFILE,&limit,NULL),EPERM,"prlimit override");
    denied(kill(getppid(),0),EPERM,"outside signal");
    denied(syscall(SYS_tgkill,getppid(),getppid(),0),EPERM,"outside thread signal");
    printf("PASS policy %d\n",count); return 0;
}
