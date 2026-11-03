// J0 development probe: trusted repository fixtures only, NOT a sandbox/worker.
#include "config.h"
#include <JavaScriptCore/JavaScript.h>
#include <JavaScriptCore/JSBasePrivate.h>
#include <JavaScriptCore/JSContextRefPrivate.h>
#include <JavaScriptCore/Options.h>
#include <chrono>
#include <cstdlib>
#include <fstream>
#include <iostream>
#include <iterator>
#include <string>
#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#endif

static bool timedOut = false;
static bool cpuWatchdog = false;
static std::chrono::steady_clock::time_point deadline;
static bool terminate(JSContextRef context, void*)
{
    if (!cpuWatchdog && std::chrono::steady_clock::now() < deadline) {
        // Explicitly re-arm: returning false alone does not reliably re-arm this
        // pinned JSC implementation after its first callback.
        JSContextGroupSetExecutionTimeLimit(JSContextGetGroup(context), 0.01, terminate, nullptr);
        return false;
    }
    timedOut = true;
    return true;
}

static std::string utf8(JSStringRef value)
{
    if (!value)
        return {};
    const auto length = JSStringGetMaximumUTF8CStringSize(value);
    // Trusted fixtures only. This is deliberately not a product value bridge.
    if (length > 16 * 1024 * 1024)
        return "\"probe-output-over-limit\"";
    std::string result(length, '\0');
    const auto written = JSStringGetUTF8CString(value, result.data(), length);
    result.resize(written ? written - 1 : 0);
    return result;
}

int main(int argc, char** argv)
{
    if (argc < 3) {
        std::cerr << "Usage: quickexcel-jsc-probe seconds trusted-fixture.js [...]\n";
        return 64;
    }
    const double seconds = std::stod(argv[1]);
    if (!(seconds > 0 && seconds <= 30))
        return 64;
    size_t processLimitBytes = 0;
#ifdef _WIN32
    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOGPFAULTERRORBOX);
    if (const char* limit = std::getenv("QX_PROBE_PROCESS_LIMIT_MIB")) {
        const auto mib = std::stoull(limit);
        if (mib < 16 || mib > 1024) return 64;
        processLimitBytes = mib * 1024 * 1024;
        // Development self-containment experiment, not an AppContainer launcher.
        // The unnamed handle is retained for this process lifetime and reclaimed by the OS.
        auto job = CreateJobObjectW(nullptr, nullptr);
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION limits {};
        limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_PROCESS_MEMORY;
        limits.ProcessMemoryLimit = processLimitBytes;
        BOOL inJob = FALSE;
        if (!job || !SetInformationJobObject(job, JobObjectExtendedLimitInformation, &limits, sizeof(limits)) ||
            !AssignProcessToJobObject(job, GetCurrentProcess()) ||
            !QueryInformationJobObject(job, JobObjectExtendedLimitInformation, &limits, sizeof(limits), nullptr) ||
            limits.ProcessMemoryLimit != processLimitBytes ||
            !IsProcessInJob(GetCurrentProcess(), job, &inJob) || !inJob) {
            std::cerr << "probe-job-setup-failed\n";
            return 68;
        }
        std::cerr << "probe-job-limit-installed=" << processLimitBytes << "\n";
    }
#endif
    cpuWatchdog = std::getenv("QX_PROBE_CPU_WATCHDOG") != nullptr;
    deadline = std::chrono::steady_clock::now() + std::chrono::duration_cast<std::chrono::steady_clock::duration>(
        std::chrono::duration<double>(seconds));
    auto group = JSContextGroupCreate();
    // JSC's time limit measures CPU time. Poll a monotonic wall deadline through
    // the native watchdog; no user JavaScript timer/microtask is involved.
    JSContextGroupSetExecutionTimeLimit(group, cpuWatchdog ? seconds : 0.01, terminate, nullptr);
    auto context = JSGlobalContextCreateInGroup(group, nullptr);
    const auto start = std::chrono::steady_clock::now();
    int exitCode = 0;
    std::string output = "null";
    for (int i = 2; i < argc; ++i) {
        std::ifstream stream(argv[i], std::ios::binary);
        if (!stream) { exitCode = 66; break; }
        const std::string source((std::istreambuf_iterator<char>(stream)), {});
        if (source.size() > 16 * 1024 * 1024) { exitCode = 65; break; }
        auto script = JSStringCreateWithUTF8CString(source.c_str());
        auto origin = JSStringCreateWithUTF8CString("quickexcel-probe");
        JSValueRef exception = nullptr;
        auto value = JSEvaluateScript(context, script, nullptr, origin, 1, &exception);
        JSStringRelease(origin);
        JSStringRelease(script);
        if (exception || !value || timedOut) {
            // Never read a thrown object's stack/message/toString getter.
            exitCode = timedOut ? 124 : 65;
            break;
        }
        if (i == argc - 1) {
            auto json = JSValueCreateJSONString(context, value, 0, &exception);
            if (exception || timedOut) { exitCode = timedOut ? 124 : 65; break; }
            output = json ? utf8(json) : "null";
            if (json) JSStringRelease(json);
        }
    }
    const auto elapsed = std::chrono::duration<double, std::milli>(
        std::chrono::steady_clock::now() - start).count();
    size_t peakBytes = 0, peakCommitBytes = 0, currentCommitBytes = 0;
#ifdef _WIN32
    PROCESS_MEMORY_COUNTERS counters {};
    counters.cb = sizeof(counters);
    if (GetProcessMemoryInfo(GetCurrentProcess(), &counters, sizeof(counters))) {
        peakBytes = counters.PeakWorkingSetSize;
        peakCommitBytes = counters.PeakPagefileUsage;
        currentCommitBytes = counters.PagefileUsage;
    }
#endif
    std::cout << "{\"exitCode\":" << exitCode << ",\"interrupted\":"
              << (timedOut ? "true" : "false") << ",\"evalMilliseconds\":" << elapsed
              << ",\"peakWorkingSetBytes\":" << peakBytes << ",\"peakCommitBytes\":" << peakCommitBytes
              << ",\"currentCommitBytes\":" << currentCommitBytes << ",\"processLimitBytes\":" << processLimitBytes
              << ",\"jitReservationOverrideBytes\":" << JSC::Options::jitMemoryReservationSize()
              << ",\"gcMaxHeapSize\":" << JSC::Options::gcMaxHeapSize()
              << ",\"maxPerThreadStackUsage\":" << JSC::Options::maxPerThreadStackUsage()
              << ",\"result\":" << output << "}\n";
    JSGlobalContextRelease(context);
    JSContextGroupRelease(group);
    return exitCode;
}
