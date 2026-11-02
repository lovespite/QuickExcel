# Platform integration probes (development only)

This tool is never referenced by a production project and must never be packaged.
Build on/for Linux x64 with `QuickExcelNativeRid=linux-x64`, self-contained runtime,
and `--artifacts-path` under this repository. Build the normal bridge and sandbox
first, then `CC=clang-18 bash tools/QuickExcel.PlatformSmoke/native/build-linux.sh`.
Copy only its probe executable into `TestProbe/QuickExcel.ScriptWorker` beside the
development tool; keep the normal worker in `ScriptWorker/`.

Run as an ordinary user with an active systemd user manager and cgroup v2:

- `QuickExcel.PlatformSmoke NEW_DIRECTORY`: v1/v2, batching, actions, limits, recovery.
- `QuickExcel.PlatformSmoke --sandbox NEW_DIRECTORY`: actual native policy, thread,
  committed memory and lifetime attacks through the production launcher and guard.
- `python3 scripts/test-linux-worker-lifetime.py ABSOLUTE_TOOL NEW_DIRECTORY`:
  terminate the owned host only, check worker exit and safe journal recovery.
- `dbus-run-session -- python3 scripts/test-linux-secret-service.py ABSOLUTE_TOOL NEW_DIRECTORY`:
  fresh private keyring with a stdin-only random password; positive and unavailable service.
- `python3 scripts/test-cups-handoff.py ABSOLUTE_TOOL NEW_DIRECTORY`: owned non-root
  CUPS/IPP services on free ports 8631/8632, generated PDF, media checks and Poppler geometry.
  Requires cups-daemon, cups-client, cups-ipp-utils, cups-filters, poppler-utils and
  running Avahi for the IPP emulator. Never submits to real printers or system queues.

`--hold` and `--recover` support the lifecycle fixture; `--bootstrap` is diagnostic.
All failures and resource limits remain in result JSON and logs. These are engineering
checks, not proof of physical printing, IME, accessibility, Wayland or clean devices.
