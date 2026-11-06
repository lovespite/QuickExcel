QuickExcel Avalonia — platform package

Linux x64 tested baseline: Ubuntu 24.04.5, kernel 6.8, systemd 255 with an ordinary
user session and unified cgroup v2. Needs libwebkit2gtk-4.1-0, libgtk-3-0,
libfontconfig1, libfreetype6, libicu74, libssl3t64 and fonts-noto-cjk.
Printing needs libcups2t64, cups-client and a configured reachable CUPS server;
server-side PDF conversion normally needs cups-filters. Persistent script trust
needs libsecret-tools and an unlocked Secret Service (such as GNOME Keyring).

Extract the archive and run ./Start-QuickExcel.sh in your desktop session.
For installation diagnosis run ./Start-QuickExcel.sh --self-check /absolute/new/path.
This uses isolated settings and files and does not print or change user trust.
No .NET SDK, Node or development tools are needed on the destination machine.

Linux script execution requires x64, Landlock ABI >= 3, seccomp, a systemd user
bus, and enforceable memory/swap/tasks cgroup limits. Every worker launch verifies
these requirements before sending workbook code. Unsupported environments fail
closed: there is no in-process or same-permission script fallback. Do not disable
AppArmor, browser sandboxing or kernel security to make the application run.

macOS arm64 ships the pinned JSC engine and the isolated worker. Execution uses a
denied-by-default Seatbelt profile, host-verified file/write/network/fork denials,
a parent-death watchdog and polling footprint/thread watchdogs; the 256 MiB bound
is a polling limit with bounded overshoot, not a cgroup-equivalent hard cap.
Persistent trust needs a signed bundle with an application identifier; without it
only explicit one-time consent is allowed (the Keychain returns -34018). WKWebView,
Keychain fail-closed behavior and CUPS print preview were exercised locally; signing,
notarization, signed-bundle trust and physical acceptance remain unverified.

Physical IME, DPI/multi-screen, accessibility, printer and clean-device acceptance
are deferred, not passed. Linux Chinese IME had no candidate window in the prior
VM experiment. See repository docs/67-avalonia-platforms.md for current evidence.
Keep release.json, metadata and license files with the package.

JavaScriptCore is the only script engine. JS API v1 remains compatible; API v2
modules can use JavaScript or TypeScript. TypeScript is fully checked by the
isolated compiler before Apply/Run; failed checks keep the draft and previous
applied source. Files save the original source. Older packages reject the
required typescript-sources feature: keep an original file backup before migration.
Compilation uses a separate 1 GiB / 5-second native / 6-second host budget;
execution retains 256 MiB process / 256 KiB JS stack / 1-second operation / 5-second round.
OPEN JSC-MEM-001: the 64 MiB total engine allocation hard limit is not implemented.
It is explicitly deferred and is not a progression gate. A process limit or GC
threshold is not an equivalent engine limit. JSC-Licenses contains the upstream
notices, corresponding sources and instructions for rebuilding/replacing libraries.
Rollback replaces the whole host/worker package; never mix engine or IPC versions.
