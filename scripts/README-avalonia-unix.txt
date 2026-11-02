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

macOS arm64 is an engineering preview only. Script execution is disabled and no
worker is shipped until its separate platform sandbox is implemented and qualified.
WKWebView, Keychain, native binaries, printing, signing and notarization require
macOS validation; a cross-build is not evidence of working Mac support.

Physical IME, DPI/multi-screen, accessibility, printer and clean-device acceptance
are deferred, not passed. Linux Chinese IME had no candidate window in the prior
VM experiment. See repository docs/67-avalonia-platforms.md for current evidence.
Keep release.json, metadata and license files with the package.
