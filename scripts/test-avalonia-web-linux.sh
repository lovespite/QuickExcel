#!/usr/bin/env bash
# Run an explicitly published test payload, as a normal user in an existing desktop/Xvfb session.
set -euo pipefail
if [[ $# != 2 ]]; then echo 'Usage: test-avalonia-web-linux.sh APPLICATION_DIRECTORY NEW_EVIDENCE_DIRECTORY' >&2; exit 2; fi
if [[ $(id -u) == 0 ]]; then echo 'Run as an ordinary user; WebKit sandbox must stay enabled.' >&2; exit 2; fi
if [[ -z ${DISPLAY:-}${WAYLAND_DISPLAY:-} ]]; then echo 'A graphical session is required (Xvfb is engineering-only evidence).' >&2; exit 2; fi
if [[ -n ${WEBKIT_DISABLE_SANDBOX_THIS_IS_DANGEROUS:-} ]]; then echo 'Refusing disabled WebKit sandbox.' >&2; exit 2; fi
application=$(realpath -e -- "$1")
evidence=$(realpath -m -- "$2")
if [[ -e $evidence ]]; then echo 'Evidence directory must be new.' >&2; exit 2; fi
mkdir -p -- "$evidence"
uname -a > "$evidence/environment.txt"
sha256sum "$application/QuickExcel.Avalonia.dll" "$application/Web/manifest.json" > "$evidence/payload.sha256"
timeout --kill-after=5s 180s "$application/QuickExcel.Avalonia" --web-smoke "$evidence" > "$evidence/process.log" 2>&1
python3 - "$evidence/result.json" <<'PY'
import json, sys
with open(sys.argv[1]) as stream:
    result = json.load(stream)
assert result['passed'], result
print(f"PASS: {len(result['checks'])} checks. Skipped: {result['skipped']}")
PY
