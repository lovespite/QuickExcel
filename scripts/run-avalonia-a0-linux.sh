#!/usr/bin/env bash
# Run the packaged A0 probe in an existing Linux display session, as an ordinary user.
# Bundle layout: bin/, web/, fonts/. Neither Node nor a .NET SDK is needed at runtime.
set -euo pipefail
if [[ $# != 2 || $(id -u) == 0 || -z ${DISPLAY:-} ]]; then
  echo 'Usage (non-root, DISPLAY set): run-avalonia-a0-linux.sh BUNDLE NEW_OUTPUT_DIRECTORY' >&2
  exit 2
fi
bundle=$(realpath "$1")
mkdir "$2" # Never overwrite evidence from an earlier run.
output=$(realpath "$2")
test -x "$bundle/bin/QuickExcel.AvaloniaProbe"
test -f "$bundle/web/manifest.json"
test -f "$bundle/fonts/NotoSansCJKsc-Regular.otf"
set +e
timeout --kill-after=5s 60s "$bundle/bin/QuickExcel.AvaloniaProbe" "$output" \
  --web "$bundle/web" --fonts "$bundle/fonts" >"$output/stdout.log" 2>"$output/stderr.log"
code=$?
set -e
python3 - "$output" "$code" <<'PY'
import json, os, platform, pathlib, sys
output = pathlib.Path(sys.argv[1])
code = int(sys.argv[2])
result, error = {}, None
try:
    result = json.loads((output / 'result.json').read_text())
except (OSError, ValueError) as failure:
    error = str(failure)
passed = code == 0 and result.get('passed') is True and len(result.get('checks', [])) == 3
(output / 'process.json').write_text(json.dumps({
    'schemaVersion': 1, 'passed': passed, 'exitCode': code, 'resultError': error,
    'display': os.environ.get('DISPLAY'), 'session': os.environ.get('XDG_SESSION_TYPE'),
    'kernel': platform.platform(), 'uid': os.getuid(),
}, indent=2))
print(('PASS' if passed else 'FAIL') + ': automated checks (desktop IME/DPI acceptance is separate): ' + str(output))
sys.exit(0 if passed else 1)
PY
