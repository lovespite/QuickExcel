#!/usr/bin/env bash
set -euo pipefail
here=$(cd -- "$(dirname -- "$0")" && pwd)
exec bash "$here/../../../tools/QuickExcel.JscProbe/build-macos.sh" "$@"
