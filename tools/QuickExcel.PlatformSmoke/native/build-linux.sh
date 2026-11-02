#!/usr/bin/env bash
set -euo pipefail
native=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$native/../../.." && pwd)
mkdir -p "$root/artifacts/native/linux-x64/probe"
${CC:-cc} -std=c11 -O2 -Wall -Wextra -Werror -pthread "$native/linux-probe.c" -L"$root/artifacts/native/linux-x64" -lquickexcel-sandbox '-Wl,-rpath,$ORIGIN' -o "$root/artifacts/native/linux-x64/probe/QuickExcel.ScriptWorker"
