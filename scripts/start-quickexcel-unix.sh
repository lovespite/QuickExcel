#!/usr/bin/env bash
set -euo pipefail
application=$(cd "$(dirname "$0")" && pwd)
if [[ "$(uname -s)" == Linux ]]; then
  if [[ $(id -u) == 0 ]]; then echo 'Use an ordinary desktop user; WebKit and script isolation require it.' >&2; exit 2; fi
  if [[ -n ${WEBKIT_DISABLE_SANDBOX_THIS_IS_DANGEROUS:-} ]]; then echo 'Refusing disabled browser sandbox.' >&2; exit 2; fi
fi
exec "$application/QuickExcel.Avalonia" "$@"
