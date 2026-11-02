#!/usr/bin/env bash
set -euo pipefail
native=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$native/../.." && pwd)
[[ "$(uname -s):$(uname -m)" == Linux:x86_64 ]] || exit 2
output="$root/artifacts/native/linux-x64"
mkdir -p "$output"
cc=${CC:-cc}
"$cc" -std=c11 -O2 -Wall -Wextra -Werror -fstack-protector-strong -D_FORTIFY_SOURCE=2 -DQX_SANDBOX_LAUNCHER "$native/linux-sandbox.c" -o "$output/quickexcel-sandbox" -Wl,-z,relro,-z,now
"$cc" -std=c11 -O2 -Wall -Wextra -Werror -fstack-protector-strong -D_FORTIFY_SOURCE=2 -fPIC -shared "$native/linux-sandbox.c" -o "$output/libquickexcel-sandbox.so" -Wl,-z,relro,-z,now
{
  printf 'Linux x64 sandbox policy 1\n'
  "$cc" --version | head -1
  sha256sum "$native/linux-sandbox.c" "$native/build-linux.sh" "$output/quickexcel-sandbox" "$output/libquickexcel-sandbox.so"
} > "$output/sandbox-build-identity.txt"
