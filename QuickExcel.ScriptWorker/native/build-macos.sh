#!/usr/bin/env bash
set -euo pipefail
native=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$native/../.." && pwd)
[[ "$(uname -s):$(uname -m)" == Darwin:arm64 ]] || exit 2
output="$root/artifacts/native/osx-arm64"
mkdir -p "$output"
cc=${CC:-clang}
flags=(-std=c11 -O2 -Wall -Wextra -Werror -Wno-deprecated-declarations -fstack-protector-strong -D_FORTIFY_SOURCE=2 -D_DARWIN_C_SOURCE)
"$cc" "${flags[@]}" -DQX_SANDBOX_LAUNCHER "$native/macos-sandbox.c" -o "$output/quickexcel-sandbox" -Wl,-dead_strip
"$cc" "${flags[@]}" -fPIC -dynamiclib "$native/macos-sandbox.c" -o "$output/libquickexcel-sandbox.dylib" -Wl,-install_name,@loader_path/libquickexcel-sandbox.dylib
codesign --force --sign - "$output/quickexcel-sandbox" "$output/libquickexcel-sandbox.dylib"
{
  printf 'macOS arm64 sandbox policy 1\n'
  "$cc" --version | head -1
  shasum -a 256 "$native/macos-sandbox.c" "$native/build-macos.sh" "$output/quickexcel-sandbox" "$output/libquickexcel-sandbox.dylib"
} > "$output/sandbox-build-identity.txt"
