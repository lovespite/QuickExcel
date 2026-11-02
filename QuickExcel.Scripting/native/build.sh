#!/usr/bin/env bash
set -euo pipefail
native=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$native/../.." && pwd)
case "$(uname -s):$(uname -m)" in
  Linux:x86_64) rid=linux-x64; library=libquickexcel-js.so ;;
  Darwin:arm64) rid=osx-arm64; library=libquickexcel-js.dylib ;;
  *) echo 'Supported native builds: Linux x64 and macOS arm64.' >&2; exit 2 ;;
esac
version=0.17.0
sha=9de1bf573606d2aebbbab6217e616a66e09e3613518489b37b9b3157cf398dcc
output="$root/artifacts/native/$rid"
mkdir -p "$output"
archive="$root/artifacts/native/quickjs-ng-v$version.zip"
if [[ ! -f "$archive" ]]; then
  curl --fail --location --proto '=https' --tlsv1.2 "https://github.com/quickjs-ng/quickjs/archive/refs/tags/v$version.zip" --output "$archive.download"
  mv "$archive.download" "$archive"
fi
if command -v sha256sum >/dev/null; then actual=$(sha256sum "$archive" | cut -d' ' -f1); else actual=$(shasum -a 256 "$archive" | cut -d' ' -f1); fi
[[ "$actual" == "$sha" ]] || { echo 'QuickJS source hash mismatch.' >&2; exit 1; }
# Extract into a new owned directory: modified or stale vendor trees are never reused.
run=$(mktemp -d "$output/build-XXXXXXXX")
unzip -q "$archive" -d "$run/source"
cmake -S "$native" -B "$run/cmake" -DCMAKE_BUILD_TYPE=Release -DQX_QUICKJS_SOURCE="$run/source/quickjs-$version"
cmake --build "$run/cmake" --target quickexcel-js --parallel 4
cp "$run/cmake/$library" "$output/$library.new"
mv "$output/$library.new" "$output/$library"
printf 'QuickJS %s\nSource SHA-256 %s\nRID %s\nBuild %s\n' "$version" "$sha" "$rid" "$run" > "$output/build-identity.txt"
echo "Built $output/$library"
{
  "${CC:-cc}" --version | head -1
  if command -v sha256sum >/dev/null; then sha256sum "$native/quickexcel-js.c" "$native/CMakeLists.txt" "$output/$library"; else shasum -a 256 "$native/quickexcel-js.c" "$native/CMakeLists.txt" "$output/$library"; fi
} >> "$output/build-identity.txt"
