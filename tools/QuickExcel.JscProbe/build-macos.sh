#!/usr/bin/env bash
set -euo pipefail
here=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$here/../.." && pwd)
[[ "$(uname -s):$(uname -m)" == Darwin:arm64 ]] || exit 2
cd "$here"
python3 - <<'PY'
from prepare import ROOT, PINS, download, archive_source
path = ROOT / 'source/webkitgtk-2.54.2/Source/cmake/WebKitFS.cmake'
# Older caches may predate the Apple framework-header adjustment.
if path.exists() and 'JavaScriptCore_PRIVATE_FRAMEWORK_HEADERS_DIR' not in path.read_text(encoding='utf-8'):
    path.unlink()
path = ROOT / 'source/webkitgtk-2.54.2/Source/WTF/wtf/PlatformJSCOnly.cmake'
if path.exists() and 'cocoa/TimeZoneCocoa.cpp' in path.read_text(encoding='utf-8'):
    path.unlink()
for name in ('webkit', 'icu'):
    archive_source(download(PINS[name]), ROOT / 'source', webkit=name == 'webkit')
archive_source(download(PINS['typescript']), ROOT / 'typescript')
for source in (ROOT / 'source/icu').rglob('*'):
    if not source.is_file():
        continue
    target = ROOT / 'source/icu-macos' / source.relative_to(ROOT / 'source/icu')
    data = source.read_bytes()
    if b'\0' not in data:
        data = data.replace(b'\r\n', b'\n')
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists() or target.read_bytes() != data:
        target.write_bytes(data)
    if data.startswith(b'#!'):
        target.chmod(0o755)
PY
python3 "$here/prepare-macos.py"
python3 "$here/prepare_compiler.py"
# The GTK release archive does not preserve executable bits; the Darwin build
# invokes this wrapper directly.
chmod +x "$root/artifacts/jsc/source/webkitgtk-2.54.2/Source/cmake/clang-wrapper"
out="$root/artifacts/jsc"
run="$out/logs/osx-arm64-$(date -u +%Y%m%d-%H%M%S)"
mkdir -p "$run" "$out/build/icu-macos" "$out/build/osx-arm64"
export CC=${CC:-clang} CXX=${CXX:-clang++}
cd "$out/build/icu-macos"
if [[ ! -f Makefile ]]; then
  bash "$out/source/icu-macos/source/configure" --prefix="$out/icu-macos" --disable-tests --disable-samples > "$run/icu-configure.log" 2>&1
fi
make -j "${QX_BUILD_PARALLEL:-4}" > "$run/icu-build.log" 2>&1
make install > "$run/icu-install.log" 2>&1
cmake -S "$out/source/webkitgtk-2.54.2" -B "$out/build/osx-arm64" -G Ninja \
  -DCMAKE_BUILD_TYPE=Release -DPORT=JSCOnly -DENABLE_FTL_JIT=OFF -DENABLE_WEBASSEMBLY=OFF \
  -DUSE_SYSTEM_MALLOC=OFF -DUSE_MIMALLOC=ON -DUSE_HEADER_MAPS=OFF -DENABLE_REMOTE_INSPECTOR=OFF -DENABLE_API_TESTS=OFF -DDEVELOPER_MODE=OFF \
  -DICU_ROOT="$out/icu-macos" -DCMAKE_OSX_ARCHITECTURES=arm64 -DCMAKE_OSX_DEPLOYMENT_TARGET=13.0 \
  -DCMAKE_BUILD_WITH_INSTALL_RPATH=ON -DCMAKE_INSTALL_RPATH='@loader_path' \
  -DRuby_EXECUTABLE=/usr/bin/ruby -DRUBY_EXECUTABLE=/usr/bin/ruby \
  -DQX_JSC_PROBE_SOURCE="$here" -DCMAKE_PROJECT_WebKit_INCLUDE="$here/cmake-hook.cmake" > "$run/configure.log" 2>&1
cmake --build "$out/build/osx-arm64" --target JavaScriptCore --parallel "${QX_BUILD_PARALLEL:-4}" > "$run/build.log" 2>&1
payload="$out/build/osx-arm64/payload"
rm -rf "$payload"
mkdir -p "$payload"
jsc=$(find "$out/build/osx-arm64" -path '*JavaScriptCore.framework*' -type f -name JavaScriptCore | head -1)
[[ -n "$jsc" ]] || jsc="$out/build/osx-arm64/libJavaScriptCore.dylib"
cp -L "$jsc" "$payload/libJavaScriptCore.dylib"
# WebKit's Apple port links the OS ICU (libicucore) with unversioned symbols.
# Resolve it through the dyld shared cache, which the worker profile can read.
install_name_tool -id "@loader_path/libJavaScriptCore.dylib" "$payload/libJavaScriptCore.dylib"
while IFS= read -r dependency; do
  [[ -n "$dependency" ]] || continue
  install_name_tool -change "$dependency" /usr/lib/libicucore.A.dylib "$payload/libJavaScriptCore.dylib"
done < <(otool -L "$payload/libJavaScriptCore.dylib" | awk '/libicucore/{print $1}')
codesign --force --sign - "$payload/libJavaScriptCore.dylib"
cp "$out/compiler/typescript.bundle.json" "$payload/typescript.bundle.json"
{
  printf 'JavaScriptCore/2.54.2/b95c91b1/qx-sync-2 macOS arm64; JSC-MEM-001 open\n'
  clang --version | head -1
  cmake --version | head -1
  ruby --version
  shasum -a 256 "$here/sources.json" "$here/source-adjustments.json" "$here/macos-sources.json" "$here/build-macos.sh" \
    "$root/QuickExcel.Scripting/native/jsc/bridge.cpp" "$payload"/libJavaScriptCore.dylib "$payload/typescript.bundle.json"
  otool -L "$payload/libJavaScriptCore.dylib" | grep libicucore || true
} > "$payload/identity.txt"
echo "Built JSC payload $payload; logs $run"
