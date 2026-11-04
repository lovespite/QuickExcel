#!/usr/bin/env bash
set -euo pipefail
here=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$here/../.." && pwd)
[[ "$(uname -s):$(uname -m)" == Linux:x86_64 ]] || exit 2
cd "$here"
python3 - <<'PY'
from prepare import ROOT, PINS, download, archive_source
for name in ('webkit', 'icu'):
    archive_source(download(PINS[name]), ROOT / 'source', webkit=name == 'webkit')
archive_source(download(PINS['typescript']), ROOT / 'typescript')
# The pinned ICU ZIP carries CRLF shell/configure inputs. Keep its verified
# extraction intact and build a separate, reproducible LF-normalized copy.
for source in (ROOT / 'source/icu').rglob('*'):
    if not source.is_file():
        continue
    target = ROOT / 'source/icu-linux' / source.relative_to(ROOT / 'source/icu')
    data = source.read_bytes()
    if b'\0' not in data:
        data = data.replace(b'\r\n', b'\n')
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists() or target.read_bytes() != data:
        target.write_bytes(data)
    if data.startswith(b'#!'):
        target.chmod(0o755)
PY
python3 "$here/prepare_compiler.py"
out="$root/artifacts/jsc"
run="$out/logs/linux-$(date -u +%Y%m%d-%H%M%S)"
mkdir -p "$run" "$out/build/icu-linux" "$out/build/linux-x64"
export CC=${CC:-clang-18} CXX=${CXX:-clang++-18}
if [[ ! -x "$out/tools/ruby-linux/bin/ruby" ]]; then bash "$here/prepare-ruby-linux.sh"; fi
export PATH="$out/tools/ruby-linux/bin:$PATH"
cd "$out/build/icu-linux"
if [[ ! -f Makefile ]]; then
  bash "$out/source/icu-linux/source/configure" --prefix="$out/icu-linux" --disable-tests --disable-samples > "$run/icu-configure.log" 2>&1
fi
make -j "${QX_BUILD_PARALLEL:-4}" > "$run/icu-build.log" 2>&1
if [[ ! -f "$out/icu-linux/lib/libicuuc.so.77.1" || "$out/build/icu-linux/lib/libicuuc.so.77.1" -nt "$out/icu-linux/lib/libicuuc.so.77.1" ]]; then
  make install > "$run/icu-install.log" 2>&1
fi
export PKG_CONFIG_PATH="$out/icu-linux/lib/pkgconfig"
cmake -S "$out/source/webkitgtk-2.54.2" -B "$out/build/linux-x64" -G Ninja \
  -DCMAKE_BUILD_TYPE=Release -DPORT=JSCOnly -DENABLE_FTL_JIT=OFF -DENABLE_WEBASSEMBLY=OFF \
  -DUSE_SYSTEM_MALLOC=OFF -DUSE_MIMALLOC=ON -DUSE_HEADER_MAPS=OFF -DENABLE_REMOTE_INSPECTOR=OFF -DENABLE_API_TESTS=OFF -DDEVELOPER_MODE=OFF \
  -DICU_ROOT="$out/icu-linux" -DCMAKE_BUILD_WITH_INSTALL_RPATH=ON -DCMAKE_INSTALL_RPATH='$ORIGIN' -DCMAKE_SHARED_LINKER_FLAGS=-Wl,--disable-new-dtags \
  -DRuby_EXECUTABLE="$out/tools/ruby-linux/bin/ruby" -DRUBY_EXECUTABLE="$out/tools/ruby-linux/bin/ruby" \
  -DQX_JSC_PROBE_SOURCE="$here" -DCMAKE_PROJECT_WebKit_INCLUDE="$here/cmake-hook.cmake" > "$run/configure.log" 2>&1
cmake --build "$out/build/linux-x64" --target JavaScriptCore --parallel "${QX_BUILD_PARALLEL:-4}" > "$run/build.log" 2>&1
payload="$out/build/linux-x64/payload"
mkdir -p "$payload"
cp -L "$out/build/linux-x64/lib/libJavaScriptCore.so" "$payload/libJavaScriptCore.so"
for name in libicudata.so.77 libicuuc.so.77 libicui18n.so.77; do cp -L "$out/icu-linux/lib/$name" "$payload/$name"; done
{
  printf 'JavaScriptCore/2.54.2/b95c91b1/qx-sync-2 Linux x64; JSC-MEM-001 open\n'
  "$CXX" --version | head -1
  cmake --version | head -1
  ruby --version
  sha256sum "$here/sources.json" "$here/source-adjustments.json" "$here/build-linux.sh" "$root/QuickExcel.Scripting/native/jsc/bridge.cpp" "$payload/"*
} > "$out/build/linux-x64/identity.txt"
echo "Built JSC payload $payload; logs $run"
