#!/usr/bin/env bash
set -euo pipefail
here=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$here/../.." && pwd)
cd "$here"
python3 - <<'PY'
from prepare import ROOT, PINS, download, archive_source
import os, tarfile
archive = download(PINS['ruby_source'])
archive_source(archive, ROOT / 'source')
with tarfile.open(archive) as source:
    for entry in source:
        if entry.isfile():
            os.utime(ROOT / 'source' / entry.name, (entry.mtime, entry.mtime))
for path in (ROOT / 'source/ruby-4.0.9').rglob('*'):
    if path.is_file() and path.read_bytes().startswith(b'#!'):
        path.chmod(0o755)
PY
out="$root/artifacts/jsc"
run="$out/logs/ruby-linux-$(date -u +%Y%m%d-%H%M%S)"
mkdir -p "$run" "$out/build/ruby-linux"
cd "$out/build/ruby-linux"
if [[ ! -f Makefile ]]; then
  CC=clang-18 bash "$out/source/ruby-4.0.9/configure" --prefix="$out/tools/ruby-linux" --disable-install-doc --disable-yjit --disable-zjit --with-out-ext=openssl > "$run/configure.log" 2>&1
fi
# Bootstrap with the freshly built Ruby, avoiding the old distribution Ruby's
# object-ID/GC failure. The release archive contains the generated C sources.
make -j "${QX_BUILD_PARALLEL:-2}" miniruby > "$run/miniruby.log" 2>&1
base="./miniruby -I. -I$out/source/ruby-4.0.9/lib --disable=gems"
make -j "${QX_BUILD_PARALLEL:-2}" ruby BASERUBY="$base" > "$run/ruby.log" 2>&1
base="./ruby -I. -I.ext/x86_64-linux -I$out/source/ruby-4.0.9/lib --disable=gems"
make -j "${QX_BUILD_PARALLEL:-2}" main BASERUBY="$base" > "$run/build.log" 2>&1
make install BASERUBY="$base" > "$run/install.log" 2>&1
"$out/tools/ruby-linux/bin/ruby" --version
