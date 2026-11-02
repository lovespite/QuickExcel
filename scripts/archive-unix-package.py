#!/usr/bin/env python3
"""Archive a fresh package with portable modes and verify the serialized manifest."""
import hashlib,json,pathlib,sys,tarfile
package=pathlib.Path(sys.argv[1]).resolve();archive=pathlib.Path(sys.argv[2]).resolve()
assert package.is_dir() and not archive.exists()
manifest=json.loads((package/'release.json').read_text(encoding='utf-8-sig'))
executables={'QuickExcel.Avalonia','ScriptWorker/QuickExcel.ScriptWorker','ScriptWorker/quickexcel-sandbox','Start-QuickExcel.sh'}
with tarfile.open(archive,'x:gz') as output:
 for file in sorted(package.rglob('*')):
  assert not file.is_symlink()
  name=file.relative_to(package).as_posix()
  info=output.gettarinfo(str(file),arcname='QuickExcel/'+name)
  info.uid=info.gid=0;info.uname=info.gname='';info.mode=0o755 if file.is_dir() or name in executables else 0o644
  if file.is_file():
   with file.open('rb') as stream:output.addfile(info,stream)
  else:output.addfile(info)
with tarfile.open(archive,'r:gz') as source:
 for file in manifest['files']:
  data=source.extractfile('QuickExcel/'+file['path']).read()
  assert len(data)==file['size'] and hashlib.sha256(data).hexdigest().upper()==file['sha256'].upper()
archive.with_suffix(archive.suffix+'.sha256').write_text(hashlib.sha256(archive.read_bytes()).hexdigest()+'  '+archive.name+'\n')
print(archive)
