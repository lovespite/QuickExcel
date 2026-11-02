#!/usr/bin/env python3
"""Ordinary-user Linux package checks, inside an existing graphical session."""
import argparse,hashlib,json,os,pathlib,shutil,subprocess,time
parser=argparse.ArgumentParser();parser.add_argument('package',type=pathlib.Path);parser.add_argument('output',type=pathlib.Path);args=parser.parse_args()
assert os.geteuid()!=0 and (os.getenv('DISPLAY') or os.getenv('WAYLAND_DISPLAY'))
assert not os.getenv('WEBKIT_DISABLE_SANDBOX_THIS_IS_DANGEROUS')
source=args.package.resolve(strict=True);root=args.output.resolve();assert not root.exists();root.mkdir(mode=0o700)
manifest=json.loads((source/'release.json').read_text(encoding='utf-8-sig'));assert manifest['rid']=='linux-x64' and manifest['diagnostics'] is False
for record in manifest['files']:
 path=source/record['path'];assert not path.is_symlink() and hashlib.sha256(path.read_bytes()).hexdigest().lower()==record['sha256'].lower()
checks=[]
def check(package,name,expected):
 with (root/(name+'.log')).open('wb') as log:
  result=subprocess.run([str(package/'Start-QuickExcel.sh'),'--self-check',str(root/name)],stdout=log,stderr=log,timeout=90)
 data=json.loads((root/name/'result.json').read_text());assert data['passed'] is expected and result.returncode==(0 if expected else 1),data
 checks.append(dict(name=name,passed=True,exitCode=result.returncode,detail=data))
check(source,'original',True)
assert subprocess.run([str(source/'QuickExcel.Avalonia'),'--web-smoke',str(root/'forbidden')],timeout=5).returncode==2
checks.append(dict(name='diagnostic-flag-rejected',passed=True))
copy=root/'只读 安装包';shutil.copytree(source,copy)
for path in copy.rglob('*'):
 assert not path.is_symlink()
 path.chmod(0o555 if path.is_dir() or os.access(path,os.X_OK) else 0o444)
copy.chmod(0o555)
check(copy,'readonly-unicode-spaces',True)
# Mutate only this exact owned fixture. Keep all per-case logs and evidence.
web=copy/'Web';web.chmod(0o755);asset=web/'script.html';asset.unlink()
check(copy,'missing-web',False)
shutil.copy2(source/'Web'/'script.html',asset)
worker=copy/'ScriptWorker';worker.chmod(0o755);(worker/'QuickExcel.ScriptWorker').unlink()
check(copy,'missing-worker',False)
(root/'result.json').write_text(json.dumps(dict(passed=True,checks=checks,source=str(source)),indent=2))
print(f'PASS: {len(checks)} Linux package checks')
