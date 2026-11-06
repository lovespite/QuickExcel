#!/usr/bin/env python3
"""Build/sign a macOS preview bundle on a Mac; notarize only when explicitly requested.

Multi-file .NET bundles cannot be sealed by codesign because every file under
Contents/MacOS is treated as code; a single-file or Resources layout is required
before this becomes a supported signed .app path. Raw package self-check remains
the local acceptance path until then."""
import argparse,hashlib,json,pathlib,plistlib,platform,shutil,subprocess
parser=argparse.ArgumentParser()
parser.add_argument('package',type=pathlib.Path)
parser.add_argument('output',type=pathlib.Path,help='New .app path under repository artifacts')
parser.add_argument('--identity',default='-',help='Developer ID identity, or ad-hoc local preview')
parser.add_argument('--notary-profile',help='Explicit opt-in: submit signed preview using this keychain profile')
parser.add_argument('--self-check',action='store_true')
args=parser.parse_args()
assert platform.system()=='Darwin' and platform.machine()=='arm64','Run on macOS arm64'
root=pathlib.Path(__file__).resolve().parent.parent;artifacts=root/'artifacts'
package=args.package.resolve(strict=True);bundle=args.output.resolve()
assert artifacts in package.parents and artifacts in bundle.parents and bundle.suffix=='.app' and not bundle.exists()
manifest=json.loads((package/'release.json').read_text(encoding='utf-8-sig'))
assert manifest['rid']=='osx-arm64' and manifest['diagnostics'] is False
assert (package/'ScriptWorker'/'QuickExcel.ScriptWorker').is_file(),'macOS isolated worker is missing from the package'
assert manifest.get('scriptCapability','').startswith('Seatbelt'),'macOS package must record its sandbox capability'
assert not args.notary_profile or args.identity!='-','Notarization requires Developer ID signing'
for record in manifest['files']:
 file=package/record['path']
 assert file.is_file() and not file.is_symlink() and hashlib.sha256(file.read_bytes()).hexdigest().lower()==record['sha256'].lower()
for file in package.rglob('*'):assert not file.is_symlink(),'Unexpected linked package entry'

def run(*command):subprocess.run(command,check=True)

entitlements=bundle.with_suffix('.entitlements.plist')
with entitlements.open('xb') as file:plistlib.dump({'com.apple.security.cs.allow-jit':True},file)
# The main executable must be signed while it is still outside bundle context:
# inside Contents/MacOS codesign validates every managed DLL as nested code.
# Sign a staging copy so the verified source package stays unchanged.
staging=bundle.with_suffix('.staging')
assert not staging.exists()
shutil.copytree(package,staging)
app=staging/'QuickExcel.Avalonia'
with app.open('rb') as stream:magic=stream.read(4)
if magic in [b'\xcf\xfa\xed\xfe',b'\xfe\xed\xfa\xcf',b'\xca\xfe\xba\xbe',b'\xbe\xba\xfe\xca']:
 run('/usr/bin/lipo','-verify_arch','arm64',str(app))
 run('/usr/bin/codesign','--force','--sign',args.identity,*(['--options','runtime']+(['--timestamp'] if args.identity!='-' else [])),'--entitlements',str(entitlements),str(app))
macos=bundle/'Contents'/'MacOS';shutil.copytree(staging,macos);shutil.rmtree(staging)
info=dict(CFBundleIdentifier='org.quickexcel.desktop',CFBundleName='QuickExcel',CFBundleDisplayName='QuickExcel',CFBundleExecutable='QuickExcel.Avalonia',CFBundlePackageType='APPL',CFBundleShortVersionString=manifest['version'],CFBundleVersion=manifest['version'],LSMinimumSystemVersion='13.0',NSHighResolutionCapable=True,NSPrincipalClass='NSApplication')
with (bundle/'Contents'/'Info.plist').open('wb') as file:plistlib.dump(info,file)
(macos/'QuickExcel.Avalonia').chmod(0o755)
(macos/'metadata').mkdir(exist_ok=True)
(macos/'release.json').rename(macos/'metadata'/'unsigned-release.json')

common=['--force','--sign',args.identity,'--options','runtime']
if args.identity!='-':common+=['--timestamp']
# Sign every nested Mach-O inside out; the bundle is sealed last.
for file in sorted(macos.rglob('*'),key=lambda p:len(p.parts),reverse=True):
 if not file.is_file():continue
 # The bundle's main executable is signed with the bundle below; signing it
 # early makes codesign validate every managed DLL as an unsigned subcomponent.
 if file.name=='QuickExcel.Avalonia':continue
 with file.open('rb') as stream:magic=stream.read(4)
 is_macho=magic in [b'\xcf\xfa\xed\xfe',b'\xfe\xed\xfa\xcf',b'\xca\xfe\xba\xbe',b'\xbe\xba\xfe\xca']
 # Managed assemblies are PE files, but codesign still treats them as nested
 # code inside Contents/MacOS and refuses to seal an unsigned bundle.
 if not is_macho and file.suffix!='.dll':continue
 if is_macho:run('/usr/bin/lipo','-verify_arch','arm64',str(file))
 extra=['--entitlements',str(entitlements)] if file.name=='QuickExcel.ScriptWorker' else []
 run('/usr/bin/codesign',*common,*extra,str(file))
run('/usr/bin/codesign',*common,'--entitlements',str(entitlements),str(bundle))
run('/usr/bin/codesign','--verify','--strict','--verbose=2',str(bundle))
if args.self_check:run(str(macos/'QuickExcel.Avalonia'),'--self-check',str(bundle.with_suffix('.self-check')))
if args.notary_profile:
 submission=bundle.with_suffix('.submission.zip')
 assert not submission.exists()
 run('/usr/bin/ditto','-c','-k','--keepParent',str(bundle),str(submission))
 with bundle.with_suffix('.notary.json').open('xb') as log:
  subprocess.run(['/usr/bin/xcrun','notarytool','submit',str(submission),'--keychain-profile',args.notary_profile,'--wait','--output-format','json'],check=True,stdout=log)
 result=json.loads(bundle.with_suffix('.notary.json').read_text());assert result['status']=='Accepted',result
 run('/usr/bin/xcrun','stapler','staple',str(bundle));run('/usr/bin/xcrun','stapler','validate',str(bundle))
 run('/usr/sbin/spctl','--assess','--type','execute','--verbose=2',str(bundle))
files=[dict(path=p.relative_to(bundle).as_posix(),sha256=hashlib.sha256(p.read_bytes()).hexdigest()) for p in sorted(bundle.rglob('*')) if p.is_file()]
bundle.with_suffix('.manifest.json').write_text(json.dumps(dict(schemaVersion=1,preview=True,scriptsEnabled=True,notarized=bool(args.notary_profile),files=files),indent=2))
print(f'Prepared {bundle}; macOS platform acceptance and sandbox qualification remain separate.')
