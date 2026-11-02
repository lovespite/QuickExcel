import argparse,hashlib,json,os,pathlib,shutil,socket,subprocess,time,xml.etree.ElementTree as ET
parser=argparse.ArgumentParser(description="Isolated non-root CUPS/IPP PDF handoff; never uses system queues")
parser.add_argument("executable",type=pathlib.Path);parser.add_argument("output",type=pathlib.Path);args=parser.parse_args()
assert os.geteuid()!=0,"Run as an ordinary user"
assert args.executable.is_absolute() and args.executable.is_file()
assert args.output.is_absolute() and not args.output.exists()
root=args.output;root.mkdir(mode=0o700)
for name in ['config','state','spool','cache','tmp','device']:(root/name).mkdir(mode=0o700)
(root/'cups-files.conf').write_text(f'User {__import__("pwd").getpwuid(os.getuid()).pw_name}\nGroup {__import__("grp").getgrgid(os.getgid()).gr_name}\nSystemGroup lpadmin\n'+''.join(f'{key} {root/value}\n' for key,value in [('ServerRoot','config'),('StateDir','state'),('RequestRoot','spool'),('CacheDir','cache'),('TempDir','tmp'),('AccessLog','access.log'),('ErrorLog','error.log'),('PageLog','page.log')])+'DataDir /usr/share/cups\nServerBin /usr/lib/cups\nDocumentRoot /usr/share/cups/doc-root\n')
(root/'cupsd.conf').write_text('Listen 127.0.0.1:8631\nServerName localhost\nServerAlias localhost\nBrowsing Off\nWebInterface No\nLogLevel debug\n<Location />\nOrder allow,deny\nAllow localhost\n</Location>\n<Location /admin>\nOrder allow,deny\nAllow localhost\n</Location>\n<Policy default>\n<Limit All>\nOrder allow,deny\nAllow localhost\n</Limit>\n</Policy>\n')
# The system cupsd AppArmor profile intentionally refuses private home paths.
# Run an owned copy for this fixture only, with no root or real queues.
shutil.copy2("/usr/sbin/cupsd",root/"cupsd-fixture")
processes=[];logs=[]
def start(command,name):
 log=open(root/(name+'.log'),'wb');logs.append(log)
 p=subprocess.Popen(command,stdout=log,stderr=log,cwd=root);processes.append(p);return p
def ready(port,p):
 for _ in range(100):
  assert p.poll() is None, 'service exited: '+str(p.args)
  try:
   with socket.create_connection(('127.0.0.1',port),timeout=.1):return
  except OSError:time.sleep(.05)
 raise RuntimeError('service did not start')
try:
 for port in [8631,8632]:
  with socket.socket() as check:check.bind(('127.0.0.1',port))
 printer=start(['/usr/sbin/ippeveprinter' if pathlib.Path('/usr/sbin/ippeveprinter').exists() else '/usr/bin/ippeveprinter','-p','8632','-n','localhost','-d',str(root/'device'),'-k','-f','application/pdf','-F','application/pdf','QuickExcel-A6'],'device');ready(8632,printer)
 cups=start([str(root/'cupsd-fixture'),'-f','-c',str(root/'cupsd.conf'),'-s',str(root/'cups-files.conf')],'daemon');ready(8631,cups)
 subprocess.run(['/usr/sbin/lpadmin','-h','localhost:8631','-p','QuickExcel-A6','-E','-v','ipp://localhost:8632/ipp/print','-m','everywhere'],check=True,timeout=10)
 env=os.environ.copy();env['CUPS_SERVER']='localhost:8631'
 subprocess.run([str(args.executable),'--cups',str(root/'result')],check=True,env=env,timeout=30)
 files=[]
 for _ in range(200):
  files=[p for p in (root/'device').iterdir() if p.is_file() and p.read_bytes().startswith(b'%PDF')]
  if files:break
  time.sleep(.05)
 assert len(files)==1,'No unique PDF received by virtual printer'
 expected_file=root/'result'/'expected.pdf'
 expected=hashlib.sha256(expected_file.read_bytes()).hexdigest();actual=hashlib.sha256(files[0].read_bytes()).hexdigest()
 spool=[p for p in (root/'spool').glob('d*') if p.is_file() and p.read_bytes().startswith(b'%PDF')]
 assert len(spool)==1 and hashlib.sha256(spool[0].read_bytes()).hexdigest()==expected,'Scheduler must receive exact generated PDF'
 # cups-filters normalizes the A4 box from rounded application DIPs to exact mm.
 # Independently check page geometry and all text positions with Poppler.
 def geometry(path):
  xml=subprocess.run(['/usr/bin/pdftotext','-bbox',str(path),'-'],capture_output=True,check=True,timeout=5).stdout
  document=ET.fromstring(xml);ns={'x':'http://www.w3.org/1999/xhtml'}
  pages=document.findall('.//x:page',ns);words=document.findall('.//x:word',ns)
  return pages,words
 before,words_before=geometry(expected_file);after,words_after=geometry(files[0])
 assert len(before)==len(after)==1 and len(words_before)==len(words_after)>0
 assert any('中文' in (w.text or '') for w in words_after)
 deltas=[]
 for a,b in zip(before+words_before,after+words_after):
  assert a.text==b.text
  for key,value in a.attrib.items():deltas.append(abs(float(value)-float(b.attrib[key])))
 assert max(deltas)<.1,'Printer normalized page/text positions by more than 0.1 point'
 (root/'handoff.json').write_text(json.dumps(dict(passed=True,schedulerExact=True,expected=expected,received=actual,printerByteIdentical=expected==actual,maxGeometryDeltaPoints=max(deltas),file=str(files[0])),indent=2))
 print('PASS: exact scheduler PDF; virtual printer page and Chinese text geometry within 0.1 point')
finally:
 for p in reversed(processes):
  if p.poll() is None:
   p.terminate()
   try:p.wait(timeout=3)
   except subprocess.TimeoutExpired:p.kill();p.wait()
 for log in logs:log.close()
