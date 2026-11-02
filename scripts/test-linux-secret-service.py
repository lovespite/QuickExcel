import argparse,os,pathlib,secrets,subprocess,time
parser=argparse.ArgumentParser(description="Run inside dbus-run-session, with a new isolated keyring")
parser.add_argument("executable",type=pathlib.Path);parser.add_argument("output",type=pathlib.Path);parser.add_argument("--desktop",action="store_true");args=parser.parse_args()
assert os.geteuid()!=0 and args.output.is_absolute() and not args.output.exists()
assert args.executable.is_absolute() and args.executable.is_file()
root=args.output
root.mkdir(mode=0o700)
(root/'data').mkdir(mode=0o700)
(root/'control').mkdir(mode=0o700)
env=os.environ.copy();env['XDG_DATA_HOME']=str(root/'data')
with open(root/'startup.log','wb') as log:
 daemon=subprocess.Popen(['/usr/bin/gnome-keyring-daemon','--foreground','--unlock','--components=secrets','--control-directory',str(root/'control')],stdin=subprocess.PIPE,stdout=log,stderr=log,env=env)
 try:
  daemon.stdin.write(secrets.token_bytes(32).hex().encode());daemon.stdin.close()
  time.sleep(.3)
  ready=subprocess.run(['busctl','--address='+env['DBUS_SESSION_BUS_ADDRESS'],'get-property','org.freedesktop.secrets','/org/freedesktop/secrets/aliases/default','org.freedesktop.Secret.Collection','Locked'],capture_output=True,env=env,timeout=2)
  assert ready.returncode==0 and ready.stdout.strip()==b'b false',ready.stderr.decode()
  cases=[('--workspace-smoke','workspace'),('--grid-smoke','grid'),('--web-smoke','web')] if args.desktop else [('--trust','trust'),('--trust-unavailable','unavailable')]
  for mode,name in cases:
   with open(root/(name+'.log'),'wb') as output:
    subprocess.run([str(args.executable),mode,str(root/name)],env=env,stdout=output,stderr=output,check=True,timeout=180 if args.desktop else 20)
   print('PASS: '+name)
 finally:
  daemon.terminate();daemon.wait(timeout=3)
