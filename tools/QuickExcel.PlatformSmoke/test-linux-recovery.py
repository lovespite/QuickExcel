"""Kill only this test's owned hosts; verify both isolated roles and journal recovery."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid

parser = argparse.ArgumentParser()
parser.add_argument('executable', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
root = Path(__file__).resolve().parents[2]
output = args.output.resolve()
if sys.platform != 'linux' or os.geteuid() == 0 or not output.is_relative_to(root / 'artifacts'):
    raise SystemExit('Run as a normal Linux user with evidence under this repository artifacts.')
output.mkdir(parents=True, exist_ok=False)
records = []
failure = None
try:
    for role, option in [('execute', '--hold'), ('compile', '--hold-compile')]:
        case = output / role
        with (output / (role + '.log')).open('w') as log:
            host = subprocess.Popen([str(args.executable.resolve()), option, str(case)], stdout=log, stderr=log)
            try:
                deadline = time.monotonic() + 4
                while not (case / 'live.json').exists():
                    if host.poll() is not None or time.monotonic() >= deadline:
                        raise AssertionError('Owned host did not reach the protected live barrier')
                    time.sleep(0.02)
                live = json.loads((case / 'live.json').read_text())
                assert live['host'] == host.pid
                started = time.monotonic()
                host.kill()
                host.wait(timeout=2)
                stat = Path('/proc') / str(live['worker']) / 'stat'
                while stat.exists():
                    try:
                        fields = stat.read_text().split(') ', 1)[1].split()
                        if int(fields[19]) != live['start']:
                            break
                    except FileNotFoundError:
                        break
                    if time.monotonic() - started >= 2:
                        raise AssertionError('Worker survived its parent beyond two seconds')
                    time.sleep(0.02)
                death_ms = (time.monotonic() - started) * 1000
                payload = Path(live['payload'])
                unowned = payload.parent.parent / uuid.uuid4().hex
                unowned.mkdir()
                (unowned / 'sentinel').write_text('unjournaled')
                recovered = output / (role + '-recovered.json')
                subprocess.run([str(args.executable.resolve()), '--recover', str(recovered)], check=True, timeout=8, stdout=log, stderr=log)
                assert not payload.parent.exists()
                assert (unowned / 'sentinel').read_text() == 'unjournaled'
                (unowned / 'sentinel').unlink()
                unowned.rmdir()
                records.append({'role': role, 'live': live, 'parentDeathMs': death_ms, 'recovery': json.loads(recovered.read_text()), 'unjournaledPreserved': True})
            finally:
                if host.poll() is None:
                    host.kill()
                    host.wait(timeout=2)
except Exception as error:
    failure = repr(error)
(output / 'result.json').write_text(json.dumps({'passed': failure is None, 'records': records, 'error': failure}, indent=2))
print(failure or 'PASS: both Linux roles stop on parent death and recover owned leases')
raise SystemExit(failure is not None)
