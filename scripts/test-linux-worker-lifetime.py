#!/usr/bin/env python3
"""Ordinary-user destructive-process test; all filesystem mutations are owned test paths."""
import json, os, pathlib, subprocess, sys, time, uuid

tool = pathlib.Path(sys.argv[1]).resolve()
evidence = pathlib.Path(sys.argv[2]).resolve()
evidence.mkdir()  # Never reuse another run.
out = open(evidence / "host.stdout.log", "w")
err = open(evidence / "host.stderr.log", "w")
process = subprocess.Popen([str(tool), "--hold", str(evidence / "parent")], stdout=out, stderr=err)
try:
    ready = evidence / "parent" / "live.json"
    deadline = time.monotonic() + 8
    while not ready.exists() and process.poll() is None and time.monotonic() < deadline:
        time.sleep(.02)
    identity = json.loads(ready.read_text())
    assert identity["host"] == process.pid
    record = pathlib.Path(identity["payload"]).parent
    root = record.parent
    assert record.name == identity["scope"][len("quickexcel-worker-"):-len(".scope")]
    assert root.name == "QuickExcel.LinuxWorkers.v1" and not record.is_symlink()
    unowned = root / uuid.uuid4().hex
    unowned.mkdir(mode=0o700)
    marker = unowned / "keep.txt"
    marker.write_text("Unjournaled directories are not owned by worker recovery.")
    subprocess.run([str(tool), "--recover", str(evidence / "while-live.json")], check=True, timeout=3)
    assert json.loads((evidence / "while-live.json").read_text())["recovered"] == 0
    assert record.is_dir() and marker.is_file()
    began = time.monotonic()
    process.kill()  # The test kills only its own host; never directly kills its worker.
    process.wait(timeout=2)
    state = "alive"
    while time.monotonic() - began < 1.5:
        try:
            stat = pathlib.Path(f'/proc/{identity["worker"]}/stat').read_text()
            parts = stat[stat.rfind(")")+2:].split()
            if int(parts[19]) != identity["start"]: state = "pid-reused"; break
            if parts[0] == "Z": state = "terminated-zombie"; break
        except FileNotFoundError:
            state = "gone"; break
        time.sleep(.01)
    elapsed = (time.monotonic() - began) * 1000
    assert state != "alive", "Worker survived its host."
    # PID 1 may briefly own a terminated zombie; recovery must wait for reaping.
    deadline = time.monotonic() + 2
    while pathlib.Path(f'/proc/{identity["worker"]}').exists() and time.monotonic() < deadline: time.sleep(.02)
    subprocess.run([str(tool), "--recover", str(evidence / "after-death.json")], check=True, timeout=3)
    assert json.loads((evidence / "after-death.json").read_text())["recovered"] == 1
    assert not record.exists() and marker.is_file()
    marker.unlink(); unowned.rmdir()  # Only the marker created above, never a prefix sweep.
    (evidence / "result.json").write_text(json.dumps(dict(passed=True, identity=identity, workerState=state, terminationMs=elapsed, liveLeasePreserved=True, abandonedRoundRecovered=True, unjournaledDirectoryPreserved=True), indent=2))
    print("PASS: host death, live lease and journaled recovery")
finally:
    if process.poll() is None: process.kill(); process.wait(timeout=2)
    out.close(); err.close()
