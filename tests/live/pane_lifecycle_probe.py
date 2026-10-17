"""Suspend only Word WebKit processes, with disposable Word documents exclusively open.

Run after overhead_probe.py, never during a mutation. Every stopped process is
resumed in finally. Refuses to run if any open document isn't a #101 fixture.
"""

import json
import os
import signal
import subprocess
import time
from pathlib import Path

from verified_docx_mcp.live import bridge
from verified_docx_mcp.live.open_pane_live import _open_documents

names = _open_documents()
assert names and all(n.startswith("word-101-") for n in names), names
registry = bridge.start_in_background()
initial = {s.document_name: (s.connected_since, s.heartbeat_age()) for s in registry.list()}
assert initial and set(initial) <= set(names)
rows = subprocess.check_output(["ps", "-axo", "pid,command"], text=True).splitlines()
pids = []
for row in rows:
    if "/WebKit.framework/" not in row or "/com.apple.WebKit.WebContent" not in row:
        continue
    pid = int(row.strip().split()[0])
    files = subprocess.run(["lsof", "-p", str(pid), "-Fn"], capture_output=True, text=True, check=False)
    if "/Containers/com.microsoft.Word/" in files.stdout:
        pids.append(pid)
assert pids, "No Word WebKit process found"
initial = {s.document_name: (s.connected_since, s.heartbeat_age()) for s in registry.list()}
stopped = []
missing = {}
started = time.monotonic()
try:
    for pid in pids:
        os.kill(pid, signal.SIGSTOP)
        stopped.append(pid)
    while time.monotonic() - started < 18:
        present = {s.document_name for s in registry.list()}
        for name in initial:
            if name not in present:
                missing.setdefault(name, time.monotonic() - started)
        time.sleep(0.2)
finally:
    for pid in stopped:
        os.kill(pid, signal.SIGCONT)
resumed = time.monotonic()
while time.monotonic() - resumed < 30:
    current = {s.document_name: s.connected_since for s in registry.list()}
    if all(n in current and current[n] != value[0] for n, value in initial.items()):
        break
    time.sleep(0.2)
assert set(missing) == set(initial), missing
assert all(n in current and current[n] != value[0] for n, value in initial.items()), current
result = {
    "suspended_seconds": resumed - started,
    "reconnect_seconds": time.monotonic() - resumed,
    "detected_age_s": {n: initial[n][1] + elapsed for n, elapsed in missing.items()},
    "fresh_connections": True,
    "process_count": len(stopped),
}
Path("/private/tmp/word-101-lifecycle-evidence.json").write_text(json.dumps(result, indent=2))
print(json.dumps(result), flush=True)
