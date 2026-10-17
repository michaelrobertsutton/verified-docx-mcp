"""Sample heartbeat/hash diagnostics for disposable #101/#102 panes only."""

import json
import sys
import time
from pathlib import Path

from verified_docx_mcp.live import bridge

registry = bridge.start_in_background()


def panes():
    return [s for s in registry.list() if s.document_name.startswith("word-101-")]


label = sys.argv[1]
duration = float(sys.argv[2])
started = time.monotonic()
before = {s.document_name: (s.heartbeat_count, s.hash_sample_count) for s in panes()}
max_age = {}
while time.monotonic() - started < duration:
    for s in panes():
        age = s.heartbeat_age()
        max_age[s.document_name] = max(max_age.get(s.document_name, 0), age)
    time.sleep(0.2)
after = {s.document_name: (s.heartbeat_count, s.hash_sample_count) for s in panes()}
result = {
    "label": label,
    "seconds": time.monotonic() - started,
    "panes": {
        n: {
            "heartbeats": v[0] - before.get(n, (0, 0))[0],
            "hash_requests": v[1] - before.get(n, (0, 0))[1],
            "native_hash_syncs": v[1] - before.get(n, (0, 0))[1],
            "max_heartbeat_age_s": max_age.get(n),
        }
        for n, v in after.items()
    },
}
Path("/private/tmp/word-101-" + label + "-heartbeat.json").write_text(json.dumps(result, indent=2))
print(json.dumps(result), flush=True)
