"""Real-Word acceptance for #110 (live_revision, section digests, ranged and to_file reads).

    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp \
        python tests/live/read_tokens_probe.py /private/tmp/word-110-large.docx

The argument is a DISPOSABLE copy of the large fixture (scripts/make_large_fixture.py);
the probe edits it in Word. This process must own the bridge ports (53135/53136), so no
other bridge owner may be running. Prints one JSON line per check and exits non-zero if
any check fails.
"""

from __future__ import annotations

import hashlib
import json
import statistics
import sys
import tempfile
import time
from pathlib import Path

from verified_docx_mcp import projection, server, text_edit
from verified_docx_mcp.live import bridge, open_pane_live, reads_live

path = str(Path(sys.argv[1]).resolve())
ORIGINAL = Path(tempfile.mkdtemp(prefix="probe-110-")) / "original.docx"
ORIGINAL.write_bytes(Path(path).read_bytes())  # the unedited file, for file-side comparisons
results: list[dict] = []


def check(name: str, ok: bool, **detail) -> None:
    results.append({"check": name, "pass": bool(ok), **detail})
    print(json.dumps(results[-1], default=str), flush=True)


registry = bridge.start_in_background()
status = registry.status() if hasattr(registry, "status") else {}
opened = open_pane_live.execute_open_pane(path, timeout_s=90)
session = registry.get(Path(path).name)
caps = sorted(session.hello.capabilities)
check("pane connected with body_hash_v1", "body_hash_v1" in caps,
      method=opened["method"], capabilities_with_110_112=[c for c in caps if c in ("body_hash_v1", "compact_evidence_v1")])

# 1. live_revision: native path, timing, and token equals a live read's.
times = []
rev = None
for _ in range(5):
    t0 = time.monotonic()
    rev = reads_live.live_revision(path)
    times.append(time.monotonic() - t0)
check("live_revision is native and under 1 s", rev["via"] == "body_hash" and statistics.median(times) < 1.0,
      via=rev["via"], seconds=[round(t, 3) for t in times], paragraphs=300)
live_text = server.execute_read_document(path, source="live", format="text")
check("live_revision equals the live read's revision", live_text["revision"] == rev["revision"],
      revision=rev["revision"][:24] + "...")

# 2. section digests: file projection of the unedited document vs the live pane.
file_digests = [[s["section_key"], s["sha12"]] for s in projection.find_sections_impl(ORIGINAL, digests=True)
                if s["kind"] == "heading"]
pane_digests = rev["sections"]
mismatched = [(a, b) for a, b in zip(file_digests, pane_digests, strict=False) if a != b]
check("file and live sha12 agree on every section",
      file_digests == pane_digests, file_sections=len(file_digests), live_sections=len(pane_digests),
      first_mismatches=mismatched[:3])
compact = server.live_list_sections(path, compact=True)
check("live_list_sections(compact) agrees", [[k, s] for k, _h, s in compact["sections"]] == file_digests,
      sections=len(compact["sections"]))
live_find = server.execute_find_sections(path, source="live", digests=True)
check("find_sections(source=live, digests) agrees",
      [[s["section_key"], s["sha12"]] for s in live_find["sections"] if s["kind"] == "heading"] == file_digests,
      source=live_find["source"])

# 3. ranged and to_file reads over the live body.
file_range = server.execute_read_document(str(ORIGINAL), source="file", format="text", start_para_ref="p10", end_para_ref="p19")
live_range = server.execute_read_document(path, source="live", format="text", start_para_ref="p10", end_para_ref="p19")
check("live ranged read matches the file range", live_range["text"] == file_range["text"],
      chars=len(live_range["text"]), paragraphs=10)
for fmt in ("markdown", "runs"):
    fr = server.execute_read_document(str(ORIGINAL), source="file", format=fmt, start_para_ref="p10", end_para_ref="p19")
    lr = server.execute_read_document(path, source="live", format=fmt, start_para_ref="p10", end_para_ref="p19")
    check(f"live ranged {fmt} read matches the file range", lr[fmt] == fr[fmt])
out = Path(tempfile.gettempdir()).parent / "probe-110-body.md"
out.unlink(missing_ok=True)  # to_file refuses an existing file unless overwrite=true
summary = server.execute_read_document(path, source="live", format="markdown", to_file=str(out))
written = out.read_bytes()
tokens = len(json.dumps(summary, separators=(",", ":"))) // 4
check("live to_file: summary only, sha matches, under 500 tokens",
      summary["sha256"] == hashlib.sha256(written).hexdigest() and "markdown" not in summary and tokens < 500,
      tokens=tokens, chars=summary["chars"], section_count=summary["section_count"])

# 4. the revision is accepted as revision_before; a stale one is still refused.
edit = {"find": "Section 020", "replace": "Section 020 (accepted)", "expected_matches": 1}
fresh = reads_live.live_revision(path)["revision"]
done = text_edit.execute_replace_text_batch_live(path, [edit], fresh)
check("write with live_revision's revision is accepted", done["applied"] is True and done["batch_status"] == "complete",
      batch_status=done["batch_status"])
stale = text_edit.execute_replace_text_batch_live(path, [{**edit, "find": "Section 021", "replace": "S21"}], fresh)
check("the now-stale revision is still refused",
      stale["batch_status"] == "none_applied" and stale["results"][0]["code"] == "LIVE_STALE",
      code=stale["results"][0].get("code"), message=stale["results"][0].get("message"))

# 5. old-pane fallback against the real body_ooxml (the capability is hidden from the server only).
reads_live.BODY_HASH_CAPABILITY = "body_hash_v1_not_offered"
t0 = time.monotonic()
fallback = reads_live.live_revision(path)
fallback_s = time.monotonic() - t0
reads_live.BODY_HASH_CAPABILITY = "body_hash_v1"
native = reads_live.live_revision(path)
check("fallback reports via and agrees with the native result",
      fallback["via"] == "body_ooxml_fallback" and fallback["revision"] == native["revision"]
      and fallback["sections"] == native["sections"],
      seconds=round(fallback_s, 3), sections=len(fallback["sections"]))

failed = [r["check"] for r in results if not r["pass"]]
print(json.dumps({"summary": {"checks": len(results), "failed": failed}}), flush=True)
sys.exit(1 if failed else 0)
