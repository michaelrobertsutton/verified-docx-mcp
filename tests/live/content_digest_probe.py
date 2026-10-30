"""Real-Word acceptance for #117 (content_sha12: text, tracked deletions, table structure, comments).

    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp \
        python tests/live/content_digest_probe.py /private/tmp/content-digest-probe.docx [/private/tmp/images.docx]

The first argument is a DISPOSABLE copy of tests/fixtures/content-digest.docx; the probe saves
it from Word and resolves a comment in it. The optional second argument is an image-heavy
document used for timing only (it is opened, never edited). This process must own the bridge
ports (53135/53136), so no other bridge owner may be running. Prints one JSON line per check and
exits non-zero if any check fails.
"""

from __future__ import annotations

import json
import statistics
import sys
import time
from pathlib import Path

from verified_docx_mcp import projection, server, text_edit
from verified_docx_mcp.live import bridge, comments_live, open_pane_live, reads_live

path = str(Path(sys.argv[1]).resolve())
heavy = str(Path(sys.argv[2]).resolve()) if len(sys.argv) > 2 else None
results: list[dict] = []


def check(name: str, ok: bool, **detail) -> None:
    results.append({"check": name, "pass": bool(ok), **detail})
    print(json.dumps(results[-1], default=str), flush=True)


def by_key(rows: list[list[str]]) -> dict[str, str]:
    return {key: value for key, value in rows}


registry = bridge.start_in_background()
opened = open_pane_live.execute_open_pane(path, timeout_s=90)
session = registry.get(Path(path).name)
check("pane connected with body_ooxml", "body_ooxml" in session.hello.capabilities, method=opened["method"])

# 1. Live content digests from one export, then the same document saved by Word, read as a file.
live = reads_live.live_revision(path, content_digests=True)
check("live_revision(content_digests) uses one body_ooxml export", live["via"] == "body_ooxml",
      sections=len(live["sections"]), content_sections=len(live["content_sections"]))
saved = text_edit.execute_live_save(path)
check("live_save", saved.get("saved") is True)
after_save = reads_live.live_revision(path, content_digests=True)
check("saving changed no digest", after_save["content_sections"] == live["content_sections"]
      and after_save["sections"] == live["sections"]
      and after_save["preamble_content_sha12"] == live["preamble_content_sha12"])
file_side = server.execute_find_sections(path, source="file", digests=True)
heads = [s for s in file_side["sections"] if s["kind"] == "heading"]
file_content = [[s["section_key"], s["content_sha12"]] for s in heads]
file_sha = [[s["section_key"], s["sha12"]] for s in heads]
mismatched = [(f, l) for f, l in zip(file_content, live["content_sections"], strict=False) if f != l]
check("file and live content_sha12 agree on every section", file_content == live["content_sections"],
      file_sections=len(file_content), live_sections=len(live["content_sections"]), first_mismatches=mismatched[:3])
check("file and live preamble_content_sha12 agree",
      file_side["preamble_content_sha12"] == live["preamble_content_sha12"],
      file=file_side["preamble_content_sha12"], live=live["preamble_content_sha12"])
check("file and live sha12 agree on every section", file_sha == live["sections"])
items = projection.section_content_items(Path(path))
check("the saved file still has comments, a table and a deletion",
      any(i[0] == "c" for _k, its in items["sections"] for i in its)
      and any(i[0] == "tbl" for _k, its in items["sections"] for i in its)
      and any(i[0] == "del" for _k, its in items["sections"] for i in its))
compact = server.live_list_sections(path, compact=True, content_digests=True)
check("live_list_sections(content_digests) agrees", compact["content_sections"] == live["content_sections"]
      and all(len(row) == 3 for row in compact["sections"]))

# 2. Resolve the unresolved root comment live: exactly one content_sha12 changes, no sha12.
open_items = comments_live.execute_list_open_items_live(path)
target = next((c for c in open_items["comments"] if c["content"] == "Root comment one"), None)
check("found the unresolved root comment", target is not None,
      open=[c["content"] for c in open_items["comments"]])
if target is not None:
    comments_live.execute_resolve_comment_live(path, target["comment_id"])
    resolved = reads_live.live_revision(path, content_digests=True)
    before, after = by_key(live["content_sections"]), by_key(resolved["content_sections"])
    changed = sorted(k for k in before if before[k] != after.get(k))
    check("resolve changes exactly one section's content_sha12", changed == ["overview-1"], changed=changed)
    check("resolve changes no sha12", resolved["sections"] == live["sections"])
    check("resolve leaves the preamble digest", resolved["preamble_content_sha12"] == live["preamble_content_sha12"])


# 3. Timing: 5 calls each, median seconds.
def median_seconds(target_path: str, content: bool) -> tuple[float, list[float]]:
    times = []
    for _ in range(5):
        t0 = time.monotonic()
        reads_live.live_revision(target_path, content_digests=content)
        times.append(time.monotonic() - t0)
    return round(statistics.median(times), 3), [round(t, 3) for t in times]


for content in (False, True):
    median, all_times = median_seconds(path, content)
    check(f"timing live_revision(content_digests={content})", True, median_s=median, seconds=all_times)

if heavy is not None:
    open_pane_live.execute_open_pane(heavy, timeout_s=90)
    heavy_session = registry.get(Path(heavy).name)
    reply = reads_live._fetch_body_ooxml(heavy_session)
    reply_bytes = len(json.dumps(reply).encode("utf-8"))
    for content in (False, True):
        median, all_times = median_seconds(heavy, content)
        check(f"timing (image-heavy) live_revision(content_digests={content})", True,
              median_s=median, seconds=all_times, body_ooxml_reply_bytes=reply_bytes,
              stripped_parts=len(reply.get("strippedParts") or []))

failed = [r["check"] for r in results if not r["pass"]]
print(json.dumps({"summary": {"checks": len(results), "failed": failed}}), flush=True)
sys.exit(1 if failed else 0)
