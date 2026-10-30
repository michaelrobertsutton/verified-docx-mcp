"""Real-Word acceptance for #108 (partial font segments, FORMAT_DRIFT, heading reads, compact tables).

    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp \
        python tests/live/token_efficiency_probe.py /private/tmp/vdm-night.docx [TAG]

The argument is a DISPOSABLE .docx that is already open in Word, with "Next steps text." in its
body. The probe owns the bridge ports (53135/53136), then waits up to 15 minutes for the pane to
connect, so the pane can be (re)opened after the probe starts and load this checkout's pane
(v69+, `partial_font_segments_v1`). TAG (default "R3") keeps a re-run's headings distinct.
It writes its own body live, after "Next steps text.", and prints one JSON line per check.

Checks:
  (a) a live replace_text whose single segment sets only bold keeps the run's other formatting
      (italic, colour) and comes back bold;
  (b) an approved run with partial segments prepares (baseline bound in the manifest), and a
      reformat of the target between prepare and execute is refused with FORMAT_DRIFT,
      nothing written; re-prepared, it completes;
  (c) read_document(source="live", section_key=<heading>) returns only that section;
  (d) get_table(source="live") is compact by default.
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from verified_docx_mcp import paragraph_insert, server, tables, text_edit
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, reads_live

path = str(Path(sys.argv[1]).resolve())
tag = sys.argv[2] if len(sys.argv) > 2 else "R3"
results: list[dict] = []


def check(name: str, ok: bool, **detail) -> None:
    results.append({"check": name, "pass": bool(ok), **detail})
    print(json.dumps(results[-1], default=str), flush=True)


registry = bridge.start_in_background()
deadline = time.monotonic() + 900
# An older pane may reconnect first; wait for one that reports the #108 capability.
while (session := registry.get(Path(path).name)) is None or "partial_font_segments_v1" not in session.hello.capabilities:
    if time.monotonic() > deadline:
        sys.exit("no pane with partial_font_segments_v1 connected (reopen the pane in Word)")
    time.sleep(1)
caps = sorted(session.hello.capabilities)
check("pane connected with partial_font_segments_v1", "partial_font_segments_v1" in caps)

FONT_TARGET = f"Font target {tag}: the gamma word inherits."
APPROVED_TARGET = f"Approved font {tag}: proof is current."
FONTS, APPROVED, TABLE = f"Fonts {tag}", f"Approved {tag}", f"Table {tag}"
paragraph_insert.execute_insert_paragraphs(path, "Next steps text.", [
    {"text": FONTS, "style": "Heading 1"}, {"text": FONT_TARGET},
    {"text": APPROVED, "style": "Heading 1"}, {"text": APPROVED_TARGET},
    {"text": TABLE, "style": "Heading 1"}, {"text": f"Table anchor {tag}."},
], 1, write_mode="live")
approved_slug = APPROVED.lower().replace(" ", "-")


def revision() -> str:
    return reads_live.live_revision(path)["revision"]


# (a) Make "gamma" italic and red, then replace it with a bold-only segment.
text_edit.execute_format_text_live(path, "gamma word inherits", {"italic": True, "color": "#C00000"}, 1)
evidence = text_edit.execute_replace_text(
    path, "gamma word inherits", "GAMMA word inherits", 1, write_mode="live",
    replacement_segments=[{"text": "GAMMA", "format": {"bold": True}}, {"text": " word inherits", "format": {}}],
)
live_runs = server.execute_read_document(path, source="live", format="runs")["runs"]
gamma = [r for r in live_runs if "GAMMA" in r["text"]]
check("(a) bold-only segment applied", evidence["applied"] and bool(gamma))
rpr = gamma[0]["rPr"] if gamma else {}
check("(a) the segment inherited italic and colour and added bold",
      rpr.get("bold") is True and rpr.get("italic") is True and str(rpr.get("color") or "").upper().lstrip("#") == "C00000",
      rPr=rpr)

# (b) Approved run with a partial segment; a reformat between prepare and execute is FORMAT_DRIFT.
op = {"op": "replace_text", "section": approved_slug,
      "args": {"find": "proof is current.", "replace": "proof: current.", "expected_matches": 1,
               "replacement_segments": [{"text": "proof:", "format": {"bold": True}},
                                        {"text": " current.", "format": {}}]}}
prepared = runs.prepare(path, [op])
manifest = json.loads(Path(prepared["manifest_path"]).read_text())
check("(b) prepare binds the font baseline", manifest["operations"][0].get("font_baseline") is not None,
      baseline=manifest["operations"][0].get("font_baseline"))
text_edit.execute_format_text_live(path, "proof is current.", {"color": "#0070C0"}, 1)  # the "manual" reformat
report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "token-efficiency probe")
lines = server.execute_read_document(path, source="live", format="text")["text"].split("\n")
codes = [issue.get("code") for issue in report.get("issues", [])]
check("(b) reformat after prepare is FORMAT_DRIFT and nothing is written",
      report["status"] != "complete" and "FORMAT_DRIFT" in codes and APPROVED_TARGET in lines,
      status=report["status"], codes=codes)
prepared = runs.prepare(path, [op])
report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "token-efficiency probe")
lines = server.execute_read_document(path, source="live", format="text")["text"].split("\n")
check("(b) re-prepared, the run completes", report["status"] == "complete"
      and f"Approved font {tag}: proof: current." in lines, status=report["status"])

# (c) Heading-section live read.
section = server.execute_read_document(path, source="live", format="text", section_key=f"{FONTS.lower().replace(' ', '-')}-1")
text = section["text"]
check("(c) heading section read returns only that section",
      FONTS in text and "GAMMA word inherits" in text and APPROVED not in text and "Next steps text." not in text,
      chars=len(text))

# (d) Compact live get_table.
tables.execute_insert_table_live(path, [["Head A", "Head B"], ["a1", "b1"], ["a2", "b2"]],
                                 anchor={"paragraph_text": f"Table anchor {tag}.", "position": "after"})
listing = server.list_tables(path, source="live")
table_id = next(t["table_id"] for t in listing["tables"] if t.get("header_row") == ["Head A", "Head B"])
compact = server.get_table(path, table_id, source="live")
check("(d) live get_table is compact by default",
      compact.get("rows") == [["Head A", "Head B"], ["a1", "b1"], ["a2", "b2"]] and "merged" in compact,
      keys=sorted(compact))

failed = [r["check"] for r in results if not r["pass"]]
print(json.dumps({"summary": {"checks": len(results), "failed": failed}}), flush=True)
sys.exit(1 if failed else 0)
