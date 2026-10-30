"""Real-Word acceptance for #113 part A (anchored find, para_ref, approved long replacements).

    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp \
        python tests/live/anchored_find_probe.py /private/tmp/word-113-anchored.docx

The argument is a DISPOSABLE .docx (any content; the probe overwrites its body in file mode
BEFORE opening Word, then edits it in Word). This process must own the bridge ports
(53135/53136), so no other bridge owner may be running, and the pane must be v68+
(anchored_find_v1). Prints one JSON line per check and exits non-zero if any check fails.

Checks:
  (a) a live ellipsis replace_text of a ~600-character paragraph whose prefix also opens
      another paragraph, with NBSP and curly quotes in the document typed as plain
      spaces / straight quotes in `find`;
  (b) a para_ref + expected_sha256 replace;
  (c) an approved run replacing a ~600-character paragraph by ellipsis;
  (d) a stale paragraph hash (PARA_HASH_MISMATCH) and a stale revision (LIVE_STALE) refused;
  (e) an anchored format_text and an anchored delete_paragraph;
  (f) a 2-edit replace_text_batch whose second, anchored edit targets the first edit's output.
(b) also replaces a > 230-character paragraph by para_ref in plain live mode. expected_length
counts characters of the normalized span (straight quotes, whitespace runs -> one space).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from verified_docx_mcp import locate, mutations, server, text_edit
from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, open_pane_live, paragraphs_live, reads_live

path = str(Path(sys.argv[1]).resolve())
results: list[dict] = []
E = locate.ELLIPSIS
NBSP, LDQ, RDQ = chr(0xA0), chr(0x201C), chr(0x201D)


def check(name: str, ok: bool, **detail) -> None:
    results.append({"check": name, "pass": bool(ok), **detail})
    print(json.dumps(results[-1], default=str), flush=True)


def code_of(fn, *args, **kwargs) -> str | None:
    try:
        fn(*args, **kwargs)
    except VerifyError as exc:
        return exc.envelope.error_code.value
    return None


FILLER = "with measured, incremental releases that keep each state partner informed "
LONG = (f"Skyward will modernize the {LDQ}claims{RDQ} platform{NBSP}" + FILLER * 7
        + "and retire the legacy mainframe batch jobs one workload at a time by 2027.")
DECOY = "Skyward will modernize the claims platform, but this paragraph never reaches the suffix."
PARA = "Para-ref target: this paragraph is replaced by position and hash, not by retyping it."
PARA_LONG = ("Long para-ref target: " + "replaced by position and hash in plain live mode, " * 6
             + "not by retyping it.")
DELETE_ME = ("Delete-me paragraph: " + "this whole paragraph is deleted by an anchored anchor_text " * 5
             + "and nothing else goes.")
FORMAT_ME = "Format-me paragraph: " + "these words turn bold through an anchored find " * 6 + "in one call."
BATCH = "Batch paragraph one starts here and is rewritten first."
BATCH_1 = "Batch paragraph rewritten by edit one, " + "then edited again by the batch's second, anchored edit " * 2 \
    + "at the end."
APPROVED = "Skyward will operate the help desk " + "with tiered escalation and measured response times, " * 11 + \
    "reported monthly to the program office."


def nlen(text: str) -> int:
    """What expected_length counts: characters of the normalized span."""
    return len(locate.normalize_span_text(text))


# 0. Known body, written in file mode before Word opens the document.
body = f"# Anchored\n\n{DECOY}\n\n{LONG}\n\n{PARA}\n\n{PARA_LONG}\n\n{DELETE_ME}\n\n{FORMAT_ME}\n\n{BATCH}\n\n# Approved\n\n{APPROVED}\n\nClosing paragraph."
mutations.execute_replace_body_markdown(path, body, allow_concurrent_editor=True)
registry = bridge.start_in_background()
opened = open_pane_live.execute_open_pane(path, timeout_s=90)
session = registry.get(Path(path).name)
caps = sorted(session.hello.capabilities)
check("pane connected with anchored_find_v1", "anchored_find_v1" in caps, method=opened["method"])


def live_lines() -> list[str]:
    return server.execute_read_document(path, source="live", format="text")["text"].split("\n")


def revision() -> str:
    return reads_live.live_revision(path)["revision"]


lines = live_lines()
long_live = next(line for line in lines if line.startswith("Skyward will modernize the") and line.endswith("2027."))
check("setup: the long paragraph kept NBSP and curly quotes", NBSP in long_live and LDQ in long_live,
      chars=len(long_live), normalized_chars=nlen(long_live))

# (a) live ellipsis replace, typed with straight quotes and plain spaces.
find = f'Skyward will modernize the "claims" platform {E}batch jobs one workload at a time by 2027.'
rev = revision()
evidence = text_edit.execute_replace_text(
    path, find, "Skyward will modernize the claims platform by 2027.", None,
    expected_length=nlen(long_live), revision_before=rev, write_mode="live",
)
lines = live_lines()
check("(a) ellipsis replace of a long paragraph",
      evidence["applied"] and "Skyward will modernize the claims platform by 2027." in lines and DECOY in lines
      and long_live not in lines, match_count=evidence["match_count"], before_chars=len(evidence["before"]))
check("(a) expected_length off by 12 is refused",
      code_of(text_edit.execute_replace_text, path, f"Skyward will operate the help desk{E}to the program office.",
              "x", None, expected_length=nlen(APPROVED) - 12, write_mode="live") == "ZERO_MATCH")

# (b) para_ref replace, short (plain find on the wire) and long (> 230: the anchored payload).
index = live_lines().index(PARA)
digest = locate.span_sha256(PARA)
text_edit.execute_replace_text(
    path, None, "Para-ref target replaced.", None, para_ref=f"p{index}", expected_sha256=digest, write_mode="live",
)
check("(b) para_ref replace", "Para-ref target replaced." in live_lines(), para_ref=f"p{index}")
long_index = live_lines().index(PARA_LONG)
text_edit.execute_replace_text(
    path, None, "Long para-ref target replaced.", None, para_ref=f"p{long_index}",
    expected_sha256=locate.span_sha256(PARA_LONG), write_mode="live",
)
check("(b) para_ref replace of a > 230-character paragraph", "Long para-ref target replaced." in live_lines()
      and PARA_LONG not in live_lines(), chars=len(PARA_LONG))

# (d) stale paragraph hash and stale revision.
check("(d) stale paragraph hash is PARA_HASH_MISMATCH",
      code_of(text_edit.execute_replace_text, path, None, "x", None, para_ref=f"p{index}", expected_sha256=digest,
              write_mode="live") == "PARA_HASH_MISMATCH")
check("(d) stale revision_before is LIVE_STALE",
      code_of(text_edit.execute_replace_text, path, f"Skyward will operate the help desk{E}to the program office.",
              "x", None, expected_length=nlen(APPROVED), revision_before=rev, write_mode="live") == "LIVE_STALE")

# (e) anchored format_text and anchored delete_paragraph.
formatted = text_edit.execute_format_text(
    path, f"Format-me paragraph: these{E}bold through an anchored find in one call.", {"bold": True}, None,
    find_sha256=locate.span_sha256(FORMAT_ME), write_mode="live",
)
check("(e) anchored format_text", formatted["applied"] and formatted["match_count"] == 1)
deleted = paragraphs_live.delete_paragraph(
    path, f"Delete-me paragraph: this whole{E}and nothing else goes.", revision(), False, nlen(DELETE_ME), None,
)
check("(e) anchored delete_paragraph", deleted["applied"] and DELETE_ME not in live_lines())

# (f) a 2-edit batch whose second, anchored edit targets the first edit's output.
batch = text_edit.execute_replace_text_batch_live(path, [
    {"find": BATCH, "expected_matches": 1, "replace": BATCH_1},
    {"find": f"Batch paragraph rewritten by edit one,{E}anchored edit at the end.",
     "expected_length": nlen(BATCH_1), "replace": "Batch paragraph final."},
], revision())
check("(f) batch: an anchored edit on the previous edit's output", batch["batch_status"] == "complete"
      and "Batch paragraph final." in live_lines(), results=batch["results"])

# (c) approved run: a ~600-character ellipsis replacement.
ops = [{"op": "replace_text", "section": "approved",
        "args": {"find": f"Skyward will operate the help desk{E}reported monthly to the program office.",
                 "replace": "Skyward operates the help desk.", "expected_length": nlen(APPROVED)}}]
prepared = runs.prepare(path, ops)
manifest = json.loads(Path(prepared["manifest_path"]).read_text())
check("(c) prepare stores the resolved full find",
      manifest["operations"][0]["args"]["find"] == APPROVED and prepared["preview"][0]["find_chars"] == len(APPROVED))
report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "anchored-find probe")
check("(c) approved run executes", report["status"] == "complete" and "Skyward operates the help desk." in live_lines(),
      status=report["status"], issues=report["issues"])

failed = [r["check"] for r in results if not r["pass"]]
print(json.dumps({"summary": {"checks": len(results), "failed": failed}}), flush=True)
sys.exit(1 if failed else 0)
