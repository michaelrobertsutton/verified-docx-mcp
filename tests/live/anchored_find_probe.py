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
  (d) a stale paragraph hash (PARA_HASH_MISMATCH) and a stale revision (LIVE_STALE) refused.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from verified_docx_mcp import locate, mutations, server, text_edit
from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, open_pane_live, reads_live

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
APPROVED = "Skyward will operate the help desk " + "with tiered escalation and measured response times, " * 11 + \
    "reported monthly to the program office."

# 0. Known body, written in file mode before Word opens the document.
mutations.execute_replace_body_markdown(
    path,
    f"# Anchored\n\n{DECOY}\n\n{LONG}\n\n{PARA}\n\n# Approved\n\n{APPROVED}\n\nClosing paragraph.",
    allow_concurrent_editor=True,
)
registry = bridge.start_in_background()
opened = open_pane_live.execute_open_pane(path, timeout_s=90)
session = registry.get(Path(path).name)
caps = sorted(session.hello.capabilities)
check("pane connected with anchored_find_v1", "anchored_find_v1" in caps, method=opened["method"])


def live_lines() -> list[str]:
    return server.execute_read_document(path, source="live", format="text")["text"].split("\n")


lines = live_lines()
long_live = next(line for line in lines if line.startswith("Skyward will modernize the") and line.endswith("2027."))
check("setup: the long paragraph kept NBSP and curly quotes", NBSP in long_live and LDQ in long_live,
      chars=len(long_live))

# (a) live ellipsis replace, typed with straight quotes and plain spaces.
find = f'Skyward will modernize the "claims" platform {E}batch jobs one workload at a time by 2027.'
rev = reads_live.live_revision(path)["revision"]
evidence = text_edit.execute_replace_text(
    path, find, "Skyward will modernize the claims platform by 2027.", None,
    expected_length=len(long_live), revision_before=rev, write_mode="live",
)
lines = live_lines()
check("(a) ellipsis replace of a long paragraph",
      evidence["applied"] and "Skyward will modernize the claims platform by 2027." in lines and DECOY in lines
      and long_live not in lines, match_count=evidence["match_count"], before_chars=len(evidence["before"]))
check("(a) expected_length off by 12 is refused",
      code_of(text_edit.execute_replace_text, path, f"Skyward will operate the help desk{E}to the program office.",
              "x", None, expected_length=len(APPROVED) - 12, write_mode="live") == "ZERO_MATCH")

# (b) para_ref replace.
index = live_lines().index(PARA)
digest = locate.span_sha256(PARA)
evidence = text_edit.execute_replace_text(
    path, None, "Para-ref target replaced.", None, para_ref=f"p{index}", expected_sha256=digest, write_mode="live",
)
check("(b) para_ref replace", "Para-ref target replaced." in live_lines(), para_ref=f"p{index}")

# (d) stale paragraph hash and stale revision.
check("(d) stale paragraph hash is PARA_HASH_MISMATCH",
      code_of(text_edit.execute_replace_text, path, None, "x", None, para_ref=f"p{index}", expected_sha256=digest,
              write_mode="live") == "PARA_HASH_MISMATCH")
check("(d) stale revision_before is LIVE_STALE",
      code_of(text_edit.execute_replace_text, path, f"Skyward will operate the help desk{E}to the program office.",
              "x", None, expected_length=len(APPROVED), revision_before=rev, write_mode="live") == "LIVE_STALE")

# (c) approved run: a ~600-character ellipsis replacement.
ops = [{"op": "replace_text", "section": "approved",
        "args": {"find": f"Skyward will operate the help desk{E}reported monthly to the program office.",
                 "replace": "Skyward operates the help desk.", "expected_length": len(APPROVED)}}]
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
