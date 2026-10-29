"""Real-Word acceptance for #112 (bounded live write responses).

    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp \
        python tests/live/write_tokens_probe.py /private/tmp/word-112-large.docx

The argument is a DISPOSABLE copy of the large fixture (scripts/make_large_fixture.py); the
probe edits it in Word, and starts from a pristine copy. This process must own the bridge
ports (53135/53136). Prints one JSON line per check; exits non-zero if any check fails.
"""

from __future__ import annotations

import asyncio
import json
import re
import sys
import time
from pathlib import Path

from fastmcp import Client

from verified_docx_mcp import audit, server, text_edit
from verified_docx_mcp.live import bridge, open_pane_live, paragraphs_live, reads_live

path = str(Path(sys.argv[1]).resolve())
results: list[dict] = []


def check(name: str, ok: bool, **detail) -> None:
    results.append({"check": name, "pass": bool(ok), **detail})
    print(json.dumps(results[-1], default=str), flush=True)


def tokens(value) -> int:
    return len(json.dumps(value, separators=(",", ":"))) // 4


registry = bridge.start_in_background()
opened = open_pane_live.execute_open_pane(path, timeout_s=90)
session = registry.get(Path(path).name)
caps = session.hello.capabilities
check("pane has compact_evidence_v1 and body_hash_v1", {"compact_evidence_v1", "body_hash_v1"} <= set(caps),
      method=opened["method"])


def revision() -> str:
    return reads_live.live_revision(path)["revision"]


# 1. live delete_paragraph on the 300-paragraph document: counts and neighbours only.
before_rev = revision()
t0 = time.monotonic()
deleted = paragraphs_live.delete_paragraph(path, "Section 100", before_rev)
check("delete_paragraph is compact (under 500 tokens)",
      tokens(deleted) < 500 and "before_paragraphs" not in deleted and "after_paragraphs" not in deleted,
      tokens=tokens(deleted), seconds=round(time.monotonic() - t0, 2),
      paragraphs=(deleted.get("paragraphs_before"), deleted.get("paragraphs_after")))
check("delete_paragraph reports the neighbouring paragraphs",
      deleted["neighbours"]["previous"] is not None and deleted["neighbours"]["next"] is not None,
      previous=(deleted["neighbours"]["previous"] or "")[:30], next=(deleted["neighbours"]["next"] or "")[:30])
check("delete_paragraph evidence_ref fetches the audit record",
      bool(deleted.get("evidence_ref")) and audit.find_audit_record(deleted["evidence_ref"]) is not None,
      evidence_ref=deleted.get("evidence_ref"))
check("the paragraph is really gone", "Section 100" not in server.execute_read_document(
    path, source="live", format="text")["text"].split("\n"))

# 2. a 50-edit batch: compact, hook fields intact, every edit applied.
edits = [{"find": f"Section {n:03d}", "replace": f"Sec {n:03d}", "expected_matches": 1} for n in range(1, 51)]
rev = revision()
t0 = time.monotonic()
batch = text_edit.execute_replace_text_batch_live(path, edits, rev)
seconds = round(time.monotonic() - t0, 1)
hook_ok = (
    batch["applied"] is True and batch["batch_status"] == "complete" and batch["applied_count"] == 50
    and batch["edit_count"] == 50 and batch["revision_before"] == rev
    and re.fullmatch(r"live:sha256:[0-9a-f]{64}", batch["revision_after"]) is not None
    and all(r["status"] == "applied" and r["evidence"]["match_count"] == 1 for r in batch["results"])
    and batch["results"][-1]["evidence"]["revision_after"] == batch["revision_after"]
)
check("50-edit batch: every edit applied, hook fields intact", hook_ok, seconds=seconds, tokens=tokens(batch))
check("a write chained from the batch's revision_after is accepted",
      text_edit.execute_replace_text_batch_live(
          path, [{"find": "Section 051", "replace": "Sec 051", "expected_matches": 1}], batch["revision_after"]
      )["batch_status"] == "complete")

# 3. a mid-batch refusal names the index, code and message.
mid = text_edit.execute_replace_text_batch_live(
    path,
    [{"find": "Section 052", "replace": "Sec 052", "expected_matches": 1},
     {"find": "no such heading anywhere", "replace": "x", "expected_matches": 1},
     {"find": "Section 053", "replace": "Sec 053", "expected_matches": 1}],
    revision(),
)
bad = mid["results"][1]
check("mid-batch refusal reports index, code and message",
      mid["batch_status"] == "partial" and (bad["index"], bad["code"]) == (1, "ZERO_MATCH") and bool(bad.get("message"))
      and mid["results"][2]["status"] == "not_attempted",
      code=bad.get("code"), message=(bad.get("message") or "")[:60])

# 4. errors, through the MCP layer so the lean envelope and error shaping apply.
async def call(tool: str, args: dict):
    async with Client(server.mcp) as client:
        return await client.call_tool(tool, args, raise_on_error=False)


def error_json(result) -> dict:
    text = result.content[0].text
    return json.loads(text[text.index("{"):])


stale = asyncio.run(call("replace_text", {"path": path, "find": "Section 060", "replace": "Sec 060",
                                          "expected_matches": 1, "write_mode": "live",
                                          "revision_before": rev}))  # rev is now old
sj = error_json(stale)
check("LIVE_STALE carries 12-hex prefixes and no retryable",
      sj["error_code"] == "LIVE_STALE" and set(sj["diagnostics"]) == {"expected", "actual"}
      and all(re.fullmatch(r"[0-9a-f]{12}", v) for v in sj["diagnostics"].values()) and "retryable" not in sj,
      diagnostics=sj["diagnostics"])
many = asyncio.run(call("replace_text", {"path": path, "find": "Section", "replace": "S", "expected_matches": 1,
                                         "write_mode": "live", "revision_before": revision()}))
mj = error_json(many)
check("MATCH_COUNT_MISMATCH from the live pane is bounded", mj["error_code"] == "MATCH_COUNT_MISMATCH"
      and tokens(mj) < 300, tokens=tokens(mj), diagnostics_keys=sorted(mj["diagnostics"]))
zero = asyncio.run(call("replace_text", {"path": path, "find": "text that is not here", "replace": "x",
                                         "expected_matches": 1, "write_mode": "live", "revision_before": revision()}))
zj = error_json(zero)
check("ZERO_MATCH has no ladder_report", zj["error_code"] == "ZERO_MATCH" and "ladder_report" not in zj["diagnostics"],
      tokens=tokens(zj))

# 5. the lean envelope on a real live write, and verbose=false default.
ok = asyncio.run(call("replace_text", {"path": path, "find": "Section 070", "replace": "Sec 070", "expected_matches": 1,
                                       "write_mode": "live", "revision_before": revision()}))
oj = ok.structured_content
gone = [k for k in ("verified_via", "track_changes_author", "document_name", "audit_logged", "write_mode",
                    "orphaned_comment_ids") if k in oj]
check("live write result is lean", ok.is_error is False and not gone and "revision_after" in oj and oj["applied"],
      still_present=gone, tokens=tokens(oj))

# 6. text-unchanged formatting result.
fmt = asyncio.run(call("format_text", {"path": path, "find": "Sec 001", "style": {"bold": True}, "expected_matches": 1,
                                       "write_mode": "live", "revision_before": revision()}))
fj = fmt.structured_content
check("format_text reports text_unchanged and one excerpt",
      fmt.is_error is False and fj.get("text_unchanged") is True and "excerpt" in fj and "before" not in fj,
      tokens=tokens(fj))

failed = [r["check"] for r in results if not r["pass"]]
print(json.dumps({"summary": {"checks": len(results), "failed": failed}}), flush=True)
sys.exit(1 if failed else 0)
