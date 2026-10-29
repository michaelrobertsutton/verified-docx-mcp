"""Response-size measurements for the #110 read tools, on the ~20K-word document.

    PYTHONPATH=<checkout>/src python scripts/measure_read_tokens.py <large.docx>

Run it once against a checkout from before the change and once after; a scenario
that needs a parameter the older code lacks reports "n/a". Size is the compact JSON
of the tool result; tokens are estimated as characters / 4 (the same rough measure
the issue uses -- an estimate, not a tokenizer count).
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

DOCX = Path(sys.argv[1]).resolve()
TMP = Path(tempfile.mkdtemp(prefix="measure-read-"))
os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = f"{DOCX.parent}:{TMP}"

from verified_docx_mcp import projection, server


def size(name: str, call) -> None:
    try:
        result = call()
    except TypeError as exc:  # a parameter this checkout does not have yet
        print(json.dumps({"scenario": name, "chars": None, "tokens": None, "note": f"n/a ({exc})"}))
        return
    chars = len(json.dumps(result, separators=(",", ":"), ensure_ascii=False))
    detail = {k: result[k] for k in ("identical", "identical_raw") if isinstance(result, dict) and k in result}
    if isinstance(result, dict) and "hunks" in result:
        detail["changed_hunks"] = sum(1 for h in result["hunks"] if h.get("tag") != "equal")
    print(json.dumps({"scenario": name, "chars": chars, "tokens": chars // 4, **detail}))


body_md, _, _ = projection.read_document_markdown(DOCX)
identical_md = TMP / "identical.md"
identical_md.write_text(body_md, encoding="utf-8")

# Differences a text pipeline introduces without changing any words: NBSP for some
# spaces, trailing spaces, CRLF line endings, a final newline.
typographic_md = TMP / "typographic.md"
typographic_md.write_bytes(
    (body_md.replace("program office", "program office").replace("\n", "  \r\n") + "\r\n").encode("utf-8")
)

edited = body_md.replace("Section 010", "Section 010 (revised)").replace("Section 075", "Section 075 (revised)")
edited = edited.replace("Section 140", "Section 140 (revised)")
edited_md = TMP / "edited.md"
edited_md.write_text(edited, encoding="utf-8")

d = str(DOCX)
size("diff_body_vs_file: identical document", lambda: server.execute_diff_body_vs_file(d, str(identical_md)))
size("diff_body_vs_file: typographic-only differences", lambda: server.execute_diff_body_vs_file(d, str(typographic_md)))
size("diff_body_vs_file: three real edits", lambda: server.execute_diff_body_vs_file(d, str(edited_md)))
size("read_document markdown, whole body", lambda: server.execute_read_document(d, format="markdown"))
size(
    "read_document to_file",
    lambda: server.execute_read_document(d, format="markdown", to_file=str(TMP / "body.md")),
)
size(
    "read_document markdown, 10 paragraphs",
    lambda: server.execute_read_document(d, format="markdown", start_para_ref="p100", end_para_ref="p109"),
)
size("find_sections", lambda: server.execute_find_sections(d))
size("find_sections digests=true", lambda: server.execute_find_sections(d, digests=True))
