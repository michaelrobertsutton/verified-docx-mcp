"""Response-size measurements for the #111 read-envelope trimming, through the MCP client
(so the lean-envelope middleware and tool annotations are included).

    PYTHONPATH=<checkout>/src python scripts/measure_read_payloads.py [some.docx]

The document defaults to tests/fixtures/approved-changes.docx. Size is the compact JSON of the
tool's structured result; tokens are estimated as characters / 4 -- an estimate, not a tokenizer
count. live_status needs a connected pane and is not measured here (a unit test bounds it).
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
SRC = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else REPO / "tests" / "fixtures" / "approved-changes.docx"
TMP = Path(tempfile.mkdtemp(prefix="measure-payloads-"))
os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = str(TMP)
DOC = TMP / SRC.name
shutil.copyfile(SRC, DOC)

from fastmcp import Client

from verified_docx_mcp import server


async def main() -> None:
    async with Client(server.mcp) as client:

        async def size(name: str, tool: str, **args) -> None:
            try:
                result = await client.call_tool(tool, {"path": str(DOC), **args})
            except Exception as exc:  # noqa: BLE001 - a parameter this checkout lacks, or a tool error
                print(json.dumps({"scenario": name, "chars": None, "note": f"n/a ({type(exc).__name__}: {exc})"[:160]}))
                return
            chars = len(json.dumps(result.structured_content, separators=(",", ":"), ensure_ascii=False))
            print(json.dumps({"scenario": name, "chars": chars, "tokens": chars // 4}))

        await size("read_document text (default)", "read_document", format="text")
        await size("read_document text (verbose)", "read_document", format="text", verbose=True)
        await size("read_document markdown", "read_document")
        await size("find_sections", "find_sections")
        await size("list_tables", "list_tables")
        await size("list_styles (unfiltered)", "list_styles")
        await size("list_styles (paragraph, Heading)", "list_styles", type="paragraph", name_contains="Heading")
        await size("lock_status (default)", "lock_status")
        await size("lock_status (verbose)", "lock_status", verbose=True)
        await size("read_header_footer", "read_header_footer")


asyncio.run(main())
shutil.rmtree(TMP, ignore_errors=True)
