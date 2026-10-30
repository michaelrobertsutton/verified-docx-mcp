"""Response size of the #108 acceptance session: 20 cell edits plus 3 table reads.

    PYTHONPATH=<checkout>/src python scripts/measure_table_session.py

Runs through the real FastMCP server in-process (`fastmcp.Client(server.mcp)`), so what is
measured is the model-facing result after the lean envelope. Every call uses the tools'
DEFAULTS, so running it against a checkout from before #108 and against one after compares
default to default. The session, on a copy of tests/fixtures/approved-changes.docx (a 12x4
table): read the table, edit 10 cells, read it again, edit 10 more, read it a third time.

Prints one JSON line per call kind and a summary with the total response characters, an
estimate of tokens (characters / 4, not a tokenizer count) and the sha256 of the final
word/document.xml, so the two runs can be checked to produce the same document.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="measure-table-"))
os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = str(TMP)
os.environ["XDG_STATE_HOME"] = str(TMP / "state")

from fastmcp import Client

from verified_docx_mcp import mutations, server

mutations._QUIESCE_INTERVAL_SECONDS = 0.02
FIXTURE = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "approved-changes.docx"

# 20 (row, cell) targets: rows 2..11, cells 1..2 (the header row and the "Keep" column stay).
EDITS = [(row, cell) for row in range(2, 12) for cell in (1, 2)]


def _size(result) -> int:
    text = "".join(getattr(block, "text", "") for block in result.content)
    if "{" in text:
        text = json.dumps(json.loads(text[text.index("{") :]), separators=(",", ":"))
    return len(text)


async def main() -> None:
    doc = TMP / "session.docx"
    shutil.copyfile(FIXTURE, doc)
    sizes: dict[str, list[int]] = {"get_table": [], "replace_cell_markdown": []}
    async with Client(server.mcp) as client:

        async def call(tool: str, args: dict) -> None:
            result = await client.call_tool(tool, {"path": str(doc), **args}, raise_on_error=False)
            if result.is_error:
                raise SystemExit(f"{tool} failed: {_size(result)} chars")
            sizes[tool].append(_size(result))

        await call("get_table", {"table_id": 1, "source": "file"})
        for n, (row, cell) in enumerate(EDITS):
            await call("replace_cell_markdown", {"table_id": 1, "row_index": row, "cell_index": cell,
                                                 "markdown": f"Edited {row}:{cell}", "write_mode": "file"})
            if n == 9:
                await call("get_table", {"table_id": 1, "source": "file"})
        await call("get_table", {"table_id": 1, "source": "file"})

    for tool, values in sizes.items():
        print(json.dumps({"tool": tool, "calls": len(values), "chars": sum(values), "tokens": sum(values) // 4}))
    with zipfile.ZipFile(doc) as z:
        body = z.read("word/document.xml")
    total = sum(sum(v) for v in sizes.values())
    print(json.dumps({"summary": {"calls": sum(len(v) for v in sizes.values()), "chars": total, "tokens": total // 4,
                                  "document_xml_sha256": hashlib.sha256(body).hexdigest()}}))


asyncio.run(main())
