"""Response-size measurements for the #112 write tools, through the MCP layer.

    PYTHONPATH=<checkout>/src python scripts/measure_write_tokens.py <large.docx>

Each scenario runs on its own copy of the ~20K-word document (see
scripts/make_large_fixture.py), calls the real FastMCP server in-process so the
model-facing result is what is measured, and prints the compact JSON size of the
tool result or error. Run once against a checkout from before the change and once
after. Tokens are characters / 4: an estimate, not a tokenizer count.
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

SOURCE = Path(sys.argv[1]).resolve()
TMP = Path(tempfile.mkdtemp(prefix="measure-write-"))
os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = str(TMP)
os.environ["XDG_STATE_HOME"] = str(TMP / "state")

from fastmcp import Client

from verified_docx_mcp import server


def fresh() -> str:
    target = TMP / f"doc-{len(list(TMP.glob('doc-*.docx')))}.docx"
    shutil.copyfile(SOURCE, target)
    return str(target)


async def call(client: Client, tool: str, args: dict) -> tuple[bool, str]:
    result = await client.call_tool(tool, args, raise_on_error=False)
    text = "".join(getattr(block, "text", "") for block in result.content)
    return result.is_error, text


async def main() -> None:
    async with Client(server.mcp) as client:

        async def measure(name: str, tool: str, args: dict) -> None:
            is_error, text = await call(client, tool, args)
            if is_error and "unexpected keyword" in text:
                print(json.dumps({"scenario": name, "tokens": None, "note": "n/a"}))
                return
            compact = json.dumps(json.loads(text[text.index("{") :]), separators=(",", ":")) if "{" in text else text
            print(json.dumps({"scenario": name, "error": is_error, "chars": len(compact), "tokens": len(compact) // 4}))

        await measure("append_markdown, one paragraph", "append_markdown", {"path": fresh(), "markdown": "One more paragraph."})
        await measure(
            "replace_body_markdown, small replacement",
            "replace_body_markdown",
            {"path": fresh(), "markdown": "# Only\n\nSmall replacement body.", "force": True},
        )
        await measure(
            "replace_range_markdown, one section",
            "replace_range_markdown",
            {"path": fresh(), "section_key": "section-007-1", "markdown": "# Section 007\n\nRewritten section body."},
        )
        await measure(
            "format_text, bold one phrase",
            "format_text",
            {"path": fresh(), "find": "Section 020", "style": {"bold": True}, "expected_matches": 1, "write_mode": "file"},
        )
        await measure(
            "replace_text, one edit",
            "replace_text",
            {"path": fresh(), "find": "Section 030", "replace": "Section 030 (revised)", "expected_matches": 1, "write_mode": "file"},
        )
        await measure(
            "MATCH_COUNT_MISMATCH, a string with 150 hits",
            "replace_text",
            {"path": fresh(), "find": "Section", "replace": "S", "expected_matches": 1, "write_mode": "file"},
        )
        await measure(
            "ZERO_MATCH",
            "replace_text",
            {"path": fresh(), "find": "this text is not in the document", "replace": "x", "expected_matches": 1, "write_mode": "file"},
        )


asyncio.run(main())
