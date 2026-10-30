"""Issue #108: keep the MCP tool descriptions (the docstrings FastMCP publishes) small.

Every client pays for the whole `tools/list` response in context on each session, so the
descriptions are budgeted. Long contracts live in docs/tool-contracts.md.
"""

from __future__ import annotations

import asyncio
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from fastmcp import Client

from verified_docx_mcp import server

MAX_TOOL_CHARS = 3_000
MAX_TOTAL_CHARS = 60_000


async def _descriptions() -> dict[str, str]:
    async with Client(server.mcp) as client:
        return {t.name: t.description or "" for t in await client.list_tools()}


class ToolDescriptionBudgetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.descriptions = asyncio.run(_descriptions())

    def test_every_tool_description_fits_the_per_tool_budget(self):
        too_long = {n: len(d) for n, d in self.descriptions.items() if len(d) > MAX_TOOL_CHARS}
        self.assertEqual(too_long, {}, f"descriptions over {MAX_TOOL_CHARS} characters")

    def test_total_description_size_fits_the_budget(self):
        total = sum(len(d) for d in self.descriptions.values())
        self.assertLess(total, MAX_TOTAL_CHARS, f"{total} description characters")

    def test_replace_text_stays_under_two_thousand(self):
        # The issue's named target: it was 12,200 characters, the largest description.
        self.assertLess(len(self.descriptions["replace_text"]), 2_000)

    def test_every_tool_has_a_description(self):
        self.assertEqual([n for n, d in self.descriptions.items() if not d.strip()], [])


if __name__ == "__main__":
    unittest.main()
