"""Tool-call ARGUMENT size for the #113 cheaper-write-input tools, through the MCP layer.

    PYTHONPATH=<checkout>/src python scripts/measure_write_inputs.py

Model OUTPUT tokens are the expensive direction, and a tool call's output is its arguments.
Every scenario runs each variant through the real FastMCP server in-process
(`fastmcp.Client(server.mcp)`), on its own copy of a repo fixture, checks that no call
errored, and prints the compact JSON size of each call's ARGUMENTS summed over the
variant's calls. Tokens are characters / 4: an estimate, not a tokenizer count.

For `markdown_file=` variants only the PATH counts: the file is assumed to be on disk
already because the workflow wrote it (JennyStack draft-to-word writes `clean.md` first).
The write of that file is not an argument of any call measured here.

Add a scenario (e.g. part A's ellipsis finds) by decorating a function that returns a
`Scenario` with `@scenario`; `SCENARIOS` is the list the runner walks.
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
import tempfile
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="measure-inputs-"))
os.environ["VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"] = str(TMP)
os.environ["XDG_STATE_HOME"] = str(TMP / "state")

from fastmcp import Client  # noqa: E402

from verified_docx_mcp import mutations, server  # noqa: E402

REPO = Path(__file__).resolve().parents[1]
FIXTURE = REPO / "tests" / "fixtures" / "sections.docx"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

Call = tuple[str, dict]


class Workspace:
    """A fresh copy of the fixture (and a place for input files) per variant."""

    _count = 0

    def __init__(self) -> None:
        Workspace._count += 1
        self.dir = TMP / f"w{Workspace._count}"
        self.dir.mkdir()
        self.docx = str(self.dir / "doc.docx")
        shutil.copyfile(FIXTURE, self.docx)

    def write(self, name: str, text: str) -> str:
        path = self.dir / name
        path.write_text(text, encoding="utf-8")
        return str(path)


@dataclass
class Scenario:
    name: str
    # variant label -> builder of the ordered tool calls (arguments included).
    variants: dict[str, Callable[[Workspace], list[Call]]]
    baseline: str = ""  # the variant the others are compared with (default: the first)
    prepare: Callable[[], Workspace] = Workspace  # builds each variant's fresh workspace
    notes: list[str] = field(default_factory=list)


SCENARIOS: list[Callable[[], Scenario]] = []


def scenario(fn: Callable[[], Scenario]) -> Callable[[], Scenario]:
    SCENARIOS.append(fn)
    return fn


# --- scenarios ------------------------------------------------------------------------------

PARAGRAPHS = [
    f"Skyward will deliver requirement {n} within the first ninety days of award, reporting monthly."
    for n in range(1, 13)
]
FINDINGS = [
    f"Finding {n}: no owner or date is named for this commitment; name the accountable Skyward role and "
    "tie it to the staffing table so the evaluator can verify it."
    for n in range(1, 13)
]
QUOTES = [f"Skyward will deliver requirement {n} within" for n in range(1, 13)]
SECTION_MARKDOWN = (
    "## Background\n\n"
    + "\n\n".join(
        f"Skyward delivered requirement {n} for the incumbent program, with measurable results against "
        "the performance work statement and a documented past-performance reference."
        for n in range(1, 11)
    )
    + "\n"
)


@scenario
def review_comments() -> Scenario:
    def prepared() -> Workspace:
        ws = Workspace()
        mutations.execute_replace_body_markdown(ws.docx, "\n\n".join(PARAGRAPHS), force=True)
        return ws

    def one_by_one(ws: Workspace) -> list[Call]:
        return [
            ("add_anchored_comment", {"path": ws.docx, "quote": q, "text": c, "expected_matches": 1, "write_mode": "file"})
            for q, c in zip(QUOTES, FINDINGS, strict=True)
        ]

    def batch(ws: Workspace) -> list[Call]:
        return [
            (
                "add_anchored_comment_batch",
                {"path": ws.docx, "write_mode": "file",
                 "comments": [{"quote": q, "content": c} for q, c in zip(QUOTES, FINDINGS, strict=True)]},
            )
        ]

    return Scenario(
        "12 review comments",
        {"12 x add_anchored_comment": one_by_one, "1 x add_anchored_comment_batch": batch},
        notes=["the batch is also 1 model turn instead of 12"],
        prepare=prepared,
    )


@scenario
def markdown_section_write() -> Scenario:
    def inline(ws: Workspace) -> list[Call]:
        return [("replace_range_markdown", {"path": ws.docx, "section_key": "background-1", "markdown": SECTION_MARKDOWN})]

    def from_file(ws: Workspace) -> list[Call]:
        # The file already exists on disk (the workflow wrote it); only the path is emitted.
        file = ws.write("clean.md", SECTION_MARKDOWN)
        return [("replace_range_markdown", {"path": ws.docx, "section_key": "background-1", "markdown_file": file})]

    return Scenario(
        f"markdown section write ({len(SECTION_MARKDOWN)} chars)",
        {"replace_range_markdown(markdown=...)": inline, "replace_range_markdown(markdown_file=...)": from_file},
        notes=["markdown_file: path only; the file is assumed already on disk"],
    )


# Part A hook: add ellipsis/anchored-find scenarios here (or anywhere above) with @scenario.


# --- runner ---------------------------------------------------------------------------------


def arg_chars(args: dict) -> int:
    return len(json.dumps(args, separators=(",", ":"), ensure_ascii=False))


async def run_variant(
    client: Client, builder: Callable[[Workspace], list[Call]], prepare: Callable[[], Workspace]
) -> tuple[int, int, int]:
    calls = builder(prepare())
    chars = 0
    for tool, args in calls:
        result = await client.call_tool(tool, args, raise_on_error=False)
        if result.is_error:
            text = "".join(getattr(block, "text", "") for block in result.content)
            raise SystemExit(f"{tool} failed: {text[:300]}")
        chars += arg_chars(args)
    return len(calls), chars, chars // 4


async def main() -> None:
    rows = []
    async with Client(server.mcp) as client:
        for factory in SCENARIOS:
            sc = factory()
            baseline = sc.baseline or next(iter(sc.variants))
            measured = {label: await run_variant(client, builder, sc.prepare) for label, builder in sc.variants.items()}
            base_chars = measured[baseline][1]
            print(f"\n{sc.name}")
            for label, (calls, chars, tokens) in measured.items():
                saving = "" if label == baseline else f"  ({100 * (base_chars - chars) / base_chars:.0f}% fewer than baseline)"
                print(f"  {label:<50} calls={calls:<3} arg_chars={chars:<6} ~tokens={tokens}{saving}")
                rows.append({"scenario": sc.name, "variant": label, "calls": calls, "arg_chars": chars, "tokens": tokens})
            for note in sc.notes:
                print(f"  note: {note}")
    print("\n" + json.dumps(rows, separators=(",", ":")))


if __name__ == "__main__":
    asyncio.run(main())
    sys.stdout.flush()
