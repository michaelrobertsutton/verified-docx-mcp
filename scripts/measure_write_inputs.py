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

from fastmcp import Client

from verified_docx_mcp import mutations, server

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
    # False: the calls are only sized, not sent (live mode needs Word; the new-API live calls are
    # exercised by the fake-pane unit tests and the real-Word probe instead).
    execute: bool = True


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


def _old_paragraph(n: int) -> str:
    """A ~460-character proposal paragraph (the issue's example is 412), unique per n."""
    return (
        f"For requirement {n}, Skyward will meet it through a dedicated delivery lead who owns the schedule, the "
        "risk register and the weekly status report to the contracting officer's representative. The lead "
        "coordinates the engineering, test and security staff, escalates blocking issues within one business "
        "day, and confirms each milestone against the performance work statement before the government "
        f"accepts it, so that requirement {n} is delivered on time and verified by 2027."
    )


def _new_paragraph(n: int) -> str:
    """The rewrite: every sentence changes, about the same length (the model types all of it in both APIs)."""
    return (
        f"Requirement {n} is owned end to end by a named Skyward delivery manager, who keeps the integrated "
        "schedule and risk register current and briefs the COR every Tuesday. That manager directs the "
        "engineering, test and cyber staff, raises any blocker to the program manager within four working "
        "hours, and walks each milestone through the PWS acceptance checklist with the government before "
        f"sign-off, so requirement {n} lands on schedule with evidence on file by mid-2027."
    )


OLD_PARAGRAPHS = [_old_paragraph(n) for n in range(1, 11)]
NEW_PARAGRAPHS = [_new_paragraph(n) for n in range(1, 11)]
ANCHOR = 24  # prefix/suffix characters an anchored find carries (the minimum is 12)


def _anchored(text: str) -> str:
    return f"{text[:ANCHOR]}…{text[-ANCHOR:]}"


@scenario
def rewrite_and_review_session() -> Scenario:
    """The issue's acceptance session: 10 paragraph rewrites then 12 review comments."""

    def prepared() -> Workspace:
        ws = Workspace()
        mutations.execute_replace_body_markdown(ws.docx, "\n\n".join(OLD_PARAGRAPHS + PARAGRAPHS), force=True)
        return ws

    def today(ws: Workspace) -> list[Call]:
        rewrites = [
            ("replace_text", {"path": ws.docx, "find": old, "replace": new, "expected_matches": 1, "write_mode": "file"})
            for old, new in zip(OLD_PARAGRAPHS, NEW_PARAGRAPHS, strict=True)
        ]
        comments = [
            ("add_anchored_comment", {"path": ws.docx, "quote": q, "text": c, "expected_matches": 1, "write_mode": "file"})
            for q, c in zip(QUOTES, FINDINGS, strict=True)
        ]
        return rewrites + comments

    def new_api(ws: Workspace) -> list[Call]:
        rewrites = [
            ("replace_text", {"path": ws.docx, "find": _anchored(old), "expected_length": len(old), "replace": new,
                              "write_mode": "file"})
            for old, new in zip(OLD_PARAGRAPHS, NEW_PARAGRAPHS, strict=True)
        ]
        batch = (
            "add_anchored_comment_batch",
            {"path": ws.docx, "write_mode": "file",
             "comments": [{"quote": q, "content": c} for q, c in zip(QUOTES, FINDINGS, strict=True)]},
        )
        return [*rewrites, batch]

    return Scenario(
        f"session: 10 paragraph rewrites ({len(OLD_PARAGRAPHS[0])} chars) + 12 review comments, file mode",
        {"today: full find + 12 comment calls": today, "#113: anchored find + comment batch": new_api},
        prepare=prepared,
        notes=["file mode has no 230-char find cap, so 'today' is one call per rewrite (the cheapest old form)"],
    )


LIVE_FIND_CAP = 230  # today's live replace_text find limit (Word search() caps needles at 255)
REVISION = "live:sha256:" + "0" * 64  # the shape of the token every live write passes


def _chunks(text: str, cap: int = LIVE_FIND_CAP) -> list[str]:
    """Split *text* into consecutive pieces of at most *cap* characters at sentence, then
    clause, then word boundaries: how a caller chunks a long live rewrite today."""
    out: list[str] = []
    rest = text
    while len(rest) > cap:
        head = rest[:cap]
        cut = max(head.rfind(". "), head.rfind(", "), head.rfind(" "))
        out.append(rest[: cut + 1])
        rest = rest[cut + 1 :]
    return [*out, rest]


@scenario
def rewrite_and_review_session_live() -> Scenario:
    """The same session in live mode, where today's 230-character find cap splits each rewrite."""
    doc = "/Users/me/Proposals/Volume II Technical.docx"

    def today(_ws: Workspace) -> list[Call]:
        calls: list[Call] = []
        for old, new in zip(OLD_PARAGRAPHS, NEW_PARAGRAPHS, strict=True):
            olds = _chunks(old)
            # The new text is split into as many consecutive pieces (one per old piece).
            step = -(-len(new) // len(olds))
            news = [new[i * step : (i + 1) * step] for i in range(len(olds))]
            calls += [
                ("replace_text", {"path": doc, "find": o, "replace": r, "expected_matches": 1, "write_mode": "live",
                                  "revision_before": REVISION})
                for o, r in zip(olds, news, strict=True)
            ]
        calls += [
            ("add_anchored_comment", {"path": doc, "quote": q, "text": c, "expected_matches": 1, "write_mode": "live"})
            for q, c in zip(QUOTES, FINDINGS, strict=True)
        ]
        return calls

    def new_api(_ws: Workspace) -> list[Call]:
        rewrites = [
            ("replace_text", {"path": doc, "find": _anchored(old), "expected_length": len(old), "replace": new,
                              "write_mode": "live", "revision_before": REVISION})
            for old, new in zip(OLD_PARAGRAPHS, NEW_PARAGRAPHS, strict=True)
        ]
        batch = (
            "add_anchored_comment_batch",
            {"path": doc, "write_mode": "live",
             "comments": [{"quote": q, "content": c} for q, c in zip(QUOTES, FINDINGS, strict=True)]},
        )
        return [*rewrites, batch]

    return Scenario(
        "session: 10 paragraph rewrites + 12 review comments, live mode (sized, not sent)",
        {"today: 230-char chunks + 12 comment calls": today, "#113: anchored find + comment batch": new_api},
        execute=False,
        notes=[f"each {len(OLD_PARAGRAPHS[0])}-char rewrite is {len(_chunks(OLD_PARAGRAPHS[0]))} live calls today",
               "a typical real path; revision_before on every replace_text, as live writes pass it"],
    )


# --- runner ---------------------------------------------------------------------------------


def arg_chars(args: dict) -> int:
    return len(json.dumps(args, separators=(",", ":"), ensure_ascii=False))


async def run_variant(
    client: Client, builder: Callable[[Workspace], list[Call]], prepare: Callable[[], Workspace], execute: bool = True
) -> tuple[int, int, int]:
    calls = builder(prepare())
    chars = 0
    for tool, args in calls:
        result = await client.call_tool(tool, args, raise_on_error=False) if execute else None
        if result is not None and result.is_error:
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
            measured = {
                label: await run_variant(client, builder, sc.prepare, sc.execute) for label, builder in sc.variants.items()
            }
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
