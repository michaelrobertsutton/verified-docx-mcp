#!/usr/bin/env python3
"""WP-1 id-correlation script
(https://github.com/michaelrobertsutton/JennyStack/issues/106).

Settles the question the plan calls out explicitly: how does Office.js's
`Comment.id` (what the task pane reports) relate to the OOXML `durableId`
this server's `list_open_items` reports as `comment_id`, and to the raw
`w:id`?

Usage
-----
    python3 scripts/compare_comment_ids.py PANE_JSON DOCX_PATH

`PANE_JSON` is the file the lead saved after clicking "Copy JSON" in the
task pane (docs/live-mode.md) -- either the full report object
(`{"comments": [...], ...}`) or a bare list of the same comment objects.
Each pane comment is expected to carry `id`, `content`, `authorName`,
`creationDate`, `resolved`, `anchorText` (see `addin/taskpane.js`'s
`collectComments`).

`DOCX_PATH` is the `.docx` the pane's report came from -- normally a
scratch copy of `tests/fixtures/comments/multipara-comment.docx`, per
docs/live-mode.md (never the fixture in place).

This script never writes to the `.docx`; it only calls
`tracked_changes.execute_list_open_items`, the same read tool
`list_open_items` exposes.

Matching method
----------------
`list_open_items` only reports OPEN (unresolved) comments (see that
function's own docstring), so a pane comment reported as `resolved:
true` legitimately has no server-side match here -- that is reported as
a note, not an error. Remaining pane/server comments are paired by the
best combined similarity of (a) comment content and (b) anchor text
(`difflib.SequenceMatcher`, case- and whitespace-normalized), picked
greedily highest-score-first so one server comment is never claimed by
two pane comments. A pair below `MATCH_THRESHOLD` is left unmatched
rather than guessed.
"""

from __future__ import annotations

import argparse
import difflib
import json
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))

from verified_docx_mcp import tracked_changes  # noqa: E402

MATCH_THRESHOLD = 0.55


def _normalize(text: str | None) -> str:
    return " ".join((text or "").split()).strip().lower()


def _load_pane_comments(pane_json_path: Path) -> list[dict[str, Any]]:
    raw = json.loads(pane_json_path.read_text(encoding="utf-8"))
    if isinstance(raw, list):
        comments = raw
    elif isinstance(raw, dict) and isinstance(raw.get("comments"), list):
        comments = raw["comments"]
    else:
        raise SystemExit(
            f"{pane_json_path}: expected a JSON list of comments, or an object with a "
            "'comments' list (the pane's full report) -- got neither."
        )
    normalized = []
    for c in comments:
        normalized.append(
            {
                "id": c.get("id", ""),
                "content": c.get("content", ""),
                "authorName": c.get("authorName", ""),
                "creationDate": c.get("creationDate", ""),
                "resolved": bool(c.get("resolved", False)),
                "anchorText": c.get("anchorText", ""),
            }
        )
    return normalized


def _score(pane_c: dict[str, Any], server_c: dict[str, Any]) -> float:
    content_score = difflib.SequenceMatcher(
        None, _normalize(pane_c["content"]), _normalize(server_c["content"])
    ).ratio()
    anchor_score = difflib.SequenceMatcher(
        None, _normalize(pane_c["anchorText"]), _normalize(server_c["quoted_text"])
    ).ratio()
    # Content match carries more weight: the anchor text Office.js returns
    # for a multi-paragraph comment's range can legitimately differ from
    # the OOXML-side quoted_text (paragraph-join whitespace), while the
    # comment body itself should match closely either way.
    return 0.7 * content_score + 0.3 * anchor_score


def _match(
    pane_comments: list[dict[str, Any]], server_comments: list[dict[str, Any]]
) -> tuple[list[tuple[dict[str, Any], dict[str, Any], float]], list[dict[str, Any]], list[dict[str, Any]]]:
    candidates = []
    for pi, pc in enumerate(pane_comments):
        for si, sc in enumerate(server_comments):
            candidates.append((_score(pc, sc), pi, si))
    candidates.sort(key=lambda t: t[0], reverse=True)

    matched_pane: set[int] = set()
    matched_server: set[int] = set()
    pairs: list[tuple[dict[str, Any], dict[str, Any], float]] = []
    for score, pi, si in candidates:
        if score < MATCH_THRESHOLD:
            break
        if pi in matched_pane or si in matched_server:
            continue
        matched_pane.add(pi)
        matched_server.add(si)
        pairs.append((pane_comments[pi], server_comments[si], score))

    unmatched_pane = [pc for i, pc in enumerate(pane_comments) if i not in matched_pane]
    unmatched_server = [sc for i, sc in enumerate(server_comments) if i not in matched_server]
    return pairs, unmatched_pane, unmatched_server


def _verdict(pane_id: str, server_c: dict[str, Any]) -> str:
    if pane_id and pane_id == server_c["comment_id"]:
        return "Comment.id == durableId"
    if pane_id and pane_id == server_c["w_id"]:
        return "Comment.id == w_id"
    return "unrelated"


def _truncate(text: str, width: int = 40) -> str:
    text = text.replace("\n", " ")
    return text if len(text) <= width else text[: width - 1] + "…"


def _print_table(pairs: list[tuple[dict[str, Any], dict[str, Any], float]]) -> list[str]:
    headers = ["Office.js Comment.id", "server comment_id (durableId)", "server w_id", "match", "verdict"]
    rows = []
    verdicts = []
    for pane_c, server_c, score in pairs:
        verdict = _verdict(pane_c["id"], server_c)
        verdicts.append(verdict)
        rows.append(
            [
                str(pane_c["id"]),
                str(server_c["comment_id"]),
                str(server_c["w_id"]),
                f"{score:.2f}",
                verdict,
            ]
        )
    widths = [max(len(h), *(len(r[i]) for r in rows)) if rows else len(h) for i, h in enumerate(headers)]

    def fmt_row(row: list[str]) -> str:
        return "  ".join(cell.ljust(w) for cell, w in zip(row, widths))

    print(fmt_row(headers))
    print(fmt_row(["-" * w for w in widths]))
    for row in rows:
        print(fmt_row(row))
    return verdicts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pane_json", help="Path to the JSON the pane's 'Copy JSON' button produced.")
    parser.add_argument("docx_path", help="Path to the .docx the pane's report was read from (a scratch copy).")
    args = parser.parse_args(argv)

    pane_json_path = Path(args.pane_json).expanduser().resolve()
    docx_path = Path(args.docx_path).expanduser().resolve()

    pane_comments = _load_pane_comments(pane_json_path)
    resolved_in_pane = [c for c in pane_comments if c["resolved"]]
    open_pane_comments = [c for c in pane_comments if not c["resolved"]]

    result = tracked_changes.execute_list_open_items(str(docx_path))
    server_comments = result["comments"]

    print(f"Pane comments:   {len(pane_comments)} ({len(resolved_in_pane)} resolved, skipped from matching)")
    print(f"Server comments: {len(server_comments)} (open only -- list_open_items filters resolved)")
    print()

    pairs, unmatched_pane, unmatched_server = _match(open_pane_comments, server_comments)
    if not pairs:
        print("No pairs matched above threshold; nothing to correlate.")
    verdicts = _print_table(pairs)

    if unmatched_pane:
        print()
        print(f"Unmatched pane comments ({len(unmatched_pane)}):")
        for c in unmatched_pane:
            print(f"  id={c['id']!r} content={_truncate(c['content'])!r}")
    if unmatched_server:
        print()
        print(f"Unmatched server comments ({len(unmatched_server)}):")
        for c in unmatched_server:
            print(f"  comment_id={c['comment_id']!r} w_id={c['w_id']!r} content={_truncate(c['content'])!r}")

    print()
    distinct = set(verdicts)
    if not distinct:
        print("VERDICT: no matched pairs -- cannot settle the id-correlation question from this run.")
    elif distinct == {"Comment.id == durableId"}:
        print(f"VERDICT: Comment.id == durableId for all {len(verdicts)} matched pair(s).")
    elif distinct == {"Comment.id == w_id"}:
        print(f"VERDICT: Comment.id == w_id for all {len(verdicts)} matched pair(s).")
    elif distinct == {"unrelated"}:
        print(f"VERDICT: Comment.id is unrelated to both durableId and w_id for all {len(verdicts)} matched pair(s).")
    else:
        print(f"VERDICT: mixed -- see per-row verdicts above ({dict.fromkeys(verdicts)}).")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
