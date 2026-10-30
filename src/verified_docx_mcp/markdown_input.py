"""``markdown_file=`` support for the markdown write tools (issue #113).

The model often already has the markdown on disk; passing the path instead of the text
saves it re-emitting every character. The text is resolved HERE, server-side, before any
file/live routing, so both modes see the same string a ``markdown=`` call would carry.
"""

from __future__ import annotations

import hashlib
import re
from typing import Any

from .errors import ErrorCode, _make_error
from .paths import read_input_text_file


def _read_file(markdown_file: str) -> tuple[str, dict[str, Any]]:
    _, text = read_input_text_file(markdown_file)
    return text, {
        "markdown_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "markdown_chars": len(text),
    }


def resolve_markdown_input(
    markdown: str | None, markdown_file: str | None
) -> tuple[str, dict[str, Any]]:
    """Exactly one of *markdown* / *markdown_file* -> ``(text, extra_evidence)``.

    ``extra_evidence`` is ``{}`` for inline markdown and ``markdown_sha256`` (hex of the
    UTF-8 bytes) + ``markdown_chars`` for a file, merged into the success response.
    """
    if (markdown is None) == (markdown_file is None):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "Pass exactly one of markdown or markdown_file.",
            {"markdown_given": markdown is not None, "markdown_file_given": markdown_file is not None},
        )
    if markdown is not None:
        return markdown, {}
    return _read_file(markdown_file)  # type: ignore[arg-type]


_SEPARATOR_CELL = re.compile(r"^:?-+:?$")


def _split_pipe_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|") and not line.endswith("\\|"):
        line = line[:-1]
    cells: list[str] = []
    current: list[str] = []
    i = 0
    while i < len(line):
        if line[i] == "\\" and i + 1 < len(line) and line[i + 1] == "|":
            current.append("|")
            i += 2
            continue
        if line[i] == "|":
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(line[i])
        i += 1
    cells.append("".join(current).strip())
    return cells


def parse_pipe_table(text: str) -> list[list[str]]:
    """A GFM pipe table (header row, ``---`` separator row, body rows) -> ``rows``.

    Every row, header included, becomes a list of cell-markdown strings; the separator row
    is dropped (use ``header_rows`` to mark the header). Blank lines around the table are
    ignored; anything else that is not a table row is INVALID_INPUT.
    """
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 2 or "|" not in lines[0]:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "markdown_file must hold a GFM pipe table (header row, separator row, body rows).",
        )
    separator = _split_pipe_row(lines[1])
    if not separator or not all(_SEPARATOR_CELL.match(c) for c in separator):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "markdown_file table: the second line must be the |---|---| separator row.",
        )
    rows = [_split_pipe_row(lines[0])]
    for number, line in enumerate(lines[2:], start=3):
        if "|" not in line:
            raise _make_error(ErrorCode.INVALID_INPUT, f"markdown_file table: line {number} is not a table row.")
        rows.append(_split_pipe_row(line))
    return rows


def resolve_table_rows(
    rows: list[Any] | None, markdown_file: str | None
) -> tuple[list[Any], dict[str, Any]]:
    """insert_table: exactly one of *rows* / *markdown_file* (a GFM pipe table)."""
    if (rows is None) == (markdown_file is None):
        raise _make_error(ErrorCode.INVALID_INPUT, "Pass exactly one of rows or markdown_file.")
    if rows is not None:
        return rows, {}
    text, evidence = _read_file(markdown_file)  # type: ignore[arg-type]
    return parse_pipe_table(text), evidence
