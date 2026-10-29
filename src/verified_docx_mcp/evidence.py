"""Bounded write evidence (issue #112).

A write that rewrites a whole body, section or revision set used to echo the whole
before and after text back to the caller (tens of thousands of tokens for a 20K-word
document). The full text now lives only in the audit log, under ``evidence_ref``; the
response carries hashes, sizes and a short excerpt at the edit boundary instead.

``compact_whole_document`` is applied AFTER ``audit.append_audit`` has stored the
full evidence, so nothing the audit log holds changes.
"""

from __future__ import annotations

import hashlib
from typing import Any

# Characters of context on each side of the first difference: a 200-character window.
EXCERPT_RADIUS = 100


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _blocks(text: str, sep: str) -> int:
    return sum(1 for block in text.split(sep) if block.strip())


def _common_prefix(a: str, b: str) -> int:
    n = min(len(a), len(b))
    i = 0
    while i < n and a[i] == b[i]:
        i += 1
    return i


def bounded_fields(before: str, after: str, *, sep: str = "\n\n") -> dict[str, Any]:
    """Hashes, sizes, block counts and a 200-character excerpt of each side around the
    first point where they differ (the end of the old text for a pure append).
    ``sep`` is the block separator: a blank line for markdown, a newline for plain text."""
    edge = _common_prefix(before, after)
    lo = max(0, edge - EXCERPT_RADIUS)
    hi = edge + EXCERPT_RADIUS
    return {
        "before_sha256": _sha256(before),
        "after_sha256": _sha256(after),
        "before_chars": len(before),
        "after_chars": len(after),
        "paragraphs_before": _blocks(before, sep),
        "paragraphs_after": _blocks(after, sep),
        "before_excerpt": before[lo:hi],
        "after_excerpt": after[lo:hi],
    }


def compact_whole_document(
    evidence: dict[str, Any],
    *,
    evidence_ref: str | None,
    sep: str = "\n\n",
) -> dict[str, Any]:
    """Replace the whole-document ``before``/``after`` in *evidence* (in place) with the
    bounded fields. ``evidence_ref`` is included only when the audit record was written,
    since a ref to a record that does not exist would dangle."""
    before = evidence.pop("before", "")
    after = evidence.pop("after", "")
    evidence.update(bounded_fields(str(before), str(after), sep=sep))
    if evidence_ref:
        evidence["evidence_ref"] = evidence_ref
    return evidence
