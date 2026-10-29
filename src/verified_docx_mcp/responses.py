"""Compact public responses; mutation verification and audit retain full evidence."""

from __future__ import annotations

import functools
import inspect
from typing import Any

# These fields are explanatory evidence, never mutation outcomes or identifiers.
_EVIDENCE_FIELDS = frozenset(
    {
        "before",
        "after",
        "runs_before",
        "runs_after",
        "fontBefore",
        "fontAfter",
        "segments",
        "matches",
        "metrics",
        "timings",
        "phase_timings",
        "revision_detail",
        "before_paragraphs",
        "after_paragraphs",
        "deleted_text",
        "before_cells",
        "after_cells",
    }
)


def compact_write(result: dict[str, Any]) -> dict[str, Any]:
    output = {k: v for k, v in result.items() if k not in _EVIDENCE_FIELDS}
    if isinstance(output.get("results"), list):
        rows = []
        for entry in output["results"]:
            if not isinstance(entry, dict):
                rows.append(entry)
                continue
            row = {k: v for k, v in entry.items() if k != "evidence" and k not in _EVIDENCE_FIELDS}
            evidence = entry.get("evidence")
            if isinstance(evidence, dict):
                for key in ("match_count", "revision_after", "warnings"):
                    if key in evidence:
                        row[key] = evidence[key]
            rows.append(row)
        output["results"] = rows
    return output


def public_write(fn):
    """Add verbose=false to a tool signature, compacting after its audited return."""

    @functools.wraps(fn)
    def call(*args, verbose: bool = False, **kwargs):
        result = fn(*args, **kwargs)
        return result if verbose else compact_write(result)

    call.__annotations__ = {**fn.__annotations__, "verbose": bool}
    signature = inspect.signature(fn)
    call.__signature__ = signature.replace(
        parameters=[
            *signature.parameters.values(),
            inspect.Parameter(
                "verbose", inspect.Parameter.KEYWORD_ONLY, default=False, annotation=bool
            ),
        ]
    )
    return call


def compact_table(result: dict[str, Any]) -> dict[str, Any]:
    merged = []
    for row in result["rows"]:
        for cell in row:
            if cell["grid_span"] != 1 or cell["v_merge"] != "none":
                merged.append(
                    {k: cell[k] for k in ("row_index", "cell_index", "grid_span", "v_merge")}
                )
    output = {
        **result,
        "rows": [[cell["text"] for cell in row] for row in result["rows"]],
        "merged": merged,
    }
    if result.get("source") == "live" and result.get("has_merged_cells"):
        output["warnings"] = [*result.get("warnings", []), "live_merge_metadata_unavailable"]
    return output
