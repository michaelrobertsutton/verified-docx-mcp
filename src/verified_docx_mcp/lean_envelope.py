"""Lean write envelope (issue #112).

Every write used to carry roughly 150 tokens of constants the caller already knows or
can never act on: ``verified_via``, ``track_changes_author``, ``document_name``,
``orphaned_comment_ids: []``, ``audit_logged: true``, ``ledger_logged: true``,
``conflict_copy_detected: false``, an echoed ``write_mode`` and the ``revision_before``
the caller just passed. This middleware drops them from what the MODEL sees. The audit
log is written earlier, from the full evidence, so nothing it records changes, and the
in-process ``execute_*`` functions still return the full dict.

Fields that carry information stay: ``applied``, ``match_count``, ``revision_after``,
``write_mode`` when it was not requested explicitly (``auto``), a ``revision_before``
that differs from the one passed, any ``False`` / non-empty value of the flags above,
and ``orphaned_comment_ids`` / ``conflict_copies`` when there is something to say.
``replace_text_batch`` keeps ``revision_before`` always: JennyStack's hook validates it.
"""

from __future__ import annotations

from typing import Any

import mcp.types as mt
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.base import ToolResult

from .middleware import MUTATING_TOOLS

# field -> the only value that makes it noise.
_CONSTANTS: dict[str, Any] = {
    "verified_via": "word-addin",
    "track_changes_author": "word-signed-in-user",
    "audit_logged": True,
    "ledger_logged": True,
    "conflict_copy_detected": False,
}

# The batch hook compares revision_before to the token it sent, so it must always be there.
_KEEPS_REVISION_BEFORE = frozenset({"replace_text_batch", "add_anchored_comment_batch"})


def lean_result(tool: str, arguments: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    """*result* without the constants above; a new dict, *result* itself is untouched."""
    if "applied" not in result or "error_code" in result:
        return result
    out = dict(result)
    for key, constant in _CONSTANTS.items():
        # type check: 1 == True and 0 == False must not make a count look like a flag.
        if key in out and type(out[key]) is type(constant) and out[key] == constant:
            del out[key]
    out.pop("document_name", None)
    if out.get("orphaned_comment_ids") == []:
        del out["orphaned_comment_ids"]

    requested = arguments.get("write_mode")
    if requested in ("file", "live") and out.get("write_mode") == requested:
        del out["write_mode"]

    passed = arguments.get("revision_before")
    if tool not in _KEEPS_REVISION_BEFORE and passed is not None and out.get("revision_before") == passed:
        del out["revision_before"]

    if tool == "add_anchored_comment":
        # A comment never edits text: before/after are the caller's own quote, and a
        # single match's comment_ids is [comment_id].
        quote = arguments.get("quote")
        if quote is not None and out.get("before") == quote and out.get("after") == quote:
            out.pop("before", None)
            out.pop("after", None)
        if "comment_id" in out and out.get("comment_ids") == [out["comment_id"]]:
            del out["comment_ids"]
    return out


class LeanEnvelopeMiddleware(Middleware):
    """Applies :func:`lean_result` to every successful mutating tool result."""

    async def on_call_tool(
        self,
        context: MiddlewareContext[mt.CallToolRequestParams],
        call_next: CallNext[mt.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        result = await call_next(context)
        tool = context.message.name
        if tool not in MUTATING_TOOLS or result.is_error:
            return result
        sc = result.structured_content
        if not isinstance(sc, dict):
            return result
        lean = lean_result(tool, context.message.arguments or {}, sc)
        if lean is sc:
            return result
        return ToolResult(structured_content=lean, meta=result.meta)
