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

Issue #111 applies the same idea to READ results (:func:`lean_read_result`): a ``path``
equal to the argument, a null ``section_key``, empty ``warnings`` / ``lossy_elements``,
``revision_detail`` (unless ``verbose=True``), the live hash repeated as
``live.body_sha256`` and the prose ``live_session.action`` are dropped from what the model
sees. JennyStack reads ``path``, ``source``, ``format``, ``text`` and ``revision`` of
``read_document`` and ``path`` / ``source`` of ``list_open_items``, so those never go.
"""

from __future__ import annotations

import os
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
_KEEPS_REVISION_BEFORE = frozenset({"replace_text_batch"})


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


# Every read-only tool (the ones annotated readOnlyHint in server.py).
READ_TOOLS: frozenset[str] = frozenset(
    {
        "read_document",
        "find_sections",
        "list_page_sections",
        "list_tables",
        "get_table",
        "list_open_items",
        "get_comment_thread",
        "list_styles",
        "list_parts",
        "list_shapes",
        "list_textboxes",
        "read_textbox",
        "read_header_footer",
        "diff_body_vs_file",
        "live_status",
        "live_list_sections",
        "live_revision",
        "lock_status",
        "export_pdf",
        "read_audit_evidence",
    }
)

# JennyStack reads `path` from these results (live_doc_stamp.py / live_guard_state.py
# resolve_read_response), so it is never dropped even when it equals the argument.
_KEEPS_PATH = frozenset({"read_document", "list_open_items"})
_LIVE_PREFIX = "live:sha256:"


def _same_path(result_path: Any, argument: Any) -> bool:
    """True when the result's (resolved) path is the path the caller passed."""
    if not isinstance(result_path, str) or not isinstance(argument, str):
        return False
    return result_path == argument or os.path.realpath(result_path) == os.path.realpath(argument)


def lean_read_result(tool: str, arguments: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    """*result* without read-envelope noise; a new dict, *result* itself is untouched."""
    if "error_code" in result:
        return result
    out = dict(result)
    if tool not in _KEEPS_PATH and "path" in out and _same_path(out["path"], arguments.get("path")):
        del out["path"]
    for key in ("section_key", "lossy_elements", "warnings"):
        if key in out and (out[key] is None or out[key] == [] or out[key] == ""):
            del out[key]
    if "revision_detail" in out and (arguments.get("verbose") is not True or out["revision_detail"] is None):
        del out["revision_detail"]
    live = out.get("live")
    if isinstance(live, dict) and "body_sha256" in live:
        revision = out.get("revision")
        if revision in (live["body_sha256"], f"{_LIVE_PREFIX}{live['body_sha256']}"):
            live = {k: v for k, v in live.items() if k != "body_sha256"}
            if live:
                out["live"] = live
            else:
                del out["live"]
    session = out.get("live_session")
    if isinstance(session, dict) and "action" in session:
        out["live_session"] = {k: v for k, v in session.items() if k in ("document_url", "reason")}
    parts = out.get("headers_and_footers")
    if tool == "read_header_footer" and isinstance(parts, list):
        out["headers_and_footers"] = [
            {k: v for k, v in part.items() if not (k in ("warnings", "lossy_elements") and not v)}
            if isinstance(part, dict)
            else part
            for part in parts
        ]
    return out


class LeanEnvelopeMiddleware(Middleware):
    """Applies :func:`lean_result` to every successful mutating tool result and
    :func:`lean_read_result` to every successful read tool result."""

    async def on_call_tool(
        self,
        context: MiddlewareContext[mt.CallToolRequestParams],
        call_next: CallNext[mt.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        result = await call_next(context)
        tool = context.message.name
        if result.is_error or (tool not in MUTATING_TOOLS and tool not in READ_TOOLS):
            return result
        sc = result.structured_content
        if not isinstance(sc, dict):
            return result
        shrink = lean_result if tool in MUTATING_TOOLS else lean_read_result
        lean = shrink(tool, context.message.arguments or {}, sc)
        if lean is sc:
            return result
        return ToolResult(structured_content=lean, meta=result.meta)
