# Vendored from GoogleDocs-MCP src/verified_googledocs_mcp/middleware.py,
# commit 374bcf8 (full file, 98 lines). Lifted per issue #28 WP-02 ("Lift,
# do not rewrite" — "middleware.py (EvidenceEnforcementMiddleware,
# MUTATING_TOOLS)"). The class body is unmodified; only MUTATING_TOOLS'
# contents differ (see below) and this header comment was added.
"""Enforcement middleware: ensures every mutating tool returns evidence.

A mutating tool that returns a bare success without an evidence payload would
defeat the verified-write contract.  This middleware is the structural backstop:
it checks that the result of any registered mutating tool contains either an
``applied`` key (normal success or dry-run) or an ``error_code`` key (a
surfaced VerifyError).  If neither is present the middleware raises so the tool
cannot report bare success.

Non-mutating tools pass through untouched.

Registration
------------
Add a tool name to MUTATING_TOOLS when a new mutating tool is wired up.

  mcp.add_middleware(EvidenceEnforcementMiddleware())
"""

from __future__ import annotations

from typing import Any

import mcp.types as mt
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.base import ToolResult

# Explicit registry of mutating tool names. Extend as new tools are added —
# in the SAME commit that adds the tool (core/document-backend-protocol.md
# §6: "every new docx mutating tool is added to this middleware's
# MUTATING_TOOLS set in the same commit that adds the tool").
#
# Empty through WP-03: export_pdf writes a local PDF file only and never
# modifies the .docx (mirrors the Google server's own docstring note that
# export_pdf is the one non-mutating exception); lock_status and every
# WP-03 read tool are read-only data. WP-04 adds the first three mutating
# tools below, added in the same commit that adds each tool
# (core/document-backend-protocol.md §6).
MUTATING_TOOLS: frozenset[str] = frozenset(
    {
        "replace_body_markdown",
        "replace_range_markdown",
        "append_markdown",
        # issue #28 WP-06:
        "replace_text",
        "format_text",
        # issue #28 WP-07 (list_open_items is read-only, not listed here):
        "accept_tracked_changes",
        "reject_tracked_changes",
        # issue #28 WP-08 (get_comment_thread is read-only, not listed here):
        "add_anchored_comment",
        # issue #28 WP-09:
        "reply_to_comment",
        "resolve_comment",
        # issue #28 WP-14 (list_tables/get_table are read-only, not listed here):
        "replace_table_row",
        "replace_cell_markdown",
        "insert_table",
        # issue #28 WP-15a (read_header_footer is read-only, not listed here):
        "insert_image",
        "apply_style",
        # issue #106 WP-3 (live_status is read-only, not listed here):
        "live_save",
    }
)

_EVIDENCE_KEYS = frozenset({"applied", "error_code"})


class EvidenceEnforcementMiddleware(Middleware):
    """Reject a mutating tool result that carries no evidence payload.

    Hooks into ``on_call_tool``.  For tools in MUTATING_TOOLS:
      - Lets ToolErrors (already typed failures) pass through unchanged.
      - Inspects the ToolResult's structured_content for an evidence key.
      - Raises RuntimeError if no evidence key is found on a normal return.

    Non-mutating tools are not touched.
    """

    async def on_call_tool(
        self,
        context: MiddlewareContext[mt.CallToolRequestParams],
        call_next: CallNext[mt.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        tool_name = context.message.name
        result = await call_next(context)

        if tool_name not in MUTATING_TOOLS:
            return result

        # Errors (is_error=True) are already typed failures — they carry the
        # VerifyError envelope as their payload, which counts as evidence.
        if result.is_error:
            return result

        # Check structured_content for an evidence key.
        sc = result.structured_content
        if isinstance(sc, dict) and _EVIDENCE_KEYS.intersection(sc.keys()):
            return result

        # Fallback: check the text content for an "applied" or "error_code" key.
        # (FastMCP serialises dict returns to JSON text when output_schema is None.)
        if result.content:
            import json

            for block in result.content:
                text = getattr(block, "text", None)
                if text:
                    try:
                        parsed: Any = json.loads(text)
                        if isinstance(parsed, dict) and _EVIDENCE_KEYS.intersection(parsed.keys()):
                            return result
                    except (json.JSONDecodeError, ValueError):
                        pass

        raise RuntimeError(
            f"Mutating tool {tool_name!r} returned a result without evidence. "
            "Every mutating tool must include 'applied' or 'error_code' in its "
            "return payload so the verified-write contract cannot be bypassed."
        )
