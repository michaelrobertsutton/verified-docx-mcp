"""Live revision state never comes from an on-disk stand-in (#48)."""
from typing import Any

from ..errors import ErrorCode, _make_error
from . import comments_live, write_mode


def list_revisions(session: Any) -> dict[str, Any]:
    if "live_revisions" not in session.hello.capabilities:
        return {"revisions": None, "coverage": "unavailable", "reason": "pane lacks live_revisions"}
    state = comments_live._request(session, "revisions_list")
    revisions, markup = state.get("revisions"), state.get("ooxml_revision_count")
    if revisions is not None and isinstance(markup, int) and markup > len(revisions):
        # Word's API merged or dropped items: never present this list as complete.
        state = {**state, "coverage": "partial", "reason": (
            f"Word's API returned {len(revisions)} revisions but the document body has {markup}; "
            "adjacent tracked insertions and deletions can be merged into one item, so this list is incomplete.")}
    return state


def mutate_revisions(path: str, action: str, ids: list[str] | None,
                     revision_before: str | None) -> dict[str, Any]:
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "live_revisions", feature_description="live revisions")
    if ids is not None and (not ids or len(ids) != len(set(ids))):
        raise _make_error(ErrorCode.INVALID_INPUT, "revision_ids must be nonempty and unique")
    pre = comments_live._request(session, "describe")["bodySha256"]
    write_mode.check_not_stale(revision_before, pre)
    result = comments_live._request(session, f"revisions_{action}",
                                    {"revision_ids": ids, "expectedBodySha256": pre})
    evidence = write_mode.live_evidence(
        applied=True, match_count=result["changed_count"], rung=1,
        before=str(result["before_count"]), after=str(result["after_count"]),
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool=f"{action}_tracked_changes", path=path)
    evidence.update(result)
    return evidence
