"""Whole-paragraph deletion through Word, guarded and independently verified."""
from typing import Any

from .. import audit
from ..errors import ErrorCode, VerifyError, _make_error
from . import comments_live, write_mode


def delete_paragraph(path: str, anchor_text: str, revision_before: str,
                     track_changes: bool = False) -> dict[str, Any]:
    if not anchor_text or not revision_before.startswith("live:sha256:"):
        raise _make_error(ErrorCode.INVALID_INPUT, "A whole-paragraph anchor and live body revision are required")
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "delete_paragraph", feature_description="whole-paragraph deletion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    write_mode.check_not_stale(revision_before, pre)
    try:
        result = comments_live._request(session, "paragraph_delete", {
            "anchor_text": anchor_text, "expectedBodySha256": pre, "track_changes": track_changes})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="delete_paragraph", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    evidence = write_mode.live_evidence(
        applied=True, match_count=1, rung=1, before=anchor_text, after="",
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="delete_paragraph", path=path)
    evidence.update(result)
    return evidence
