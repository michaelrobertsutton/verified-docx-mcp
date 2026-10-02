"""Live text-box reads and scope revisions (#35/#46)."""
from typing import Any

from ..errors import ErrorCode, VerifyError, _make_error
from . import comments_live, write_mode


def scope_state(session: Any, scope: str, revision_before: str | None) -> dict[str, Any]:
    write_mode.require_capability(session, "textboxes", feature_description="text-box scopes")
    result = comments_live._request(session, "scope_describe", {"scope": scope})
    if revision_before and revision_before.startswith("live:scope:"):
        if revision_before != f"live:scope:{result['scopeSha256']}":
            raise _make_error(ErrorCode.LIVE_STALE, "The selected scope changed")
    else:
        write_mode.check_not_stale(revision_before, result["bodySha256"])
    return result


def read_scope(path: str, scope: str | None = None) -> dict[str, Any]:
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "textboxes", feature_description="text-box reads")
    for attempt in range(2):
        try:
            result = comments_live._request(session, "textboxes_read" if scope else "textboxes_list",
                                            {"scope": scope} if scope else {})
            return {"source": "live", **result}
        except VerifyError as exc:
            if exc.envelope.error_code != ErrorCode.HOST_SHAPE_READ_FAILED or attempt:
                raise
    raise AssertionError("unreachable")
