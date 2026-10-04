"""Tag a document so Word opens the Live pane with it (#66).

The pane writes the Office document setting ``Office.AutoShowTaskpaneWithDocument``
(manifest ``TaskpaneId`` of the same name), after which Word shows and connects the
pane every time that document is opened, without a click.
"""
from typing import Any

from . import comments_live, write_mode


def execute_autoopen(path: str, enabled: bool | None = None) -> dict[str, Any]:
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "autoopen", feature_description="auto-open pane")
    if enabled is None:
        state = comments_live._request(session, "autoopen_get")
        return {**state, "applied": False, "document_name": session.document_name}
    state = comments_live._request(session, "autoopen_set", {"enabled": enabled})
    return {**state, "applied": True, "document_name": session.document_name,
            "note": "document setting changed; the document is now marked modified"}
