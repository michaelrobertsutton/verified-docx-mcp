"""Live revision state never comes from an on-disk stand-in (#48)."""
from datetime import UTC, datetime
from typing import Any
from uuid import uuid4
from xml.etree import ElementTree as ET

from ..errors import ErrorCode, _make_error
from . import comments_live, write_mode


def list_revisions(session: Any) -> dict[str, Any]:
    if "live_revisions" not in session.hello.capabilities:
        return {"revisions": None, "coverage": "unavailable", "reason": "pane lacks live_revisions"}
    state = comments_live._request(session, "revisions_list")
    revisions, markup = state.get("revisions"), state.get("ooxml_revision_count")
    xml = state.pop("revision_ooxml", None)
    if xml is not None:
        try:
            inventory = _markup_revisions(xml, state.get("session_epoch", "unknown"))
        except (ET.ParseError, ValueError):
            return {**state, "coverage": "partial", "reason": "Live revision OOXML could not be parsed"}
        # WordApiDesktop 1.4 lists every content revision, deletions included, in document order, each as
        # its own accept()/reject()-able object: align by order, verified pairwise (#58).
        desktop = state.pop("desktop_revisions", None)
        # Otherwise metadata is not a safe handle correlation when the API merges revisions.
        api = revisions or []
        if _align_desktop(inventory, desktop):
            pass
        elif len(api) == len(inventory):
            for item in inventory:
                matches = [r for r in api if _identity(r) == _identity(item)]
                peers = [r for r in inventory if _identity(r) == _identity(item)]
                if len(matches) == len(peers) == 1 and matches[0].get("revision_id"):
                    item.update(revision_id=matches[0]["revision_id"], actionable=True,
                                actionability_reason=None, source="ooxml+officejs")
        return {**state, "revisions": inventory, "ooxml_revision_count": len(inventory),
                "coverage": "body", "reason": None}
    if revisions is not None and markup is None:
        # The body markup could not be read, so the API list cannot be cross-checked.
        state = {**state, "coverage": "partial", "reason": (
            "Body revision markup could not be read, so this list may omit revisions "
            "Word merged into neighbours (e.g. deletions)")}
    elif revisions is not None and isinstance(markup, int) and markup > len(revisions):
        # Word's API merged or dropped items: never present this list as complete.
        state = {**state, "coverage": "partial", "reason": (
            f"Word's API returned {len(revisions)} revisions but the document body has {markup}; "
            "adjacent tracked insertions and deletions can be merged into one item, so this list is incomplete.")}
    return state


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
_REVISION_TYPES = {
    "ins": "Added", "del": "Deleted", "moveFrom": "MoveFrom", "moveTo": "MoveTo",
    **{name: "Formatted" for name in ("rPrChange", "pPrChange", "sectPrChange",
        "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange")},
    "cellIns": "CellInserted", "cellDel": "CellDeleted", "cellMerge": "CellMerged",
}


_DESKTOP_KIND = {"Insert": "ins", "Delete": "del", "Property": "rPrChange", "ParagraphProperty": "pPrChange",
                 "TableProperty": "tblPrChange", "SectionProperty": "sectPrChange",
                 "MovedFrom": "moveFrom", "MovedTo": "moveTo"}


def _is_structural(item: dict[str, Any]) -> bool:
    """An empty insertion/deletion mark (paragraph mark, table row/cell): in the markup, not in Word's lists."""
    return item["markup_type"] in ("ins", "del") and not item["text"]


def _align_desktop(inventory: list[dict[str, Any]], desktop: Any) -> bool:
    """Give each content revision the handle of its WordApiDesktop ``Revision`` by document order.

    Only when the two sequences agree pairwise (type, author, instant, and text for insertions);
    any disagreement leaves every entry read-only. Word's ``Revision.date`` is local time labelled
    ``Z`` like ``w:date``, so it is compared against both ``w:date`` and ``w16du:dateUtc``.
    """
    if not isinstance(desktop, list) or not desktop:
        return False
    candidates = [i for i in inventory if not _is_structural(i)]
    if len(candidates) != len(desktop):
        return False
    for item, rev in zip(candidates, desktop, strict=True):
        when = _instant(rev.get("date"))
        if (not rev.get("revision_id") or _DESKTOP_KIND.get(rev.get("type")) != item["markup_type"]
                or rev.get("author") != item["author"] or not isinstance(when, datetime)
                or when not in (_instant(item.get("date")), _instant(item.get("date_utc")))):
            return False
        if item["markup_type"] == "ins" and str(rev.get("text", "")).rstrip() != item["text"].rstrip():
            return False
    for item, rev in zip(candidates, desktop, strict=True):
        item.update(revision_id=rev["revision_id"], actionable=True, actionability_reason=None,
                    source="ooxml+officejs-desktop")
    return True


_W16DU_DATE_UTC ="{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc"


def _instant(value: Any) -> Any:
    """A timezone-aware datetime for an ISO-8601 string with an explicit zone, else the raw value.

    Never truncates: ``...00.000Z`` equals ``...00Z`` but ``...00.500Z`` does not. A string
    without a zone is left as-is, so it can only match by exact string equality.
    """
    if isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError:
            return value
        if parsed.tzinfo is not None:
            return parsed.astimezone(UTC)
    return value


def _identity(item: dict[str, Any]) -> tuple[Any, ...]:
    # Word writes w:date as LOCAL wall-clock time labelled "Z"; the true UTC instant is
    # w16du:dateUtc. Office.js reports a real UTC instant, so prefer dateUtc when present.
    when = item.get("date_utc") or item.get("date")
    # Office.js reports the text of a Deleted revision as "" (observed on Word for Mac
    # 16.113.3) while the markup carries it, so deletions are identified without text.
    # Uniqueness on both sides is still required, so identical deletions stay read-only.
    text = None if item.get("type") == "Deleted" else item.get("text")
    return (item.get("type"), item.get("author"), _instant(when), text)


def _markup_revisions(xml: str, epoch: str) -> list[dict[str, Any]]:
    if "<!DOCTYPE" in xml or "<!ENTITY" in xml:
        raise ValueError("DTD/entity declarations are not allowed")
    root = ET.fromstring(xml)
    if root.tag == _PKG + "package":
        parts = [p for p in root.findall(_PKG + "part")
                 if p.get(_PKG + "name") == "/word/document.xml"]
        if len(parts) != 1:
            raise ValueError("Missing or duplicate document part")
        document = parts[0].find(_PKG + "xmlData/" + _W + "document")
    else:
        document = root if root.tag == _W + "document" else None
    if document is None or document.find(_W + "body") is None:
        raise ValueError("Missing document body")
    body = document.find(_W + "body")
    assert body is not None
    parents = {child: parent for parent in body.iter() for child in parent}

    def text(node: ET.Element) -> str:
        return "".join(n.text or "" for n in node.iter()
                       if n.tag in (_W + "t", _W + "delText"))

    result = []
    for node in body.iter():
        kind = node.tag.removeprefix(_W)
        if node.tag != _W + kind or kind not in _REVISION_TYPES:
            continue
        paragraph = node
        while paragraph.tag != _W + "p" and paragraph in parents:
            paragraph = parents[paragraph]
        result.append({
            "revision_id": f"revision-ooxml:{epoch}:{uuid4()}",
            "ooxml_id": node.get(_W + "id"), "markup_type": kind,
            "type": _REVISION_TYPES[kind], "author": node.get(_W + "author"),
            "date": node.get(_W + "date"), "date_utc": node.get(_W16DU_DATE_UTC), "text": text(node),
            "paragraph_context": [text(paragraph)] if paragraph.tag == _W + "p" else [],
            "scope": "body", "source": "ooxml", "actionable": False,
            "actionability_reason": "No unambiguous Office.js handle; use accept/reject-all",
        })
    return result


def mutate_revisions(path: str, action: str, ids: list[str] | None,
                     revision_before: str | None) -> dict[str, Any]:
    if ids and any(i.startswith("revision-ooxml:") for i in ids):
        raise _make_error(ErrorCode.LIVE_CAPABILITY_MISSING,
                          "OOXML-only revisions cannot be edited individually; use accept/reject-all")
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "live_revisions", feature_description="live revisions")
    if ids is not None and (not ids or len(ids) != len(set(ids))):
        raise _make_error(ErrorCode.INVALID_INPUT, "revision_ids must be nonempty and unique")
    pre = comments_live._request(session, "describe")["bodySha256"]
    write_mode.check_not_stale(revision_before, pre)
    result = comments_live._request(session, f"revisions_{action}",
                                    {"revision_ids": ids, "expectedBodySha256": pre})
    evidence = write_mode.live_evidence(
        shape_result=result,
        applied=True, match_count=result["changed_count"], rung=1,
        before=str(result["before_count"]), after=str(result["after_count"]),
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool=f"{action}_tracked_changes", path=path)
    evidence.update(result)
    return evidence
