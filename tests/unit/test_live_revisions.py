"""Offline revision state and pane behavior; never contacts Word."""
import shutil
import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from verified_docx_mcp.live import comments_live, revisions_live


def test_old_pane_never_reports_placeholder_revisions():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities=set(), instance_id=None, platform=None, host=None))
    with patch.object(comments_live, "_request", return_value={}), \
         patch.object(comments_live, "_session_for", return_value=session), \
         patch.object(comments_live, "_document_name_and_path", return_value=(Path("/tmp/doc.docx"), "doc")), \
         patch.object(comments_live, "_live_state", return_value=([], [])), \
         patch.object(comments_live, "_request", return_value={"comments": [], "counts": {"total": 0, "open": 0}, "session_epoch": "epoch"}), \
         patch.object(comments_live, "_file_comments_and_suggestions", return_value=([], [{"id": "fake"}])):
        result = comments_live.execute_list_open_items_live("/tmp/doc.docx")
    assert result["pending_suggestions"] is None
    assert "pane lacks" in result["pending_suggestions_reason"]


def test_live_list_does_not_substitute_file_state():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    with patch.object(comments_live, "_request", return_value={"revisions": [{"text": "live"}], "coverage": "body"}):
        assert revisions_live.list_revisions(session)["revisions"][0]["text"] == "live"


def test_partial_coverage_is_reported_when_markup_exceeds_api_list():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    reply = {"revisions": [{"text": "a"}] * 2, "coverage": "body", "ooxml_revision_count": 5}
    with patch.object(comments_live, "_request", return_value=reply):
        state = revisions_live.list_revisions(session)
    assert state["coverage"] == "partial"
    assert "2 revisions" in state["reason"] and "has 5" in state["reason"]
    full = {**reply, "ooxml_revision_count": 2}
    with patch.object(comments_live, "_request", return_value=full):
        assert revisions_live.list_revisions(session)["coverage"] == "body"


def test_revision_js_behavior():
    node = shutil.which("node")
    assert node, "Node is required for new pane behavior tests"
    root = Path(__file__).resolve().parents[2]
    subprocess.run([node, str(root / "tests/unit/js/revisions_harness.mjs"),
                    str(root / "addin/taskpane.js")], check=True, capture_output=True, text=True)


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_full_markup_inventory_and_safe_refusal():
    import pytest

    from verified_docx_mcp.errors import VerifyError
    xml = f'''<x:document xmlns:x="{W}"><x:body><x:p>
      <x:ins x:id="1" x:author="A" x:date="d"><x:r><x:t>new</x:t></x:r></x:ins>
      <x:del x:id="2" x:author="B" x:date="e"><x:r><x:delText>old</x:delText></x:r></x:del>
    </x:p></x:body></x:document>'''
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    reply = {"revisions": [{"type": "Added", "author": "A", "date": "d",
             "text": "new", "revision_id": "api"}], "revision_ooxml": xml,
             "session_epoch": "epoch"}
    with patch.object(comments_live, "_request", return_value=reply):
        state = revisions_live.list_revisions(session)
    assert state["coverage"] == "body"
    assert [r["text"] for r in state["revisions"]] == ["new", "old"]
    assert state["revisions"][1]["author"] == "B"
    assert all(not r["actionable"] for r in state["revisions"])
    with patch.object(revisions_live.write_mode, "live_session_for") as connect:
        with pytest.raises(VerifyError):
            revisions_live.mutate_revisions("unused", "accept",
                                           [state["revisions"][1]["revision_id"]], None)
        connect.assert_not_called()


def test_unique_metadata_maps_only_complete_api_inventory():
    xml = f'<w:document xmlns:w="{W}"><w:body><w:p><w:ins w:id="7" w:author="A" w:date="d"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body></w:document>'
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    api = {"revision_id": "handle", "type": "Added", "author": "A", "date": "d", "text": "x"}
    with patch.object(comments_live, "_request", return_value={"revisions": [api], "revision_ooxml": xml}):
        item = revisions_live.list_revisions(session)["revisions"][0]
    assert item["actionable"] and item["revision_id"] == "handle"
    duplicated = xml.replace('</w:p>', '<w:ins w:id="8" w:author="A" w:date="d"><w:r><w:t>x</w:t></w:r></w:ins></w:p>')
    with patch.object(comments_live, "_request", return_value={"revisions": [api, api], "revision_ooxml": duplicated}):
        assert all(not r["actionable"] for r in revisions_live.list_revisions(session)["revisions"])


def test_markup_package_and_bad_xml():
    import pytest
    inner = f'<a:document xmlns:a="{W}"><a:body><a:p><a:rPrChange a:id="1"/></a:p></a:body></a:document>'
    package = '<q:package xmlns:q="http://schemas.microsoft.com/office/2006/xmlPackage"><q:part q:name="/word/document.xml"><q:xmlData>' + inner + '</q:xmlData></q:part></q:package>'
    assert revisions_live._markup_revisions(package, "e")[0]["type"] == "Formatted"
    for xml in ('<broken', '<!DOCTYPE x><x/>', '<x/>'):
        with pytest.raises((ValueError, revisions_live.ET.ParseError)):
            revisions_live._markup_revisions(xml, "e")
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    with patch.object(comments_live, "_request", return_value={"revisions": [], "revision_ooxml": '<broken'}):
        assert revisions_live.list_revisions(session)["coverage"] == "partial"


W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"


def _single_insert_xml(date: str, date_utc: str | None) -> str:
    extra = f' xmlns:u="{W16DU}" u:dateUtc="{date_utc}"' if date_utc else ""
    return (f'<w:document xmlns:w="{W}"><w:body><w:p><w:ins w:id="0" w:author="A" w:date="{date}"{extra}>'
            '<w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body></w:document>')


def _handle_for(api_date: str, xml: str) -> bool:
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    api = {"revision_id": "handle", "type": "Added", "author": "A", "date": api_date, "text": "x"}
    with patch.object(comments_live, "_request", return_value={"revisions": [api], "revision_ooxml": xml}):
        item = revisions_live.list_revisions(session)["revisions"][0]
    return item["actionable"] and item["revision_id"] == "handle"


def test_word_local_time_w_date_matches_via_date_utc():
    # Captured from Word for Mac 16.113.3: w:date is LOCAL time labelled Z, dateUtc is the instant,
    # Office.js reports the UTC instant with milliseconds.
    xml = _single_insert_xml("2026-10-04T10:55:00Z", "2026-10-04T14:55:00Z")
    assert _handle_for("2026-10-04T14:55:00.000Z", xml)
    assert revisions_live._markup_revisions(xml, "e")[0]["date_utc"] == "2026-10-04T14:55:00Z"


def _two_changes_xml(second_author: str) -> str:
    # w:del carries its text in w:delText; the dateUtc namespace is declared once on the root.
    return (f'<w:document xmlns:w="{W}" xmlns:u="{W16DU}"><w:body><w:p>'
            '<w:del w:id="1" w:author="A" w:date="d" u:dateUtc="2026-10-04T14:55:00Z"><w:r><w:delText>old</w:delText></w:r></w:del>'
            '<w:ins w:id="2" w:author="A" w:date="d" u:dateUtc="2026-10-04T14:55:00Z"><w:r><w:t>new</w:t></w:r></w:ins>'
            f'<w:del w:id="3" w:author="{second_author}" w:date="d" u:dateUtc="2026-10-04T14:56:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>'
            '</w:p></w:body></w:document>')


def _list_two(xml: str, api: list[dict]) -> list[dict]:
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    with patch.object(comments_live, "_request", return_value={"revisions": api, "revision_ooxml": xml}):
        return revisions_live.list_revisions(session)["revisions"]


def test_deletion_matches_although_office_js_reports_empty_text():
    # Office.js gives Deleted revisions text "" (Word for Mac 16.113.3); the markup has the text.
    iso = "2026-10-04T14:55:00.000Z"
    api = [
        {"revision_id": "h1", "type": "Deleted", "author": "A", "date": iso, "text": ""},
        {"revision_id": "h2", "type": "Added", "author": "A", "date": iso, "text": "new"},
        {"revision_id": "h3", "type": "Deleted", "author": "B", "date": "2026-10-04T14:56:00.000Z", "text": ""},
    ]
    items = _list_two(_two_changes_xml("B"), api)
    assert [(i["text"], i["revision_id"], i["actionable"]) for i in items] == [
        ("old", "h1", True), ("new", "h2", True), ("gone", "h3", True)]
    # Same author and instant on two deletions is ambiguous without text: both stay read-only.
    ambiguous = _list_two(_two_changes_xml("A").replace("14:56:00Z", "14:55:00Z"),
                          [{**api[0]}, api[1], {**api[0], "revision_id": "h4"}])
    assert [i["actionable"] for i in ambiguous] == [False, True, False]


def test_unreadable_markup_is_reported_partial_not_complete():
    # The pane omitted revision_ooxml and the count (body OOXML unreadable): the API list cannot
    # be cross-checked, so it must not be presented as complete body coverage.
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    reply = {"revisions": [{"revision_id": "h", "type": "Added", "author": "A", "date": "d", "text": "x"}],
             "coverage": "body", "ooxml_revision_count": None}
    with patch.object(comments_live, "_request", return_value=reply):
        state = revisions_live.list_revisions(session)
    assert state["coverage"] == "partial" and "could not be read" in state["reason"]
    assert len(state["revisions"]) == 1


def test_date_instants_are_not_truncated_or_guessed():
    # Sub-second difference is a different instant.
    assert not _handle_for("2026-10-04T14:55:00.500Z", _single_insert_xml("d", "2026-10-04T14:55:00Z"))
    # Without dateUtc the local-time w:date cannot be trusted as UTC: a different instant never matches.
    assert not _handle_for("2026-10-04T14:55:00.000Z", _single_insert_xml("2026-10-04T10:55:00Z", None))
    # A zone-less date only matches by exact string equality.
    assert not _handle_for("2026-10-04T14:55:00.000Z", _single_insert_xml("2026-10-04T14:55:00", None))
    assert _handle_for("2026-10-04T14:55:00", _single_insert_xml("2026-10-04T14:55:00", None))
