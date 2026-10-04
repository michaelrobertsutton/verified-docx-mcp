"""Opening the Live pane on Word for Mac (#66); never touches Word."""
from pathlib import Path
from unittest.mock import patch

import pytest

from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import open_pane_live, open_pane_mac

DOC = str(Path.home() / "doc.docx")


class Ui:
    """Scripted stand-in for osascript and the mouse; records clicks."""

    def __init__(self, *, button="absent", centre="2244,135", opens_on_click=2):
        self.button, self.centre, self.opens_on_click = button, centre, opens_on_click
        self.clicks: list[tuple[int, int]] = []
        self.connected = False

    def osascript(self, script, *args, **kw):
        if script is open_pane_mac._PRESS_LIVE_PANE:
            if self.button == "clicked":
                self.connected = True
            return self.button
        if script is open_pane_mac._ADDINS_CENTRE:
            return self.centre
        return "ready"

    def click(self, x, y):
        self.clicks.append((x, y))
        entry = (2244 + open_pane_mac.POPOVER_ENTRY_DX, 135 + open_pane_mac.POPOVER_ENTRY_DY)
        if (x, y) == entry and self.clicks.count(entry) >= self.opens_on_click:
            self.connected = True


def _run(ui, **kw):
    with patch.object(open_pane_mac, "_osascript", ui.osascript), patch.object(open_pane_mac, "_click", ui.click), \
         patch.object(open_pane_mac.time, "sleep"), patch.object(open_pane_mac.platform, "system", return_value="Darwin"):
        return open_pane_mac.open_pane("doc.docx", lambda: ui.connected, **kw)


def test_ribbon_button_is_preferred_and_needs_no_mouse():
    ui = Ui(button="clicked")
    assert _run(ui) == "ribbon-button"
    assert ui.clicks == []


def test_popover_click_position_is_computed_from_the_addins_button():
    ui = Ui()
    assert _run(ui) == "addins-popover"
    assert ui.clicks[0] == (2244, 135)
    entry = (2244 + open_pane_mac.POPOVER_ENTRY_DX, 135 + open_pane_mac.POPOVER_ENTRY_DY)
    assert ui.clicks[1:] == [entry, entry]  # the first click only hovers the entry; the second activates it


def test_activating_on_the_first_click_does_not_click_again():
    ui = Ui(opens_on_click=1)
    assert _run(ui) == "addins-popover"
    assert len(ui.clicks) == 2


def test_gives_up_with_a_hand_instructions_message_and_never_claims_success():
    ui = Ui(opens_on_click=99)
    with pytest.raises(open_pane_mac.OpenPaneError) as err:
        _run(ui, timeout_s=1e9)
    assert "Developer Add-ins" in str(err.value) and "Accessibility" in str(err.value)
    assert not ui.connected


def test_missing_addins_button_is_reported():
    ui = Ui(centre="absent")
    with pytest.raises(open_pane_mac.OpenPaneError) as err:
        _run(ui)
    assert "Add-ins button not found" in str(err.value)


def test_not_macos_refuses():
    with patch.object(open_pane_mac.platform, "system", return_value="Linux"), pytest.raises(open_pane_mac.OpenPaneError):
        open_pane_mac.open_pane("doc.docx", lambda: False)


class _Registry:
    def __init__(self, present):
        self.present = present

    def get(self, name):
        return object() if self.present else None


def test_tool_returns_immediately_when_already_connected():
    with patch.object(open_pane_live.live_bridge, "start_in_background", return_value=_Registry(True)), \
         patch.object(open_pane_mac, "open_pane") as opener:
        result = open_pane_live.execute_open_pane(DOC)
    opener.assert_not_called()
    assert result["already_connected"] is True and result["opened"] is False


def test_tool_opens_the_document_then_the_pane():
    registry = _Registry(False)
    with patch.object(open_pane_live.live_bridge, "start_in_background", return_value=registry), \
         patch.object(open_pane_live, "_open_documents", side_effect=[[], ["doc.docx"]]), \
         patch.object(open_pane_live.subprocess, "run") as run, patch.object(open_pane_live.time, "sleep"), \
         patch.object(open_pane_mac, "open_pane", return_value="ribbon-button") as opener:
        result = open_pane_live.execute_open_pane(DOC)
    assert run.call_args.args[0] == ["open", "-a", "Microsoft Word", DOC]
    assert opener.call_args.args[0] == "doc.docx"
    assert result == {"opened": True, "already_connected": False, "document_name": "doc.docx", "method": "ribbon-button"}


def test_tool_failure_is_a_typed_error_with_the_manual_instructions():
    with patch.object(open_pane_live.live_bridge, "start_in_background", return_value=_Registry(False)), \
         patch.object(open_pane_live, "_open_documents", return_value=["doc.docx"]), \
         patch.object(open_pane_mac, "open_pane", side_effect=open_pane_mac.OpenPaneError("by hand: Home > Add-ins")), \
         pytest.raises(VerifyError) as err:
        open_pane_live.execute_open_pane(DOC)
    assert "by hand" in str(err.value)
