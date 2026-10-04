"""Open the Live pane in Word for Mac without a human click (#66).

Word exposes no API to open an add-in pane, so this drives its UI:

1. If the ribbon already shows the add-in's top-level "Live pane" button (it does once the
   add-in has run in this Word session), press it through the accessibility tree.
2. Otherwise open the ribbon's **Add-ins** popover and click the developer add-in in it. The
   popover is not in the accessibility tree, so it needs real mouse events; the click position
   is computed from the Add-ins button's own on-screen position, not hard-coded.

Success is always verified by the pane actually connecting (``is_connected``), never assumed.
Needs the calling app to have macOS Accessibility permission.
"""
import platform
import subprocess
import time
from collections.abc import Callable

# Offset of the "verified-do..." entry in the Add-ins popover from the Add-ins button's centre
# (measured on Word for Mac 16.113.3: popover anchors under the button).
POPOVER_ENTRY_DX = -108
POPOVER_ENTRY_DY = 357

_JXA_CLICK = """
ObjC.import('CoreGraphics');
function clickAt(argv) {
  const pt = $.CGPointMake(parseFloat(argv[0]), parseFloat(argv[1]));
  const move = $.CGEventCreateMouseEvent(null, $.kCGEventMouseMoved, pt, $.kCGMouseButtonLeft);
  $.CGEventPost($.kCGHIDEventTap, move);
  delay(0.2);
  const down = $.CGEventCreateMouseEvent(null, $.kCGEventLeftMouseDown, pt, $.kCGMouseButtonLeft);
  const up = $.CGEventCreateMouseEvent(null, $.kCGEventLeftMouseUp, pt, $.kCGMouseButtonLeft);
  $.CGEventPost($.kCGHIDEventTap, down);
  delay(0.08);
  $.CGEventPost($.kCGHIDEventTap, up);
  return "ok";
}
"""

_PREPARE = """
on run argv
  set docName to item 1 of argv
  tell application "Microsoft Word"
    activate
    activate object (first document whose name is docName)
  end tell
  -- Wide enough that the ribbon's right-hand groups are not cut off (Word clamps this to the screen).
  -- Deliberately no Finder/other-app calls: each would trigger its own macOS permission prompt.
  tell application "Microsoft Word" to set bounds of front window to {0, 40, 4000, 1000}
  delay 1
  -- Word's window titles drop the extension ("report", not "report.docx").
  set AppleScript's text item delimiters to "."
  set parts to text items of docName
  if (count of parts) > 1 then set docBase to (items 1 thru -2 of parts) as text
  if (count of parts) = 1 then set docBase to docName
  set AppleScript's text item delimiters to ""
  tell application "System Events" to tell process "Microsoft Word"
    set frontmost to true
    -- `activate object` changes Word's active document, but window 1 can still be another window: raise
    -- the target explicitly and refuse to continue on the wrong one (the pane would attach to it).
    set target to first window whose name contains docBase
    perform action "AXRaise" of target
    delay 0.5
    if (name of window 1) does not contain docBase then error "could not bring " & docName & " to the front"
    -- The Home tab click TOGGLES the ribbon: only click it when the ribbon contents are not exposed.
    if (count of scroll areas of tab group 1 of window 1) = 0 then
      click radio button "Home" of tab group 1 of window 1
      delay 2
    end if
  end tell
  return "ready"
end run
"""

_PRESS_LIVE_PANE = """
tell application "System Events" to tell process "Microsoft Word"
  repeat with g in (every group of scroll area 1 of tab group 1 of window 1)
    try
      click (first button of g whose name is "Live pane")
      return "clicked"
    end try
  end repeat
  return "absent"
end tell
"""

_ADDINS_CENTRE = """
tell application "System Events" to tell process "Microsoft Word"
  repeat with g in (every group of scroll area 1 of tab group 1 of window 1)
    try
      set b to first button of g whose name contains "ins" and name contains "Add"
      set p to position of b
      set s to size of b
      set cx to (((item 1 of p) + (item 1 of s) / 2) as integer) as text
      set cy to (((item 2 of p) + (item 2 of s) / 2) as integer) as text
      return cx & "," & cy
    end try
  end repeat
  return "absent"
end tell
"""


class OpenPaneError(RuntimeError):
    """The pane could not be opened; the message says what to do by hand."""


def _osascript(script: str, *args: str, language: str | None = None, timeout: float = 30) -> str:
    cmd = ["osascript"] + (["-l", language] if language else []) + ["-e", script, *args]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
    if proc.returncode != 0:
        raise OpenPaneError(f"osascript failed: {proc.stderr.strip()[:300]}")
    return proc.stdout.strip()


def _click(x: int, y: int) -> None:
    # (not named `run`: osascript would invoke that automatically and click twice)
    proc = subprocess.run(["osascript", "-l", "JavaScript", "-e", _JXA_CLICK + f'\nclickAt(["{x}","{y}"]);'],
                          capture_output=True, text=True, timeout=30, check=False)
    if proc.returncode != 0:
        raise OpenPaneError(f"mouse click failed: {proc.stderr.strip()[:300]}")


def _wait(is_connected: Callable[[], bool], seconds: float) -> bool:
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if is_connected():
            return True
        time.sleep(0.5)
    return is_connected()


def open_pane(document_name: str, is_connected: Callable[[], bool], timeout_s: float = 60) -> str:
    """Open the Live pane for the open Word document *document_name*; return the method used."""
    if platform.system() != "Darwin":
        raise OpenPaneError("Opening the Live pane automatically is only implemented for Word for Mac")
    deadline = time.monotonic() + timeout_s
    problems: list[str] = []
    for attempt in range(1, 4):
        if time.monotonic() > deadline:
            break
        try:
            _osascript(_PREPARE, document_name)
            if is_connected():
                return "already-connected"
            if _osascript(_PRESS_LIVE_PANE) == "clicked" and _wait(is_connected, 8):
                return "ribbon-button"
            centre = _osascript(_ADDINS_CENTRE)
            if centre == "absent":
                problems.append("Add-ins button not found on the ribbon")
                continue
            cx, cy = (int(v) for v in centre.split(","))
            _click(cx, cy)
            time.sleep(2.5)
            entry = (cx + POPOVER_ENTRY_DX, cy + POPOVER_ENTRY_DY)
            _click(*entry)  # first click hovers the entry...
            if _wait(is_connected, 3):
                return "addins-popover"
            _click(*entry)  # ...the second one activates it
            if _wait(is_connected, 8):
                return "addins-popover"
            problems.append(f"attempt {attempt}: popover entry at {entry} did not open the pane")
        except (OpenPaneError, subprocess.TimeoutExpired, ValueError) as exc:
            problems.append(f"attempt {attempt}: {exc}")
        try:  # close any popover left open before retrying
            _osascript('tell application "System Events" to key code 53')
        except OpenPaneError:
            pass
    raise OpenPaneError(
        "Could not open the Live pane automatically (" + "; ".join(problems[-3:]) + "). "
        "Open it by hand: Home > Add-ins > Developer Add-ins > verified-docx-mcp, and check that the calling "
        "app has macOS Accessibility permission.")
