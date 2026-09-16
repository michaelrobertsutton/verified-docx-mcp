#!/usr/bin/env python3
# Vendored verbatim from JennyStack scripts/docx_render.py at commit
# 678aac1 (JennyStack PR #105, "paragraph geometry probe"), source
# sha256 4496467fe6ebf09d01d587874a198bc65d671c9e78df22cf8e1174c742a10a29,
# size 40791 bytes. Copied verbatim per issue #28 WP-02 (D3: a
# standalone, stdlib-only Python module vendored into both repos with
# a source-hash comment). Do not diverge from the JennyStack copy
# without also updating it there — see that repo's copy for the
# canonical version. Everything below this header comment, including
# the module's own docstring, is unmodified from the source file.
# NOTE: this file is deliberately exempted from `ruff check --fix` /
# autoformatting — `ruff --fix` rewrites `Optional[str]` to
# `str | None` here (a legal, semantically identical change) and
# removes the now-unused `Optional` import, which breaks
# byte-for-byte verbatim parity with the source and this header's own
# sha256. Lint it read-only; never run --fix against it (see
# pyproject.toml's per-file-ignores / the CI lint step, if either is
# later added, for the corresponding exclusion).
"""docx_render.py — render a .docx to PDF via Microsoft Word automation, and
count a PDF's pages. Issue #30 WP-04 / issue #28's identical vendored copy
(D3: a standalone Python module, stdlib only, vendored into both repos
with a source-hash comment in the copy — see that issue's own module
header for the hash).

Invocation: `uv run --python 3.12 python "$REPO/scripts/docx_render.py" …`
(also plain `python3 scripts/docx_render.py …` — nothing here needs a
third-party package). Also importable: `render_word()` and `page_count()`
are the two functions other Python code should call directly.

ENGINE LADDER (D2, ruled 2026-09-09): Word only. No LibreOffice. When Word
automation is unavailable this raises RENDER_ENGINE_UNAVAILABLE; it never
silently falls back to a different engine or a word-count estimate. The
`--engine` CLI flag accepts only "word" today so a future engine (#28) can
be added without changing the CLI shape.

=====================================================================
SPIKE FINDINGS (scripts/docx-render-probe.js, run live against Word
16.112.3 — see that file for the full JXA narrative; sdef itself needs
full Xcode and was unavailable, so Word's own embedded dictionary
resource was read directly: `/Applications/Microsoft Word.app/Contents/
Resources/Word.sdef`).

**The actual, verified-working render is CLASSIC APPLESCRIPT, not JXA.**
The first cut of this module used JXA (`Application("Microsoft
Word").open(Path(...))`, `.saveAs({fileFormat: 17})`) and reproduced a
consistent, immediate "Error: Message not understood." (AppleEvent
-1708) on `saveAs` in every JXA calling shape tried, and identically via
classic AppleScript **when the PDF format was passed as the numeric
ordinal 17**. Verified twice, by a second pass of testing: Word's
AppleScript surface will not accept a bare integer in the `file format`
slot of `save as` — it requires the ENUMERATOR by name, `format PDF` (the
literal name Word.sdef gives that WdSaveFormat enumerator), not its
numeric value. The ordinal-position research that produced `wdFormatPDF
= 17` was correct about the *value*; it just is not how this command
takes it. The two enums this module also touches:
  wdAlertsNone     = "alerts none"    (WdAlertLevel enumerator name)
  wdStatisticPages = "statistic pages" (WdStatistic enumerator name)

**Working classic-AppleScript sequence** (osascript's DEFAULT language —
no `-l JavaScript` flag — running an `on run argv` handler, argv[0] =
staged input path, argv[1] = staged output path, argv[2] = that input's
basename, all already inside Word's sandbox container per below):

```applescript
on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  set inName to item 3 of argv
  tell application "Microsoft Word"
    set display alerts to alerts none
    open POSIX file inPath
    set foundIt to false
    repeat 80 times
      try
        if (name of active document) is inName then
          set foundIt to true
          exit repeat
        end if
      end try
      delay 0.5
    end repeat
    if foundIt is false then
      error "timed out waiting for the opened document (" & inName & ") to become Word's active document (a cold Word launch, a start screen, or a document-recovery pane are the likely causes)"
    end if
    set d to active document
    set pg to compute statistics d statistic statistic pages
    save as d file name outPath file format format PDF
  end tell
  return "OK " & (pg as string)
end run
```

Two pitfalls found the hard way:
  (1) `set d to open POSIX file inPath` does NOT give a usable document
      reference — it raises "The variable d is not defined" (-2753).
      `open` must be its own statement; the document reference is then
      obtained from `active document`.
  (2) A FIXED `delay` after `open` (the first working version used
      `delay 1`) is not good enough: confirmed live when Word had been
      quit and the render triggered a COLD launch (`osascript` launches
      it automatically — the Automation grant still applies; nothing
      wrong with this being a fresh process) — a cold launch is slow
      enough, and may show a start screen or a document-recovery pane
      long enough, that `active document` is not reliably the
      just-opened file (or any usable document at all) after a single
      fixed second. The script instead POLLS: it loops (0.5s between
      tries, 80 tries = 40s bound, leaving headroom under
      `render_word`'s own `timeout`) until `name of active document`
      equals the basename it just opened, and raises a clear AppleScript
      error — mapped to `RENDER_FAILED`, never a silent hang — if that
      never happens. This is deterministic on both a warm and a cold
      Word, whereas the fixed delay only ever worked on a warm one.
  (3) `compute statistics <doc> statistic statistic pages` and
      `save as <doc> file name <path> file format format PDF` both use
      the sdef's plain-English parameter labels and enumerator NAMES,
      not their four-character or ordinal codes.

**16.112.3 finding, SUPERSEDED 2026-09-16 (see the corrected finding
below) — `close` appeared not to exist on this Word build's AppleScript
surface, for either class that could plausibly carry it.** Read directly
from Word.sdef: the `document` class declares `<responds-to>` for
NOTHING — no close, no save, no print. The `window` class DOES declare
`<responds-to command="close">` (delegating to a `handleCloseScriptCommand:`
Cocoa method) — but invoking it, in every form tried on 16.112.3 (JXA and
classic AppleScript; by document name, window name, index, and `active
document`/`active window`) failed identically with "doesn't understand
the 'close' message" (-1708). This looked like a genuine gap in Word's
Automation implementation on that build, not a syntax error on the
caller's side — confirmed by reading the dictionary, not just observing
failures. `do Visual Basic` is not in the dictionary either (fails to
compile, -2741), so there was no macro-side escape hatch. GUI automation
(System Events keystrokes to close the window) was deliberately NOT
used — it needs a separate Accessibility grant this module has no
business requesting, and is racy against whatever else is on screen.

**CORRECTED FINDING, verified live 2026-09-16 against Word 16.112.4:**
`close window "<exact window name>" saving no` WORKS — the window closes
immediately, with no save prompt, and closes ONLY that window. The
16.112.3 failure above was real on that build; it simply does not
reproduce on 16.112.4. Two other close forms were also tried against
leftover `jennystack-render-*` windows and BOTH FAIL, so this module's
close call must take exactly the working shape and no other:
  - `close document "<name>"` fails with -1728 ("object not found") —
    document names carry the `.docx` extension, window names do not, and
    `document` still declares no `<responds-to>` for `close` per the
    sdef read above.
  - `tell (first window whose …) to close saving no` fails with -1750.
  - **`close (first window whose name contains "…") saving no` is the
    most dangerous failure mode of all: it RETURNS 0 (apparent success)
    but closes the WRONG window** — a live probe closed two of the
    lead's real, unrelated documents this way. Word's `whose` filter
    against windows is UNRELIABLE and must NEVER be used here, for
    closing or anything else. The only safe reference is the EXACT
    window name this render itself captured via `name of active
    window`, taken immediately after the existing poll confirms `name of
    active document is inName` — never a `whose` clause, never an index,
    never "first window".

**Consequence: `render_word()` now closes its own staged document by
default (`close_after=True`).** A close failure is reported, never
fatal — the PDF is already written to disk before the close is
attempted, so a failed close never loses output. The result dict's
`"left_open_document"` is `None` when the close succeeded; when it fails
(or when the caller passes `close_after=False` to keep the window open
on purpose), it names the window title as before, and `"close_error"`
carries the close attempt's own error text (`None` when there was
nothing to report). `scripts/docx-doctor.sh` counts current
`jennystack-render-*` windows as its own advisory line so any leftovers
(pre-dating this change, or from a close failure) stay visible rather
than silently accumulating. A lead can close these by hand at any
point — none of them are real documents, and this module NEVER acts on
any document or window it did not just create itself (see the staging
section below on unique naming, and the safety note in
`scripts/docx-doctor.sh`).

=====================================================================
PARAGRAPH GEOMETRY PROBE (issue #102, JennyStack half — a generic,
stdlib-only probe; the verified-docx-mcp server's section-to-page
mapping built on top of this is a later, separate change). Verified
live 2026-09-16 against Word 16.112.4 on a staged 10-page copy:

```applescript
set r to text object of paragraph N of d
get range information r information type active end page number      -- integer page
get range information r information type vertical position relative to page  -- points from page top (real)
page height of page setup of d                                       -- points (792.0 for US Letter)
content of r                                                          -- the paragraph text, ends with a CR
```

`N` is a 1-based ordinal into Word's `paragraphs of document` — the
main story only (body and table-cell paragraphs in document order;
text boxes, headers, and footnotes are excluded). The probes run
inside the SAME `osascript` invocation as the render, after `compute
statistics` and strictly BEFORE `save as`/close, so the pagination the
probes observe is identical to the pagination in the saved PDF —
pagination is not guaranteed stable across a second Word session, so
measuring it in a later, separate pass would not be trustworthy. Each
ordinal is probed inside its own `try`; a bad ordinal (out of range,
or any other per-paragraph failure) emits a `PROBE<TAB>n<TAB>ERR<TAB>
<message>` line and the render continues rather than failing the
whole call. `render_word()`'s `page_height_pt` / `paragraph_geometry`
result keys and the CLI's `--probe-paragraphs` flag are documented on
`render_word()` itself, below.
=====================================================================
WORD'S FILE-ACCESS SANDBOX (found in review, NOT anticipated by the
original plan or docs/onboarding.md step 11 as merged — that step covers
only the Automation/Apple-events grant). Word is sandboxed for FILE
access as well: opening or saving a document at a path outside a
directory macOS has already granted Word access to raises a native
"<App> would like to access files in the folder ..." consent dialog —
the SAME class of failure as the Automation TCC prompt the plan already
calls out (a screen a headless process cannot answer; if it appears
under this module's `timeout` the call fails or hangs, it never blocks
indefinitely past the caller's own timeout). This reproduced for real
during development against `/private/tmp/.../scratchpad` (well outside
any of Word's granted locations) and interrupted the lead's own session
with a live consent dialog.

THE FIX (this module): every render STAGES into a scratch directory
INSIDE Word's own sandbox container —
  ~/Library/Containers/com.microsoft.Word/Data/Documents/
— which Word may read and write with NO consent dialog, because it is
Word's own app-private storage (this is what the plan's "save the PDF
into the source file's own directory (Word's sandbox)" line was gesturing
at, one level more specifically than the plan spelled out). `render_word`
copies the input .docx into a unique subdirectory there
(`tempfile.mkdtemp(dir=...)`), runs Word against that copy, and moves the
finished PDF out to the caller's requested `out_path` with `os.replace` —
falling back to `shutil.move` on `OSError` errno 18 (`EXDEV`, "Invalid
cross-device link"), which is expected here: `/private/tmp` and the
container (on the boot/home volume) are commonly different filesystems,
so a plain rename across them fails. The staging directory is removed in
a `finally` unconditionally, on both success and failure, so nothing is
ever left inside Word's container.

If that container directory does not exist at all (Word installed but
never launched at least once, so macOS has not yet created its
container), this is NOT silently treated as "render into whatever
directory was asked for" — the whole point above is that doing so
prompts. It is instead a distinct, named failure, WORD_SANDBOX_UNAVAILABLE,
naming the expected path and instructing "launch Microsoft Word once,
then retry." scripts/docx-doctor.sh checks for this directory as its own
PASS/FAIL line with the same fix text, ahead of attempting any render.

=====================================================================
ERROR CODES (core/document-backend-protocol.md's docx-backend runtime
error codes; the render-path subset):
  AUTOMATION_NOT_GRANTED    — macOS declined Automation (Apple-events)
                              control of Word. Detected here from
                              AppleEvent error numbers -1712 or -1743
                              appearing in Word's/osascript's own stderr.
  WORD_SANDBOX_UNAVAILABLE  — Word's sandbox container directory (above)
                              does not exist on this machine.
  RENDER_FAILED             — any other Word automation failure (a
                              non-zero osascript exit, a timeout, or an
                              AppleScript-reported error that isn't one
                              of the two AppleEvent numbers above). The
                              real stderr/AppleScript error text is
                              always attached — never swallowed.
  RENDER_ENGINE_UNAVAILABLE — engine requested is not "word" (D2).
=====================================================================
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

OSASCRIPT_DEFAULT = "/usr/bin/osascript"


def _osascript_bin() -> str:
    """Resolve the `osascript` binary via PATH at CALL time (not import
    time), falling back to the real macOS location. Resolving lazily via
    PATH — rather than hardcoding /usr/bin/osascript — is what lets the
    AUTOMATION_NOT_GRANTED unit test inject a fake `osascript` earlier on
    PATH (an agent cannot revoke a real macOS Automation grant to
    exercise this path for real)."""
    return shutil.which("osascript") or OSASCRIPT_DEFAULT


PDFINFO = "/opt/homebrew/bin/pdfinfo"  # confirmed present on this machine;
# _find_pdfinfo() below still falls back to PATH so this module works on a
# machine where Homebrew lives elsewhere (an Intel Mac's /usr/local/bin,
# for instance) without editing this constant.

DEFAULT_TIMEOUT = 90

# The classic-AppleScript worker script (see the module docstring for why
# JXA was abandoned and why the enumerator/active-document form below is
# required). Written to a temp file inside the SAME staging directory as
# the copied .docx (both already inside Word's sandbox container), then
# run via plain `osascript <script> <in.docx> <out.pdf> <in-basename>
# <close|keep> [ordinal ...]` (osascript's DEFAULT language is
# AppleScript; no `-l JavaScript` flag) using `on run argv`. The first
# three arguments describe the already-staged paths; the script never
# makes its own sandbox decisions. The fourth argument is "close" (close
# the staged document after rendering — render_word()'s default) or
# "keep" (leave it open, close_after=False). Any FIFTH argument onward
# is a 1-based paragraph ordinal to probe (issue #102) — see the module
# docstring's PARAGRAPH GEOMETRY PROBE section.
#
# Output is one or more lines, joined with `linefeed`, in this order:
#   PROBE<TAB>n<TAB>page<TAB>vpos<TAB>text   -- one per requested ordinal
#   PROBE<TAB>n<TAB>ERR<TAB>message          -- that ordinal failed
#   PAGEH<TAB>page-height-in-points          -- only if any ordinals were probed
#   OK <page-count> closed=<0|1> <close-error-text-or-empty>   -- ALWAYS last
# render_word() below splits stdout on newlines and treats the last
# non-empty line as the OK line, unchanged from #103; PROBE/PAGEH lines
# (if any) are everything before it. No ordinals requested -> no PROBE
# or PAGEH lines at all, output is the bare OK line exactly as before
# #102. Any failure before the OK line is produced at all (including the
# `-1712`/`-1743` automation-declined case, and the bounded
# active-document poll below timing out) raises an AppleScript error,
# which osascript reports on stderr with a non-zero exit — parsed by
# render_word() below. A close FAILURE (as opposed to a render failure)
# is never raised as an AppleScript error — it is caught, reported in
# the "closed="/error-text fields of the OK line, and never fails the
# render, because the PDF is already on disk by then. The close call
# itself must reference the window by the EXACT name this script
# captured via `name of active window` — see the module docstring's
# CORRECTED FINDING section for why a `whose` filter must NEVER be used
# here (it can close the wrong window).
_APPLESCRIPT_RENDER = r"""
on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  set inName to item 3 of argv
  set closeMode to item 4 of argv
  set probeOrdinals to {}
  if (count of argv) > 4 then
    repeat with i from 5 to (count of argv)
      set end of probeOrdinals to (item i of argv)
    end repeat
  end if
  set outputLines to {}
  tell application "Microsoft Word"
    set display alerts to alerts none
    open POSIX file inPath
    -- Bounded poll, not a fixed delay: on a COLD Word launch (osascript
    -- launches it automatically when it is not already running) a start
    -- screen or a document-recovery pane can make "active document"
    -- unready, or the wrong document, for well over a second. 80 tries
    -- * 0.5s = 40s, leaving headroom under render_word()'s own timeout
    -- for the compute-statistics/probe/save-as/close steps that follow.
    set foundIt to false
    repeat 80 times
      try
        if (name of active document) is inName then
          set foundIt to true
          exit repeat
        end if
      end try
      delay 0.5
    end repeat
    if foundIt is false then
      error "timed out waiting for the opened document (" & inName & ") to become Word's active document (a cold Word launch, a start screen, or a document-recovery pane are the likely causes)"
    end if
    set d to active document
    -- Capture the EXACT window name right after the poll confirms this
    -- is our document, and verify it belongs to it before doing
    -- anything else with it. NEVER a `whose` filter (see the module
    -- docstring: a `whose` reference against Word windows has been
    -- observed to resolve to the WRONG window).
    set winName to name of active window
    if (name of document of window winName) is not inName then
      error "active window (" & winName & ") does not belong to the just-opened document (" & inName & ") — refusing to touch a window this render did not create"
    end if
    set pg to compute statistics d statistic statistic pages
    -- Paragraph geometry probe (issue #102): runs AFTER compute
    -- statistics and strictly BEFORE save-as/close, so the pagination
    -- observed here is identical to the pagination in the saved PDF.
    -- Each ordinal is independent: a bad one emits a PROBE ERR line and
    -- the loop continues, never aborting the whole render.
    if (count of probeOrdinals) > 0 then
      repeat with probeArg in probeOrdinals
        set nStr to (probeArg as string)
        try
          set n to (probeArg as integer)
          set r to text object of paragraph n of d
          set pNum to (get range information r information type active end page number)
          set vPos to (get range information r information type vertical position relative to page)
          set txt to content of r
          if (count of txt) > 80 then
            set txt to text 1 thru 80 of txt
          end if
          if (count of txt) > 0 and (character (count of txt) of txt is return or character (count of txt) of txt is linefeed) then
            set txt to text 1 thru ((count of txt) - 1) of txt
          end if
          set txt to my _js102ReplaceChar(txt, tab, " ")
          set txt to my _js102ReplaceChar(txt, return, " ")
          set txt to my _js102ReplaceChar(txt, linefeed, " ")
          set end of outputLines to "PROBE" & tab & nStr & tab & (pNum as string) & tab & (vPos as string) & tab & txt
        on error errMsg
          set end of outputLines to "PROBE" & tab & nStr & tab & "ERR" & tab & errMsg
        end try
      end repeat
      set pgh to page height of page setup of d
      set end of outputLines to "PAGEH" & tab & (pgh as string)
    end if
    save as d file name outPath file format format PDF
    set closed to "0"
    set closeErr to ""
    if closeMode is "close" then
      try
        close window winName saving no
        if (name of every window) contains winName then
          error "window \"" & winName & "\" is still listed after close"
        end if
        set closed to "1"
      on error errMsg
        set closeErr to errMsg
      end try
    end if
  end tell
  set end of outputLines to "OK " & (pg as string) & " closed=" & closed & " " & closeErr
  set {tid, AppleScript's text item delimiters} to {AppleScript's text item delimiters, linefeed}
  set fullOutput to outputLines as text
  set AppleScript's text item delimiters to tid
  return fullOutput
end run

on _js102ReplaceChar(txt, findChar, replChar)
  set {tid, AppleScript's text item delimiters} to {AppleScript's text item delimiters, findChar}
  set theItems to text items of txt
  set AppleScript's text item delimiters to replChar
  set joined to theItems as text
  set AppleScript's text item delimiters to tid
  return joined
end _js102ReplaceChar
"""


class RenderError(Exception):
    """Raised by render_word(). .code is one of the module docstring's
    error codes; .detail carries the raw stderr/JXA text, never swallowed."""

    def __init__(self, code: str, message: str, detail: Optional[str] = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail

    def to_dict(self) -> dict:
        d = {"error_code": self.code, "error": self.message}
        if self.detail:
            d["detail"] = self.detail
        return d


def _word_sandbox_root() -> Path:
    return Path.home() / "Library" / "Containers" / "com.microsoft.Word" / "Data" / "Documents"


def _classify_automation_error(text: str) -> bool:
    """True if text names one of the two AppleEvent codes macOS uses to
    report a declined Automation (Apple-events) grant."""
    return "-1712" in text or "-1743" in text


def render_word(
    in_path: str,
    out_path: str,
    timeout: int = DEFAULT_TIMEOUT,
    close_after: bool = True,
    probe_paragraphs: Optional[list[int]] = None,
) -> dict:
    """Render in_path (.docx) to out_path (.pdf) via Word automation.

    Returns {"pdf": <out_path str>, "pages": <int|None>,
    "left_open_document": <window title str | None>, "closed_after":
    <bool>, "close_error": <str | None>, "page_height_pt": <float|None>,
    "paragraph_geometry": <dict[int, dict]>} on success. "pages" is
    Word's OWN page count (from `compute statistics ... statistic
    pages`) — the authoritative source per the module docstring; a
    caller should still cross-check against page_count() (pdfinfo/regex)
    for a second opinion, which the CLI (_main, below) does.

    close_after (default True, D3): when true, the staged document is
    closed (no save — nothing is discarded, the PDF is already written)
    after rendering. "closed_after" reports whether that close actually
    succeeded; "left_open_document" is None when it did, and the window
    title when it did not (either because close_after was False, or the
    close itself failed — "close_error" then carries the close attempt's
    own error text, None otherwise). A close failure never raises
    RenderError and never affects "pages" or the written PDF — see the
    module docstring's CORRECTED FINDING section for the verified-working
    close form and why a `whose` window reference must never be used.

    probe_paragraphs (issue #102, default None): an optional list of
    1-based ordinals into Word's `paragraphs of document` (main story
    only — see the module docstring's PARAGRAPH GEOMETRY PROBE section).
    Every ordinal must be a positive int; anything else raises
    RenderError("RENDER_FAILED", "invalid probe ordinal: ...") BEFORE
    osascript is ever launched. When None or empty, no extra argv is
    passed, no probe runs, "page_height_pt" is None, and
    "paragraph_geometry" is {}. Otherwise the probes run inside the same
    osascript invocation as the render, before save/close, so the
    pagination they observe matches the saved PDF exactly.
    "page_height_pt" is the document's page height in points (from
    `page height of page setup`). "paragraph_geometry" maps each
    requested ordinal to either {"page": int, "vpos_pt": float, "text":
    str} (the paragraph's page number, vertical position in points from
    the page top, and up to its first 80 characters) or {"error": str}
    when that one ordinal's probe failed (e.g. out of range) — a
    per-ordinal failure never aborts the render or the other probes. A
    missing PAGEH line when probes were requested (the worker and this
    parser ship together and must never drift apart) is
    RenderError("RENDER_FAILED", "unexpected form").

    Raises RenderError on any failure — never returns a partial/silent
    result, and never leaves a stray PDF in Word's sandbox staging
    directory (cleaned up in `finally` regardless of outcome) nor
    clobbers an existing file at the destination.
    """
    in_path = Path(in_path).resolve()
    out_path = Path(out_path).resolve()

    if not in_path.is_file():
        raise RenderError("RENDER_FAILED", f"input file not found: {in_path}")

    probe_paragraphs = list(probe_paragraphs) if probe_paragraphs else []
    for ordinal in probe_paragraphs:
        if isinstance(ordinal, bool) or not isinstance(ordinal, int) or ordinal < 1:
            raise RenderError(
                "RENDER_FAILED",
                f"invalid probe ordinal: {ordinal!r} (must be a positive int)",
            )

    sandbox_root = _word_sandbox_root()
    if not sandbox_root.is_dir():
        raise RenderError(
            "WORD_SANDBOX_UNAVAILABLE",
            f"Word's sandbox container directory does not exist: {sandbox_root}. "
            "Launch Microsoft Word once (so macOS creates its container), then retry.",
        )

    stage_dir = Path(tempfile.mkdtemp(prefix="jennystack-render-", dir=str(sandbox_root)))
    try:
        # The staged copy's FILENAME (not just its directory) must be
        # unique across every document Word already has open — a stale
        # document left open from an earlier render (close_after=False,
        # or a close that failed — see the module docstring) must never
        # collide with, and get mistaken for, this render's own input. The
        # "jennystack-render-" prefix also makes every window this module
        # ever opens unmistakably its own: never a real lead document, so
        # any future cleanup tooling (or a human) can act on windows
        # matching this exact pattern with total confidence.
        unique_stem = f"{stage_dir.name}-{in_path.stem}"
        staged_in = stage_dir / (unique_stem + in_path.suffix)
        staged_out = stage_dir / (unique_stem + ".pdf")
        shutil.copyfile(in_path, staged_in)

        script_path = stage_dir / "_render.applescript"
        script_path.write_text(_APPLESCRIPT_RENDER, encoding="utf-8")

        close_mode = "close" if close_after else "keep"
        probe_argv = [str(n) for n in probe_paragraphs]
        try:
            proc = subprocess.run(
                [
                    _osascript_bin(),
                    str(script_path),
                    str(staged_in),
                    str(staged_out),
                    staged_in.name,
                    close_mode,
                    *probe_argv,
                ],
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as e:
            detail = (e.stderr or "") if isinstance(e.stderr, str) else ""
            raise RenderError(
                "RENDER_FAILED",
                f"Word automation timed out after {timeout}s "
                "(a stuck native dialog is the most common cause; check the "
                "screen of the host app running Claude Code)",
                detail,
            )

        stderr = (proc.stderr or "").strip()
        stdout = (proc.stdout or "").strip()

        if proc.returncode != 0:
            if _classify_automation_error(stderr) or _classify_automation_error(stdout):
                raise RenderError(
                    "AUTOMATION_NOT_GRANTED",
                    "macOS declined this app's request to control Microsoft Word via Automation.",
                    stderr or stdout,
                )
            raise RenderError(
                "RENDER_FAILED",
                "Word automation (osascript) exited non-zero.",
                stderr or stdout or "(no output)",
            )

        # The worker's output is one or more lines (issue #102): any
        # PROBE/PAGEH lines come first, and the "OK <n> closed=<0|1>
        # <err>" line is always LAST — split on newlines and treat the
        # last non-empty line as that OK line, exactly as the worker
        # promises (see _APPLESCRIPT_RENDER's comment above).
        out_lines = [ln for ln in stdout.splitlines() if ln.strip() != ""]
        if not out_lines:
            raise RenderError(
                "RENDER_FAILED",
                "Word did not report success in the expected form.",
                stdout or "(no output)",
            )
        ok_line = out_lines[-1]
        probe_lines = out_lines[:-1]

        # ^OK\s+(\d+)\s+closed=([01])\s*(.*)$ — the "closed=" token was
        # added alongside close_after (D3); the legacy bare "OK <n>" form
        # (no "closed=" token) is now RENDER_FAILED "unexpected form", not
        # silently accepted, because the AppleScript worker and this
        # parser ship together and must never drift apart.
        m = re.match(r"^OK\s+(\d+)\s+closed=([01])\s*(.*)$", ok_line, re.DOTALL)
        if not m:
            raise RenderError(
                "RENDER_FAILED",
                "Word did not report success in the expected form.",
                stdout or "(no output)",
            )
        word_pages = int(m.group(1))
        closed_after = m.group(2) == "1"
        close_error = m.group(3).strip() or None

        # Parse PROBE/PAGEH lines (issue #102). page_height_pt stays
        # None, and paragraph_geometry stays {}, when no probes were
        # requested — the worker emits neither line in that case.
        page_height_pt: Optional[float] = None
        paragraph_geometry: dict = {}
        for ln in probe_lines:
            parts = ln.split("\t")
            if parts[0] == "PROBE" and len(parts) >= 3:
                try:
                    ordinal = int(parts[1])
                except ValueError:
                    continue
                if parts[2] == "ERR":
                    err_text = "\t".join(parts[3:]) if len(parts) > 3 else ""
                    paragraph_geometry[ordinal] = {"error": err_text}
                elif len(parts) >= 5:
                    try:
                        page_num = int(parts[2])
                        vpos_pt = float(parts[3])
                    except ValueError:
                        paragraph_geometry[ordinal] = {
                            "error": f"unparseable PROBE line: {ln!r}"
                        }
                        continue
                    text = "\t".join(parts[4:])
                    paragraph_geometry[ordinal] = {
                        "page": page_num,
                        "vpos_pt": vpos_pt,
                        "text": text,
                    }
            elif parts[0] == "PAGEH" and len(parts) >= 2:
                try:
                    page_height_pt = float(parts[1])
                except ValueError:
                    pass

        if probe_paragraphs and page_height_pt is None:
            raise RenderError(
                "RENDER_FAILED",
                "Word did not report success in the expected form.",
                "probes were requested but no PAGEH line was found in: "
                + (stdout or "(no output)"),
            )

        if not staged_out.is_file():
            raise RenderError(
                "RENDER_FAILED",
                "Word reported success but produced no PDF.",
                stdout,
            )

        out_path.parent.mkdir(parents=True, exist_ok=True)
        if out_path.exists():
            raise RenderError(
                "RENDER_FAILED",
                f"refusing to clobber an existing file at {out_path}",
            )
        try:
            os.replace(str(staged_out), str(out_path))
        except OSError as e:
            if e.errno == 18:  # EXDEV: Invalid cross-device link
                shutil.move(str(staged_out), str(out_path))
            else:
                raise

        return {
            "pdf": str(out_path),
            "pages": word_pages,
            "left_open_document": None if closed_after else staged_in.name,
            "closed_after": closed_after,
            "close_error": close_error,
            "page_height_pt": page_height_pt,
            "paragraph_geometry": paragraph_geometry,
        }
    finally:
        shutil.rmtree(stage_dir, ignore_errors=True)


def _find_pdfinfo() -> Optional[str]:
    if os.path.isfile(PDFINFO) and os.access(PDFINFO, os.X_OK):
        return PDFINFO
    return shutil.which("pdfinfo")


def page_count(pdf_path: str) -> tuple[Optional[int], Optional[str]]:
    """Return (count, source) where source is "pdfinfo", "regex", or None.

    NEVER returns 0 — a PDF page_count of 0 would flow into
    final-readiness's page budget as if it were a real, tiny document, so
    an unparseable PDF returns None (source None) instead. Prefers
    pdfinfo (poppler); falls back to a `/Type /Page` regex only when
    pdfinfo is absent, and treats a regex-derived 0 the same as
    unparseable.
    """
    pdfinfo_bin = _find_pdfinfo()
    if pdfinfo_bin:
        try:
            proc = subprocess.run(
                [pdfinfo_bin, str(pdf_path)],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if proc.returncode == 0:
                m = re.search(r"^Pages:\s*(\d+)\s*$", proc.stdout, re.MULTILINE)
                if m:
                    n = int(m.group(1))
                    if n > 0:
                        return n, "pdfinfo"
        except (OSError, subprocess.TimeoutExpired):
            pass

    try:
        data = Path(pdf_path).read_bytes()
    except OSError:
        return None, None
    n = len(re.findall(rb"/Type\s*/Page(?!s)\b", data))
    if n > 0:
        return n, "regex"
    return None, None


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="docx_render.py",
        description="Render a .docx to PDF via Word automation and report its page count.",
    )
    parser.add_argument("input", help="Path to the source .docx")
    parser.add_argument("output", help="Path to write the rendered .pdf")
    parser.add_argument(
        "--engine",
        choices=["word"],
        default="word",
        help='Render engine. Only "word" today (D2: no LibreOffice); kept as a flag '
        "so a future engine can be added without changing the CLI shape.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"Seconds to wait for Word automation (default {DEFAULT_TIMEOUT}).",
    )
    parser.add_argument(
        "--keep-open",
        action="store_true",
        help="Do not close the staged document after rendering (default: "
        "close it — D3). Useful for inspecting the rendered document in "
        "Word by hand.",
    )
    parser.add_argument(
        "--probe-paragraphs",
        default=None,
        metavar="N,N,...",
        help="Comma-separated 1-based paragraph ordinals (issue #102) to "
        "probe for page number, vertical position, and text. Adds "
        '"page_height_pt" and "paragraph_geometry" to the JSON output line.',
    )
    args = parser.parse_args(argv)

    if args.engine != "word":
        err = RenderError(
            "RENDER_ENGINE_UNAVAILABLE",
            f'no render engine available for "{args.engine}" (Word is the only engine; D2).',
        )
        print(json.dumps(err.to_dict()))
        return 1

    probe_paragraphs: Optional[list[int]] = None
    if args.probe_paragraphs:
        try:
            probe_paragraphs = [
                int(tok) for tok in args.probe_paragraphs.split(",") if tok.strip() != ""
            ]
        except ValueError:
            err = RenderError(
                "RENDER_FAILED",
                f"invalid --probe-paragraphs value: {args.probe_paragraphs!r} "
                "(must be a comma-separated list of positive integers)",
            )
            print(json.dumps(err.to_dict()))
            return 1

    try:
        result = render_word(
            args.input,
            args.output,
            timeout=args.timeout,
            close_after=not args.keep_open,
            probe_paragraphs=probe_paragraphs,
        )
    except RenderError as e:
        print(json.dumps(e.to_dict()))
        return 1

    pdf_path = result["pdf"]
    word_pages = result.get("pages")

    # Cross-check Word's own count against pdfinfo/regex (never fatal —
    # Word's count is authoritative per the module docstring; a mismatch
    # is surfaced as a warning on stderr, not a failure).
    cross_check_pages, cross_check_source = page_count(pdf_path)
    if (
        word_pages is not None
        and cross_check_pages is not None
        and word_pages != cross_check_pages
    ):
        print(
            f"docx_render.py: WARNING page-count mismatch — Word reported "
            f"{word_pages}, {cross_check_source} reported {cross_check_pages}",
            file=sys.stderr,
        )

    if word_pages is not None:
        pages, source = word_pages, "word"
    else:
        pages, source = cross_check_pages, cross_check_source

    left_open = result.get("left_open_document")
    closed_after = result.get("closed_after", False)
    close_error = result.get("close_error")
    # Prints only when a document was ACTUALLY left open (closed_after
    # False) — not on every render, now that closing is the default.
    if left_open:
        note = f'docx_render.py: NOTE Word document "{left_open}" was left open'
        if close_error:
            note += f" (close failed: {close_error})"
        else:
            note += " (--keep-open was passed)"
        note += ". Safe to close by hand; it carries no content beyond the rendered input."
        print(note, file=sys.stderr)

    geometry = result.get("paragraph_geometry") or {}
    line = {
        "engine": "word",
        "pdf": pdf_path,
        "sha256": _sha256(pdf_path),
        "page_count": pages,
        "page_count_source": source,
        "left_open_document": left_open,
        "closed_after": closed_after,
        "close_error": close_error,
        "page_height_pt": result.get("page_height_pt"),
        "paragraph_geometry": {str(k): v for k, v in geometry.items()},
    }
    print(json.dumps(line))
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
