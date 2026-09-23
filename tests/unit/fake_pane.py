"""Fake Word task pane (issue #106 WP-2:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

A Python `websockets` client that speaks exactly the wire protocol
`addin/taskpane.js` speaks (`live/protocol.py`), against an in-memory
"document" (a string body plus a comment list) instead of a real Word
document. This is the acceptance harness for WP-2: nothing here launches
Word or AppleScript (forbidden by the WP-2 hard rules), and every test in
`tests/unit/test_live_session.py` / `tests/unit/test_live_bridge.py` that
needs a connected pane uses this instead.

Fidelity notes (what this fake does NOT emulate, and why that is fine for
WP-2's purpose of testing the transport/session plumbing, not Word's own
editing engine):

  - `replace`/`format`'s optional `track_changes` toggles
    `FakeDocument.change_tracking_mode` for the duration of the op (like
    the real pane's `document.changeTrackingMode` assignment) but does
    not synthesize `w:ins`/`w:del`-equivalent revision marks -- there is
    no OOXML here at all, just a Python string.
  - `format`'s bold/italic/underline/strike/color payload fields are
    accepted and validated but not stored against a run model (this fake
    has no run/style model); the op still enforces the same
    search-and-count-gate and returns the same result shape
    (`ReplaceResult`-shaped: applied, match_count, matches, pre, post)
    the real pane does. `colorAfter`/`strikeAfter` (issue #22: the real
    pane's read-back-after-sync fields, so a caller can tell a write that
    didn't take from one that did) are present on every match for shape
    parity, but this fake ECHOES the request rather than reading anything
    back -- it cannot prove the real Office JS `font.color`/
    `.strikeThrough` read-back path works, only that the server correctly
    checks whatever the pane sends. A real pane sideload (docs/live-mode.md)
    is the only thing that exercises the actual Word object-model path.
  - A comment's `anchor_text` is captured once, at `comment_add` time,
    rather than tracked live against a `Word.Range` that could shift as
    later edits land -- good enough for the WP-2 fixture scenarios, which
    never edit a document after commenting on it in the same test.

Op refusal: `MATCH_COUNT_MISMATCH`-style refusals from `replace`/
`format`/`comment_add` (the actual count of `find` doesn't equal
`expected_matches`) are reported as `LIVE_OP_FAILED` op replies (`ok:
false`), never a protocol error and never a raised exception on the
pane's own receive loop -- matching `live/protocol.py`'s documented
op contract.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

import websockets
from websockets.exceptions import ConnectionClosed

from verified_docx_mcp.live.protocol import OP_ERROR_MATCH_COUNT_MISMATCH, OP_ERROR_ZERO_MATCH


class OpRefused(Exception):
    """Raised by a `FakeDocument` op handler to make `FakePane` reply
    `ok=false` with this code/message, instead of a raised exception
    ever reaching the WSS receive loop (which would look like a pane
    crash, not a normal refusal)."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()


def _find_all(text: str, find: str) -> list[tuple[int, int]]:
    """Non-overlapping occurrences of `find` in `text`, left to right --
    the same greedy scan `str.replace` uses internally."""
    positions: list[tuple[int, int]] = []
    start = 0
    while True:
        idx = text.find(find, start)
        if idx == -1:
            break
        positions.append((idx, idx + len(find)))
        start = idx + len(find)
    return positions


@dataclass
class FakeReply:
    id: str
    content: str
    author_name: str
    creation_date: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "content": self.content,
            "authorName": self.author_name,
            "creationDate": self.creation_date,
        }


@dataclass
class FakeComment:
    id: str
    content: str
    author_name: str
    creation_date: str
    anchor_text: str
    resolved: bool = False
    replies: list[FakeReply] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "content": self.content,
            "authorName": self.author_name,
            "creationDate": self.creation_date,
            "resolved": self.resolved,
            "anchorText": self.anchor_text,
            "replies": [r.to_dict() for r in self.replies],
        }


class FakeDocument:
    """The in-memory "document" every op below reads and writes."""

    def __init__(self, text: str = "", *, author_name: str = "Fake Pane") -> None:
        self.text = text
        self.author_name = author_name
        self.comments: list[FakeComment] = []
        self.change_tracking_mode = "Off"
        self.saved = True
        self._next_comment_id = 1
        self._next_reply_seq: dict[str, int] = {}
        # Comment ids that refuse to flip `resolved` when comment_resolve is
        # called -- issue #106 WP-4's fake-pane knob for exercising
        # COMMENT_STILL_OPEN over the live path the same way file mode's own
        # independent post-write re-read can catch a resolve that did not
        # durably stick. Empty by default (every existing test's behavior is
        # unchanged); a test opts a specific comment id in.
        self.stuck_ids: set[str] = set()

    def sha256(self) -> str:
        return hashlib.sha256(self.text.encode("utf-8")).hexdigest()

    # -- ops ---------------------------------------------------------

    def describe(self) -> dict[str, Any]:
        return {
            "documentUrl": None,  # filled in by FakePane.dispatch (it owns document_url)
            "bodySha256": self.sha256(),
            "changeTrackingMode": self.change_tracking_mode,
            "saved": self.saved,
        }

    def search(self, find: str, *, match_case: bool = False, match_whole_word: bool = False) -> list[dict[str, Any]]:
        import re

        flags = 0 if match_case else re.IGNORECASE
        pattern = re.escape(find)
        if match_whole_word:
            pattern = rf"\b{pattern}\b"
        matches = list(re.finditer(pattern, self.text, flags))
        results = []
        for i, m in enumerate(matches):
            results.append(
                {
                    "index": i,
                    "text": m.group(0),
                    "contextBefore": self.text[max(0, m.start() - 20) : m.start()],
                    "contextAfter": self.text[m.end() : m.end() + 20],
                }
            )
        return results

    def replace(self, find: str, expected_matches: int, replace: str, *, track_changes: bool = False) -> dict[str, Any]:
        pre = self.sha256()
        positions = _find_all(self.text, find)
        if len(positions) != expected_matches:
            raise OpRefused(
                "LIVE_OP_FAILED",
                f"expected {expected_matches} match(es) for {find!r}, found {len(positions)}",
            )
        previous_mode = self.change_tracking_mode
        if track_changes:
            self.change_tracking_mode = "TrackAll"
        try:
            matches_result = [{"before": find, "after": replace} for _ in positions]
            new_text = self.text
            for start, end in reversed(positions):
                new_text = new_text[:start] + replace + new_text[end:]
            self.text = new_text
        finally:
            self.change_tracking_mode = previous_mode
        return {
            "applied": True,
            "match_count": len(positions),
            "matches": matches_result,
            "pre": pre,
            "post": self.sha256(),
        }

    def format(
        self,
        find: str,
        expected_matches: int,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strike: bool | None = None,
        color: str | None = None,
        track_changes: bool = False,
    ) -> dict[str, Any]:
        pre = self.sha256()
        positions = _find_all(self.text, find)
        if len(positions) != expected_matches:
            raise OpRefused(
                "LIVE_OP_FAILED",
                f"expected {expected_matches} match(es) for {find!r}, found {len(positions)}",
            )
        previous_mode = self.change_tracking_mode
        if track_changes:
            self.change_tracking_mode = "TrackAll"
        try:
            # Formatting never changes body text in this fake (no run
            # model to mutate) -- see the module docstring's fidelity
            # notes. before == after == the matched text itself.
            # issue #22: colorAfter/strikeAfter mirror the real pane's
            # read-back contract in SHAPE (present on every match), but
            # this fake has no run/style model to read back from -- it
            # echoes the request, same limitation the module docstring
            # already names for bold/italic/underline. A test that needs
            # to prove the real Office JS read-back path (not just that
            # the server checks whatever the pane sends) uses a real pane
            # sideload instead -- see docs/live-mode.md.
            matches_result = [
                {"before": find, "after": find, "colorAfter": color, "strikeAfter": strike} for _ in positions
            ]
        finally:
            self.change_tracking_mode = previous_mode
        return {
            "applied": True,
            "match_count": len(positions),
            "matches": matches_result,
            "pre": pre,
            "post": self.sha256(),
        }

    def comments_list(self) -> dict[str, Any]:
        return {"comments": [c.to_dict() for c in self.comments]}

    def comment_add(self, find: str, expected_matches: int, text: str) -> dict[str, Any]:
        pre = self.sha256()
        positions = _find_all(self.text, find)
        if len(positions) != expected_matches:
            # issue #106 WP-4: distinguish "nothing matched" from "the wrong
            # count matched" the same way live/protocol.py's OP_ERROR_*
            # constants name it -- server.py's live comment tools map these
            # two onto ZERO_MATCH/MATCH_COUNT_MISMATCH (errors.py), same as
            # file mode's locate() ladder already does.
            code = OP_ERROR_ZERO_MATCH if not positions else OP_ERROR_MATCH_COUNT_MISMATCH
            raise OpRefused(
                code,
                f"expected {expected_matches} match(es) for {find!r}, found {len(positions)}",
            )
        comment = FakeComment(
            id=f"c{self._next_comment_id}",
            content=text,
            author_name=self.author_name,
            creation_date=_now_iso(),
            anchor_text=find,
        )
        self._next_comment_id += 1
        self.comments.append(comment)
        return {"comment_id": comment.id, "pre": pre, "post": self.sha256()}

    def comment_reply(self, comment_id: str, text: str) -> dict[str, Any]:
        comment = self._find_comment(comment_id)
        seq = self._next_reply_seq.get(comment_id, 0) + 1
        self._next_reply_seq[comment_id] = seq
        reply = FakeReply(id=f"{comment_id}-r{seq}", content=text, author_name=self.author_name, creation_date=_now_iso())
        comment.replies.append(reply)
        return {"reply_id": reply.id}

    def comment_resolve(self, comment_id: str, resolved: bool) -> dict[str, Any]:
        comment = self._find_comment(comment_id)
        if comment_id not in self.stuck_ids:
            comment.resolved = resolved
        # A "stuck" comment still replies ok=true (this is not an op
        # refusal -- Word's own comment.resolved assignment has no
        # equivalent of a Drive-API write failure) but reports its
        # UNCHANGED state, so a caller that independently re-lists to
        # confirm (rather than trusting this reply) catches it -- see
        # comments_live.py's COMMENT_STILL_OPEN path.
        return {"resolved": comment.resolved}

    def save(self) -> dict[str, Any]:
        self.saved = True
        return {"saved": True}

    def _find_comment(self, comment_id: str) -> FakeComment:
        for c in self.comments:
            if c.id == comment_id:
                return c
        raise OpRefused("LIVE_OP_FAILED", f"no comment with id {comment_id!r}")


class FakePane:
    """Connects to a running bridge's WSS `/ops` endpoint and answers
    every op against a `FakeDocument`, mimicking `addin/taskpane.js`'s
    dispatcher closely enough to exercise `live/bridge.py` and
    `live/session.py` with no Word installation.
    """

    def __init__(
        self,
        *,
        document: FakeDocument | None = None,
        document_url: str = "/tmp/FakeDoc.docx",
        host: str = "Word",
        platform: str = "Mac",
        requirement_sets: dict[str, Any] | None = None,
        heartbeat_interval: float = 5.0,
    ) -> None:
        self.document = document if document is not None else FakeDocument()
        self.document_url = document_url
        self.host = host
        self.platform = platform
        self.requirement_sets = requirement_sets or {"1.4": True, "1.5": True, "1.6": True}
        self.heartbeat_interval = heartbeat_interval
        # Ops named here are received but never answered -- lets a test
        # simulate an unresponsive pane (for LIVE_DISCONNECTED-by-timeout)
        # without needing a full socket-level failure injection.
        self.drop_ops: set[str] = set()
        self._ws: Any = None
        self._recv_task: asyncio.Task | None = None
        self._heartbeat_task: asyncio.Task | None = None
        self.closed = asyncio.Event()

    async def connect(self, ops_url: str, *, ssl_context=None) -> None:
        self._ws = await websockets.connect(ops_url, ssl=ssl_context)
        await self._ws.send(
            json.dumps(
                {
                    "type": "hello",
                    "documentUrl": self.document_url,
                    "host": self.host,
                    "platform": self.platform,
                    "requirementSets": self.requirement_sets,
                    "bodySha256": self.document.sha256(),
                }
            )
        )
        self._recv_task = asyncio.create_task(self._recv_loop())
        self._heartbeat_task = asyncio.create_task(self._heartbeat_loop())

    async def send_heartbeat(self) -> None:
        """Send one heartbeat immediately -- a test helper, separate from
        the automatic `heartbeat_interval` loop, for tests that want
        precise control over heartbeat timing."""
        await self._ws.send(
            json.dumps({"type": "heartbeat", "documentUrl": self.document_url, "bodySha256": self.document.sha256()})
        )

    async def close(self) -> None:
        for task in (self._heartbeat_task, self._recv_task):
            if task is not None:
                task.cancel()
        if self._ws is not None:
            await self._ws.close()
        self.closed.set()

    async def _heartbeat_loop(self) -> None:
        try:
            while True:
                await asyncio.sleep(self.heartbeat_interval)
                await self.send_heartbeat()
        except (asyncio.CancelledError, ConnectionClosed):
            pass

    async def _recv_loop(self) -> None:
        try:
            async for raw in self._ws:
                message = json.loads(raw)
                if message.get("type") != "request":
                    continue
                request_id = message["request_id"]
                op = message["op"]
                if op in self.drop_ops:
                    continue
                reply = self._handle_op(request_id, op, message.get("payload") or {})
                await self._ws.send(json.dumps(reply))
        except (asyncio.CancelledError, ConnectionClosed):
            pass
        finally:
            self.closed.set()

    def _handle_op(self, request_id: str, op: str, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            result = self._dispatch(op, payload)
        except OpRefused as exc:
            return {"type": "reply", "request_id": request_id, "ok": False, "error": {"code": exc.code, "message": exc.message}}
        return {"type": "reply", "request_id": request_id, "ok": True, "result": result}

    def _dispatch(self, op: str, payload: dict[str, Any]) -> dict[str, Any]:
        doc = self.document
        if op == "ping":
            return {}
        if op == "describe":
            result = doc.describe()
            result["documentUrl"] = self.document_url
            return result
        if op == "search":
            matches = doc.search(
                payload["find"],
                match_case=bool(payload.get("matchCase", False)),
                match_whole_word=bool(payload.get("matchWholeWord", False)),
            )
            return {"matches": matches}
        if op == "replace":
            return doc.replace(
                payload["find"],
                int(payload["expected_matches"]),
                payload["replace"],
                track_changes=bool(payload.get("track_changes", False)),
            )
        if op == "format":
            return doc.format(
                payload["find"],
                int(payload["expected_matches"]),
                bold=payload.get("bold"),
                italic=payload.get("italic"),
                underline=payload.get("underline"),
                strike=payload.get("strike"),
                color=payload.get("color"),
                track_changes=bool(payload.get("track_changes", False)),
            )
        if op == "comments_list":
            return doc.comments_list()
        if op == "comment_add":
            return doc.comment_add(payload["find"], int(payload["expected_matches"]), payload["text"])
        if op == "comment_reply":
            return doc.comment_reply(payload["comment_id"], payload["text"])
        if op == "comment_resolve":
            return doc.comment_resolve(payload["comment_id"], bool(payload["resolved"]))
        if op == "save":
            return doc.save()
        raise OpRefused("LIVE_OP_FAILED", f"fake pane does not implement op {op!r}")
