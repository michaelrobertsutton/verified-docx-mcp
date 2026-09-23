"""Unit tests for src/verified_docx_mcp/live/protocol.py (issue #106 WP-2:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

Schema round trips (`to_json()` -> `from_json()` reproduces the original
dataclass) and validation errors (`from_json` raises `ProtocolError`, never
a bare `KeyError`/`TypeError`, on a malformed message)."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp.live.protocol import (
    VALID_OPS,
    CommentAddPayload,
    CommentInfo,
    CommentReplyInfo,
    CommentReplyPayload,
    CommentResolvePayload,
    DescribeResult,
    FormatPayload,
    HeartbeatMessage,
    HelloMessage,
    OpError,
    OpReply,
    OpRequest,
    ProtocolError,
    ReplaceMatchResult,
    ReplacePayload,
    SearchMatch,
    SearchPayload,
    peek_type,
)


class HelloMessageTests(unittest.TestCase):
    def test_round_trip(self):
        msg = HelloMessage(
            document_url="/Users/lead/Desktop/Report.docx",
            host="Word",
            platform="Mac",
            requirement_sets={"1.4": True, "1.5": True, "1.6": True},
            body_sha256="a" * 64,
        )
        self.assertEqual(HelloMessage.from_json(msg.to_json()), msg)
        # also round-trips through a JSON string, not just a dict
        import json

        self.assertEqual(HelloMessage.from_json(json.dumps(msg.to_json())), msg)

    def test_missing_field_raises_protocol_error(self):
        obj = {"type": "hello", "documentUrl": "/x.docx", "host": "Word", "platform": "Mac"}
        with self.assertRaises(ProtocolError):
            HelloMessage.from_json(obj)

    def test_wrong_type_raises_protocol_error(self):
        obj = {
            "type": "hello",
            "documentUrl": "/x.docx",
            "host": "Word",
            "platform": "Mac",
            "requirementSets": "not-a-dict",
            "bodySha256": "abc",
        }
        with self.assertRaises(ProtocolError):
            HelloMessage.from_json(obj)

    def test_wrong_message_type_raises_protocol_error(self):
        with self.assertRaises(ProtocolError):
            HelloMessage.from_json({"type": "heartbeat"})

    def test_not_json_raises_protocol_error(self):
        with self.assertRaises(ProtocolError):
            HelloMessage.from_json("{not json")

    def test_top_level_not_object_raises_protocol_error(self):
        with self.assertRaises(ProtocolError):
            HelloMessage.from_json("[1, 2, 3]")


class HeartbeatMessageTests(unittest.TestCase):
    def test_round_trip(self):
        msg = HeartbeatMessage(document_url="/x.docx", body_sha256="b" * 64)
        self.assertEqual(HeartbeatMessage.from_json(msg.to_json()), msg)

    def test_missing_field_raises(self):
        with self.assertRaises(ProtocolError):
            HeartbeatMessage.from_json({"type": "heartbeat", "documentUrl": "/x.docx"})


class OpRequestReplyTests(unittest.TestCase):
    def test_op_request_round_trip(self):
        req = OpRequest(request_id="r1", op="ping", payload={})
        self.assertEqual(OpRequest.from_json(req.to_json()), req)

    def test_op_request_rejects_unknown_op(self):
        with self.assertRaises(ProtocolError):
            OpRequest(request_id="r1", op="not-a-real-op")

    def test_op_request_from_json_defaults_missing_payload_to_empty_dict(self):
        req = OpRequest.from_json({"type": "request", "request_id": "r1", "op": "save"})
        self.assertEqual(req.payload, {})

    def test_op_request_from_json_rejects_wrong_type(self):
        with self.assertRaises(ProtocolError):
            OpRequest.from_json({"type": "reply", "request_id": "r1", "op": "ping"})

    def test_op_reply_ok_round_trip(self):
        reply = OpReply(request_id="r1", ok=True, result={"pong": True})
        self.assertEqual(OpReply.from_json(reply.to_json()), reply)

    def test_op_reply_error_round_trip(self):
        reply = OpReply(request_id="r1", ok=False, error=OpError(code="LIVE_OP_FAILED", message="nope"))
        round_tripped = OpReply.from_json(reply.to_json())
        self.assertEqual(round_tripped, reply)

    def test_op_reply_ok_false_without_error_is_rejected(self):
        with self.assertRaises(ProtocolError):
            OpReply.from_json({"type": "reply", "request_id": "r1", "ok": False})

    def test_op_reply_wrong_message_type_rejected(self):
        with self.assertRaises(ProtocolError):
            OpReply.from_json({"type": "request", "request_id": "r1", "ok": True})

    def test_all_ops_are_valid_for_op_request(self):
        for op in VALID_OPS:
            OpRequest(request_id="r", op=op)  # must not raise


class PeekTypeTests(unittest.TestCase):
    def test_peek_type_from_dict_and_string(self):
        self.assertEqual(peek_type({"type": "hello"}), "hello")
        self.assertEqual(peek_type('{"type": "reply"}'), "reply")

    def test_peek_type_missing_raises(self):
        with self.assertRaises(ProtocolError):
            peek_type({"no_type_field": True})


class SearchPayloadTests(unittest.TestCase):
    def test_round_trip_with_defaults(self):
        payload = SearchPayload(find="alpha")
        obj = payload.to_json()
        self.assertEqual(obj["matchCase"], False)
        self.assertEqual(obj["matchWholeWord"], False)
        self.assertEqual(SearchPayload.from_json(obj), payload)

    def test_round_trip_with_all_fields(self):
        payload = SearchPayload(find="alpha", match_case=True, match_whole_word=True, expected_body_sha256="c" * 64)
        self.assertEqual(SearchPayload.from_json(payload.to_json()), payload)

    def test_empty_find_rejected(self):
        with self.assertRaises(ProtocolError):
            SearchPayload.from_json({"find": ""})

    def test_missing_find_rejected(self):
        with self.assertRaises(ProtocolError):
            SearchPayload.from_json({})


class SearchMatchTests(unittest.TestCase):
    def test_from_json(self):
        m = SearchMatch.from_json(
            {"index": 0, "text": "alpha", "contextBefore": "x ", "contextAfter": " y"}
        )
        self.assertEqual(m, SearchMatch(index=0, text="alpha", context_before="x ", context_after=" y"))

    def test_context_fields_default_to_empty_string(self):
        m = SearchMatch.from_json({"index": 0, "text": "alpha"})
        self.assertEqual(m.context_before, "")
        self.assertEqual(m.context_after, "")


class ReplacePayloadTests(unittest.TestCase):
    def test_round_trip(self):
        payload = ReplacePayload(find="alpha", expected_matches=2, replace="beta", track_changes=True)
        self.assertEqual(ReplacePayload.from_json(payload.to_json()), payload)

    def test_expected_matches_must_be_positive(self):
        with self.assertRaises(ProtocolError):
            ReplacePayload.from_json({"find": "a", "expected_matches": 0, "replace": "b"})

    def test_empty_find_rejected(self):
        with self.assertRaises(ProtocolError):
            ReplacePayload.from_json({"find": "", "expected_matches": 1, "replace": "b"})

    def test_expected_body_sha256_round_trips_when_present(self):
        payload = ReplacePayload(find="a", expected_matches=1, replace="b", expected_body_sha256="d" * 64)
        obj = payload.to_json()
        self.assertEqual(obj["expectedBodySha256"], "d" * 64)
        self.assertEqual(ReplacePayload.from_json(obj).expected_body_sha256, "d" * 64)

    def test_expected_body_sha256_omitted_when_none(self):
        payload = ReplacePayload(find="a", expected_matches=1, replace="b")
        self.assertNotIn("expectedBodySha256", payload.to_json())


class FormatPayloadTests(unittest.TestCase):
    def test_round_trip(self):
        payload = FormatPayload(find="alpha", expected_matches=1, bold=True, italic=False, track_changes=True)
        self.assertEqual(FormatPayload.from_json(payload.to_json()), payload)

    def test_requires_at_least_one_toggle(self):
        with self.assertRaises(ProtocolError):
            FormatPayload.from_json({"find": "a", "expected_matches": 1})

    def test_expected_matches_must_be_positive(self):
        with self.assertRaises(ProtocolError):
            FormatPayload.from_json({"find": "a", "expected_matches": 0, "bold": True})

    def test_strike_only_round_trips(self):
        # Issue #22: live format_text used to silently drop strike -- it
        # must round-trip on its own, without bold/italic/underline/color.
        payload = FormatPayload(find="alpha", expected_matches=1, strike=True)
        self.assertEqual(FormatPayload.from_json(payload.to_json()), payload)

    def test_color_only_round_trips(self):
        payload = FormatPayload(find="alpha", expected_matches=1, color="3B3838")
        self.assertEqual(FormatPayload.from_json(payload.to_json()), payload)

    def test_color_alone_satisfies_the_at_least_one_requirement(self):
        parsed = FormatPayload.from_json({"find": "a", "expected_matches": 1, "color": "3B3838"})
        self.assertEqual(parsed.color, "3B3838")


class ReplaceMatchResultTests(unittest.TestCase):
    def test_from_json(self):
        r = ReplaceMatchResult.from_json({"before": "a", "after": "b"})
        self.assertEqual(r, ReplaceMatchResult(before="a", after="b"))


class CommentTests(unittest.TestCase):
    def test_comment_reply_info_from_json(self):
        r = CommentReplyInfo.from_json(
            {"id": "c1-r1", "content": "thanks", "authorName": "A. Reviewer", "creationDate": "2026-01-01"}
        )
        self.assertEqual(r.id, "c1-r1")

    def test_comment_info_from_json_with_replies(self):
        info = CommentInfo.from_json(
            {
                "id": "c1",
                "content": "please fix",
                "authorName": "A. Reviewer",
                "creationDate": "2026-01-01",
                "resolved": False,
                "anchorText": "the target text",
                "replies": [
                    {"id": "c1-r1", "content": "done", "authorName": "Author", "creationDate": "2026-01-02"}
                ],
            }
        )
        self.assertEqual(len(info.replies), 1)
        self.assertEqual(info.replies[0].content, "done")

    def test_comment_info_defaults_replies_to_empty_list(self):
        info = CommentInfo.from_json(
            {
                "id": "c1",
                "content": "x",
                "authorName": "A",
                "creationDate": "2026-01-01",
                "resolved": True,
                "anchorText": "x",
            }
        )
        self.assertEqual(info.replies, [])

    def test_comment_add_payload_round_trip(self):
        payload = CommentAddPayload(find="alpha", expected_matches=1, text="note")
        self.assertEqual(CommentAddPayload.from_json(payload.to_json()), payload)

    def test_comment_add_payload_rejects_empty_text(self):
        with self.assertRaises(ProtocolError):
            CommentAddPayload.from_json({"find": "a", "expected_matches": 1, "text": ""})

    def test_comment_reply_payload_round_trip(self):
        payload = CommentReplyPayload(comment_id="c1", text="thanks")
        self.assertEqual(CommentReplyPayload.from_json(payload.to_json()), payload)

    def test_comment_reply_payload_rejects_empty_comment_id(self):
        with self.assertRaises(ProtocolError):
            CommentReplyPayload.from_json({"comment_id": "", "text": "x"})

    def test_comment_resolve_payload_round_trip(self):
        payload = CommentResolvePayload(comment_id="c1", resolved=True)
        self.assertEqual(CommentResolvePayload.from_json(payload.to_json()), payload)


class DescribeResultTests(unittest.TestCase):
    def test_from_json(self):
        result = DescribeResult.from_json(
            {
                "documentUrl": "/x.docx",
                "bodySha256": "e" * 64,
                "changeTrackingMode": "Off",
                "saved": True,
            }
        )
        self.assertEqual(result.document_url, "/x.docx")
        self.assertTrue(result.saved)


if __name__ == "__main__":
    unittest.main()
