from unittest.mock import patch

import pytest

from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import comments_live

RAW = [{"id": "a", "content": "same", "authorName": "A", "anchorText": "quote",
        "anchorParagraphText": "first"},
       {"id": "b", "content": "same", "authorName": "A", "anchorText": "quote",
        "anchorParagraphText": "second"}]


def address(comment_id=None, spec=None, epoch=None):
    with patch.object(comments_live, "_session_for"), patch.object(
        comments_live, "_request", return_value={"comments": RAW, "session_epoch": "current"}
    ):
        return comments_live.address_comment("doc.docx", comment_id, spec, epoch)


def test_duplicate_text_refuses():
    with pytest.raises(VerifyError):
        address(spec={"text": "same", "author": "A"})


def test_context_disambiguates():
    handle, payload = address(spec={"text": "same", "anchor_paragraph_contains": "second"})
    assert handle == "live:b"
    assert payload["identity"]["anchorParagraphText"] == "second"


def test_stale_epoch_and_conflicting_address_refuse():
    with pytest.raises(VerifyError):
        address("live:old:a")
    with pytest.raises(VerifyError):
        address("live:a", {"text": "same"})
    assert address("live:current:a")[0] == "live:a"


def test_content_only_correlation_never_resolves():
    correlation = [{"confidence": "content-only", "comment_id": "durable", "w_id": "1",
                    "live_comment_id": "live:a"}]
    with pytest.raises(VerifyError):
        comments_live._resolve_comment_handle("durable", correlation)


def test_ambiguous_exact_correlation_refuses():
    correlation = [{"confidence": "exact", "comment_id": "durable", "w_id": "1",
                    "live_comment_id": f"live:{id}"} for id in ("a", "b")]
    with pytest.raises(VerifyError):
        comments_live._resolve_comment_handle("durable", correlation)
