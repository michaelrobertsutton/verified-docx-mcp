"""The session key is the document's file name, derived from Office.context.document.url (#61)."""
import pytest

from verified_docx_mcp.live.session import document_name_from_url

BASE = "https://skywarditsolutions-my.sharepoint.com/personal/u_skyward_com/Documents/folder"


@pytest.mark.parametrize("url, expected", [
    # Captured from Word for Mac 16.113.3: a hosted file with a space, an accent and a '#'
    # is reported with its name RAW; '#' must not be read as a URL fragment.
    (f"{BASE}/Accépt test 61 #1.docx", "Accépt test 61 #1.docx"),
    # The same file reported percent-encoded.
    (f"{BASE}/Acc%C3%A9pt%20test%2061%20%231.docx", "Accépt test 61 #1.docx"),
    (f"{BASE}/plain.docx", "plain.docx"),
    (f"{BASE}/plain.docx?web=1", "plain.docx"),
    ("file:///Users/me/Doc%20one.docx", "Doc one.docx"),
    ("/Users/me/Plan #2.docx", "Plan #2.docx"),
    ("C:\\Users\\me\\Plan #3.docx", "Plan #3.docx"),
])
def test_document_name_from_url(url, expected):
    assert document_name_from_url(url) == expected
