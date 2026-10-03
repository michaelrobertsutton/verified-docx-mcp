import shutil
import subprocess
from pathlib import Path

import pytest

from verified_docx_mcp.live.write_mode import live_evidence


@pytest.mark.skipif(not shutil.which('node'), reason='Node unavailable')
def test_shape_inventory_and_guard():
    repo = Path(__file__).resolve().parents[2]
    subprocess.run(['node', str(repo / 'tests/unit/js/shapes_harness.mjs'),
                    str(repo / 'addin/taskpane.js')], check=True, capture_output=True, text=True)


def test_shape_counts_are_returned_and_audited(monkeypatch):
    from verified_docx_mcp.live import write_mode
    captured = []
    monkeypatch.setattr(write_mode.audit, 'append_audit',
                        lambda **kw: (captured.append(kw['evidence'].copy()) or True, None))
    evidence = live_evidence(applied=True, match_count=1, rung=1, before='a', after='b',
                             pre_body_sha256='a', post_body_sha256='b', document_name='x',
                             tool='replace_text', path='x',
                             shape_result={'shapes_before':3,'shapes_after':3})
    assert evidence['shapes_before'] == evidence['shapes_after'] == 3
    assert captured[0]['shapes_after'] == 3


@pytest.mark.asyncio
async def test_old_pane_refuses_write_before_transport():
    from types import SimpleNamespace

    from verified_docx_mcp.live.session import LiveOpFailed, LiveSession
    session = SimpleNamespace(hello=SimpleNamespace(capabilities=frozenset()))
    with pytest.raises(LiveOpFailed) as error:
        await LiveSession._request_on_loop(session, 'replace', {})
    assert error.value.code == 'LIVE_CAPABILITY_MISSING'
