import copy

import pytest

from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.formatted_text import validate_segments
from verified_docx_mcp.live import approved_changes as runs
from test_approved_changes import setup  # noqa: F401

FONT = dict(bold=False, italic=False, underline='None', strikeThrough=False,
            doubleStrikeThrough=False, subscript=False, superscript=False,
            color='#000000', name='Calibri', size=11)


def segments():
    return [dict(text='Label:', format={**FONT, 'bold': True}),
            dict(text=' plain body', format=FONT)]


def test_explicit_label_and_body():
    value = validate_segments(segments(), 'Label: plain body')
    assert value[0]['format']['bold'] is True
    assert value[1]['format']['bold'] is False
    assert validate_segments([], '') == []
    assert validate_segments(None, 'legacy') is None


@pytest.mark.parametrize('change', ['text', 'extra', 'color', 'size', 'link', 'break', 'bool'])
def test_reject_invalid_segments(change):
    value = copy.deepcopy(segments())
    if change == 'text': value[0]['text'] = 'Other'
    if change == 'extra': value[0]['format']['style'] = 'Normal'
    if change == 'color': value[0]['format']['color'] = 'auto'
    if change == 'size': value[0]['format']['size'] = float('nan')
    if change == 'link': value[0]['hyperlink'] = 'javascript:alert(1)'
    if change == 'break': value[0]['text'] = 'Label:\n'
    if change == 'bool': value[0]['format']['bold'] = 1
    with pytest.raises(VerifyError): validate_segments(value, 'Label: plain body')


def test_unicode_links_and_soft_break():
    value = [dict(text='α😀\vlink', format={**FONT, 'italic': True}, hyperlink='https://example.org')]
    assert validate_segments(value, 'α😀\vlink')[0]['hyperlink'] == 'https://example.org'


def test_manifest_preview_binds_segments(setup):
    import json
    from pathlib import Path
    path, pane = setup
    pane.hello = type(pane.hello)(capabilities=[*pane.hello.capabilities, 'formatted_replacements_v1'])
    op = dict(op='replace_text', section='target', args=dict(find='Proof is current.', replace='Label: plain body', expected_matches=1, replacement_segments=segments()))
    prepared = runs.prepare(path, [op])
    manifest = json.loads(Path(prepared['manifest_path']).read_text())
    assert prepared['preview'][0]['replacement_segments'] == segments()
    assert manifest['operations'][0]['args']['replacement_segments'] == segments()
    manifest['operations'][0]['args']['replacement_segments'][1]['format']['bold'] = True
    assert runs.sha(manifest) != prepared['manifest_sha256']


def test_old_pane_refuses_before_manifest(setup):
    path, _ = setup
    op = dict(op='replace_text', section='target', args=dict(find='Proof', replace='Label: plain body', expected_matches=1, replacement_segments=segments()))
    with pytest.raises(VerifyError) as exc: runs.prepare(path, [op])
    assert exc.value.envelope.error_code.value == 'LIVE_CAPABILITY_MISSING'


def test_pane_formatting_harness():
    import subprocess
    from pathlib import Path
    root = Path(__file__).resolve().parents[2]
    subprocess.run(['node', str(root/'tests/unit/js/formatted_harness.mjs'), str(root/'addin/taskpane.js')], check=True, capture_output=True, text=True)


def test_partial_font_fields_inherit_from_baseline():
    from verified_docx_mcp.formatted_text import resolve_segments
    value=resolve_segments([{'text':'bold','format':{'bold':True}},{'text':' plain'}], 'bold plain', FONT)
    assert value[0]['format']=={**FONT,'bold':True}
    assert value[1]['format']==FONT
    assert validate_segments([{'text':'x','format':{}}], 'x')[0]['format']=={}
