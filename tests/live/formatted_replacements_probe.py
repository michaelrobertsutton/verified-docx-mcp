"""Disposable real-Word #97 probe. Create with --create, open pane, then run.

VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp PYTHONPATH=src python this_file
"""
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

PATH = Path('/private/tmp/word-formatted-matrix-v3.docx')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ET.register_namespace('w', W[1:-1])
FONT = dict(bold=False, italic=False, underline='None', strikeThrough=False,
            doubleStrikeThrough=False, subscript=False, superscript=False,
            color='#000000', name='Calibri', size=11)


def segment(text, **font):
    return dict(text=text, format={**FONT, **font})


CASES = [
    ('Mixed label: old body', 'New label: plain body', [segment('New label:', bold=True), segment(' plain body')]),
    ('Italic old body', 'Italic new body', [segment('Italic', italic=True), segment(' new body')]),
    ('Link old body', 'Link new body', [{**segment('Link'), 'hyperlink': 'https://example.org/acceptance'}, segment(' new body')]),
    ('Unicode α😀 old', 'Unicode β😀 new', [segment('Unicode β😀', bold=True), segment(' new')]),
    ('Delete mixed', '', []),
    ('Within run', 'Within new', [segment('Within new', italic=True)]),
    ('Cross boundary', 'Cross new', [segment('Cross', bold=True), segment(' new')]),
    ('Soft\vline old', 'Soft\vline new', [segment('Soft\vline', italic=True), segment(' new')]),
]


def create():
    source = Path(__file__).resolve().parents[1] / 'fixtures/approved-changes.docx'
    with ZipFile(source) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        for element in doc.iter():
            element.attrib.pop('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', None)
        cells = doc.findall('.//' + W + 'tr')
        for row, (find, _, _) in zip(cells, CASES):
            cell = row.findall(W + 'tc')[0]
            p = cell.find(W + 'p')
            p.clear()
            for text, bold, italic in [('prefix ', False, True), (find[:5], True, False), (find[5:], False, False), (' suffix', True, False)]:
                r=ET.SubElement(p,W+'r');props=ET.SubElement(r,W+'rPr')
                if bold: ET.SubElement(props,W+'b')
                if italic: ET.SubElement(props,W+'i')
                for i, part in enumerate(text.split('\v')):
                    if i: ET.SubElement(r,W+'br')
                    t=ET.SubElement(r,W+'t', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'});t.text=part
        with ZipFile(PATH, 'w') as out:
            for entry in z.infolist(): out.writestr(entry, ET.tostring(doc, encoding='utf-8', xml_declaration=True) if entry.filename=='word/document.xml' else z.read(entry.filename))
    print(PATH)


def run():
    from verified_docx_mcp.live import bridge, approved_changes as runs
    bridge.start_in_background()
    ops=[dict(op='replace_text',section='target',args=dict(find=find, replace=replace,expected_matches=1,replacement_segments=segments)) for find,replace,segments in CASES]
    ops.append(dict(op='live_save',section='',args={}))
    p=runs.prepare(str(PATH),ops)
    result=runs.execute(str(PATH),p['run_id'],p['manifest_sha256'],'User-approved #97 disposable live acceptance')
    evidence=dict(prepared=p,result=result)
    Path('/private/tmp/word-formatted-matrix-evidence.json').write_text(json.dumps(evidence,indent=2))
    print(json.dumps(result),flush=True)
    assert result['status']=='complete',result


if __name__=='__main__':
    create() if '--create' in sys.argv else run()
