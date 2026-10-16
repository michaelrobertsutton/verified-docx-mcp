"""Disposable #98 acceptance: create, open its Live pane, then run."""
import json
import sys
from pathlib import Path
from zipfile import ZipFile

PATH=Path('/private/tmp/word-callout-acceptance-v3.docx')


def create():
    source=Path(__file__).resolve().parents[1]/'fixtures/textbox.docx'
    with ZipFile(source) as z, ZipFile(PATH,'w') as out:
        for entry in z.infolist():
            content=z.read(entry.filename)
            if entry.filename=='word/document.xml':
                xml=content.decode().replace('<w:body>','<w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Target</w:t></w:r></w:p>',1)
                xml=xml.replace('<w:r><w:t>Body text after the text box.</w:t></w:r>','<w:r><w:rPr><w:b/></w:rPr><w:t>Body text</w:t></w:r><w:r><w:t xml:space="preserve"> after the text box.</w:t></w:r>')
                xml=xml.replace('</mc:AlternateContent></w:r></w:p>','</mc:AlternateContent></w:r><w:r><w:t xml:space="preserve"> Trailing editable text.</w:t></w:r></w:p>',1)
                xml=xml.replace('<w:sectPr', '<w:p><w:r><w:t>Independent final paragraph.</w:t></w:r></w:p><w:sectPr',1)
                content=xml.encode()
            out.writestr(entry,content)
    print(PATH)


def run():
    from verified_docx_mcp.live import bridge,approved_changes as runs,write_mode as wm
    bridge.start_in_background()
    pane=wm.live_session_for(str(PATH))
    assert 'anchored_text_v1' in pane.hello.capabilities
    before=pane.request_threadsafe('shapes_list')['shapes']
    assert before
    font=dict(bold=False,italic=False,underline='None',strikeThrough=False,doubleStrikeThrough=False,subscript=False,superscript=False,color='#000000',name='Calibri',size=11)
    def edit(find,replace,segments=None):
        args=dict(find=find,replace=replace,expected_matches=1)
        if segments is not None: args['replacement_segments']=segments
        return dict(op='replace_text',section='target',args=args)
    ops=[edit('Body text after','New label: plain',[dict(text='New label:',format={**font,'bold':True}),dict(text=' plain',format=font)]),edit('Trailing editable text.','Trailing verified text.'),dict(op='live_save',section='',args={})]
    p=runs.prepare(str(PATH),ops);result=runs.execute(str(PATH),p['run_id'],p['manifest_sha256'],'User-approved #98 disposable acceptance')
    after=pane.request_threadsafe('shapes_list')['shapes']
    def definitions(shapes): return sorted((s['id'],s['definition']) for s in shapes)
    assert definitions(before)==definitions(after)
    assert result['status']=='complete',result
    state=pane.request_threadsafe('execution_snapshot',{'section_slugs':['target']})['sections']
    anchored=next(p for p in state[0]['paragraphs'] if 'New label:' in p)
    refused=[dict(op='delete_paragraph',section='target',args=dict(anchor_text=anchored)),edit('Body text before','Independent text before'),dict(op='live_save',section='',args={})]
    q=runs.prepare(str(PATH),refused);report=runs.execute(str(PATH),q['run_id'],q['manifest_sha256'],'User-approved #98 refusal acceptance')
    assert report['status']=='partial' and report['applied_count']==2,report
    assert report['issues'][0]['code']=='ANCHORED_SHAPES',report
    assert definitions(after)==definitions(pane.request_threadsafe('shapes_list')['shapes'])
    evidence=dict(allowed=result,refused=report,shape_definitions_preserved=True,manifest=p['manifest_path'])
    Path('/private/tmp/word-callout-evidence.json').write_text(json.dumps(evidence,indent=2))
    print(json.dumps(evidence),flush=True)


if __name__=='__main__': create() if '--create' in sys.argv else run()
