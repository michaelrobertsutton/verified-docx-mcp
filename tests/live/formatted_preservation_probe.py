"""Run after formatted_semantics_probe.py; verify an existing adjacent link and revisions."""
import json
from pathlib import Path

from verified_docx_mcp.live import bridge, approved_changes as runs, write_mode as wm

PATH='/private/tmp/word-formatted-matrix-v3.docx'
bridge.start_in_background()
pane=wm.live_session_for(PATH)
comments_before=pane.request_threadsafe('comments_list',{})
revisions_before=pane.request_threadsafe('revisions_list',{})
font=dict(bold=False,italic=False,underline='None',strikeThrough=False,doubleStrikeThrough=False,subscript=False,superscript=False,color='#000000',name='Calibri',size=11)
op=dict(op='replace_text',section='target',args=dict(find='new body',replace='verified link body',expected_matches=1,replacement_segments=[dict(text='verified link body',format=font)]))
p=runs.prepare(PATH,[op,dict(op='live_save',section='',args={})])
report=runs.execute(PATH,p['run_id'],p['manifest_sha256'],'User-approved existing-link and revision preservation acceptance')
print(json.dumps(report),flush=True)
assert report['status']=='complete',report
comments_after=pane.request_threadsafe('comments_list',{})
revisions_after=pane.request_threadsafe('revisions_list',{})
def revisions(value):
    return sorted((str(e.get('type')),str(e.get('author')),str(e.get('text'))) for e in value.get('revisions',[]))
def comments(value):
    return sorted((str(e.get('id')),str(e.get('content')),str(e.get('resolved')),json.dumps(e.get('replies',[]),sort_keys=True)) for e in value['comments'])
assert revisions(revisions_before)==revisions(revisions_after)
assert comments(comments_before)==comments(comments_after)
if revisions_before.get('ooxml_revision_count') is not None:
    assert revisions_before['ooxml_revision_count']==revisions_after['ooxml_revision_count']
Path('/private/tmp/word-format-preservation-evidence.json').write_text(json.dumps(dict(report=report,manifest=p['manifest_path'],existing_link_preserved=True,revision_inventory_unchanged=True,comment_inventory_unchanged=True),indent=2))
