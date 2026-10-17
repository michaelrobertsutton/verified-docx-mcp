"""Follow-up live #97 acceptance on the completed formatted matrix fixture."""
import json
from pathlib import Path

from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, write_mode as wm

PATH='/private/tmp/word-formatted-matrix-v3.docx'
FONT=dict(bold=False,italic=False,underline='None',strikeThrough=False,doubleStrikeThrough=False,subscript=False,superscript=False,color='#000000',name='Calibri',size=11)
bridge.start_in_background()
pane=wm.live_session_for(PATH)
results=[]

def edit(section,find,replace,segments=None):
    args=dict(find=find,replace=replace,expected_matches=1)
    if segments is not None: args['replacement_segments']=segments
    return dict(op='replace_text',section=section,args=args)


def segments(label,body):
    return [dict(text=label,format={**FONT,'bold':True}),dict(text=body,format=FONT)]


def execute(ops,tracked=False):
    p=runs.prepare(PATH,ops+[dict(op='live_save',section='',args={})],track_changes=tracked)
    result=runs.execute(PATH,p['run_id'],p['manifest_sha256'],'User-approved disposable #97 semantics acceptance')
    results.append(dict(manifest=p['manifest_path'],report=result))
    print(json.dumps(result),flush=True)
    return result

# The same exact text elsewhere must not change the selected target.
setup_scope=pane.request_threadsafe('execution_snapshot',{'section_slugs':['other']})['sections']
inserted=pane.request_threadsafe('paragraph_insert',dict(anchor='Other',expected_matches=1,position='after',paragraphs=[{'text':'New label: plain body'}],executionSections=setup_scope,executionTargetSection='other'))
assert inserted['applied'] is True
assert execute([edit('target','New label: plain body','Repeated label: plain body',segments('Repeated label:',' plain body'))])['status']=='complete'
other=pane.request_threadsafe('execution_snapshot',{'section_slugs':['other']})['sections'][0]
assert 'New label: plain body' in '\n'.join(other['paragraphs'])

comments=pane.request_threadsafe('comments_list',{'include_anchor':True})
assert comments['comments']
# A comment refusal must not stop the independent formatted edit.
r=execute([edit('other','Collaborator comment anchor.','Do not remove the comment.'),edit('target','Italic new body','Italic verified body',[dict(text='Italic',format={**FONT,'italic':True}),dict(text=' verified body',format=FONT)])])
assert r['status']=='partial' and r['applied_count']==2 and r['issues'][0]['code']=='WOULD_DELETE_COMMENTS',r
assert len(pane.request_threadsafe('comments_list',{})['comments'])==len(comments['comments'])

# Create a tracked mixed-format replacement, restore mode, then preserve its
# pending revisions while a separate target still completes.
r=execute([edit('target','Within new','Tracked insertion',[dict(text='Tracked',format={**FONT,'italic':True}),dict(text=' insertion',format=FONT)])],tracked=True)
assert r['status']=='complete',r
revisions=pane.request_threadsafe('revisions_list',{})
assert revisions.get('revisions') or revisions.get('desktop_revisions') or revisions.get('ooxml_revision_count')
r=execute([edit('target','Tracked insertion','Must not replay',segments('Must',' not replay')),edit('target','Cross new','Cross verified',segments('Cross',' verified'))])
assert r['status']=='partial' and r['applied_count']==2 and r['issues'][0]['code']=='TRACKED_CHANGES_PRESENT',r
remaining=pane.request_threadsafe('revisions_list',{})
assert remaining.get('revisions') or remaining.get('desktop_revisions') or remaining.get('ooxml_revision_count')
Path('/private/tmp/word-formatted-semantics-evidence.json').write_text(json.dumps(dict(results=results,repeated_target_preserved=True,comments_preserved=True,revisions_preserved=True),indent=2))
