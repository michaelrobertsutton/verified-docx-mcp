"""Manual live acceptance probe; requires the disposable fixture pane.

Run with VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp from the checkout.
Never points at a production document. See docs/acceptance/approved-changes.md.
"""
import asyncio
import json
import time
from collections import Counter
from pathlib import Path

from fastmcp import Client

from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, shared
from verified_docx_mcp.server import mcp

bridge.start_in_background()
path='/private/tmp/jennystack-word-acceptance.docx'
main_counts=Counter();peer_counts=Counter();peer_mode=False
native=shared.RemoteSession.request_threadsafe
def counted(self,op,*args,**kw):
 (peer_counts if peer_mode else main_counts)[op]+=1
 return native(self,op,*args,**kw)
shared.RemoteSession.request_threadsafe=counted
ops=[{'op':'replace_cell_markdown','section':'target','args':{'table_id':1,'row_index':r,'cell_index':c,'markdown':f'**Accepted {r}:{c}**'}} for r in range(1,13) for c in range(1,4)]+[{'op':'live_save','section':'','args':{}}]
async def prepare():
 async with Client(mcp,timeout=1800) as client:
  return (await client.call_tool('live_changes_prepare',{'path':path,'operations':ops})).data
prepared=asyncio.run(prepare())
manifest=json.loads(Path(prepared['manifest_path']).read_text())
assert manifest['comments']['comments'][0]['resolved'] is False
peer=shared.Client(shared.runtime_dir()/'broker.sock');pane=peer.get('jennystack-word-acceptance.docx')
apply=runs.apply;done=0;outside_edits=0;resolved=False

def interleaved(*args,**kw):
 global done,peer_mode,outside_edits,resolved
 result=apply(*args,**kw)
 if args[1]['op']!='live_save':
  done+=1
  if done%6==0:
   peer_mode=True
   try:
    sections=pane.request_threadsafe('execution_snapshot',{'section_slugs':['other']})['sections']
    heading=sections[0]['paragraphs'][0]
    replacement='OTHER' if heading=='Other' else 'Other'
    report=pane.request_threadsafe('replace',{'find':heading,'replace':replacement,'expected_matches':1,'inherit_format':'replaced','scope':'body','allow_comment_loss':False,'executionSections':sections,'executionTargetSection':'other'})
    assert report['applied'] is True
    outside_edits+=1
    if done==18:
     sections[0]['paragraphs'][0]=replacement
     comment=manifest['comments']['comments'][0]
     result2=pane.request_threadsafe('comment_resolve',{'comment_id':comment['id'],'resolved':True,'session_epoch':manifest['comments']['session_epoch'],'identity':{k:comment[k] for k in ('content','authorName','anchorText','anchorParagraphText')},'executionSections':sections,'executionTargetSection':'other'})
     assert result2['resolved'] is True
     resolved=True
   finally:peer_mode=False
 return result
runs.apply=interleaved
start=time.monotonic()
try:
 async def execute():
  async with Client(mcp,timeout=1800) as client:
   return (await client.call_tool('live_changes_execute',{'path':path,'run_id':prepared['run_id'],'manifest_sha256':prepared['manifest_sha256'],'approved_by':'User-approved live acceptance'})).data
 result=asyncio.run(execute())
 evidence={'prepared':prepared,'report':result,'main_calls':dict(main_counts),'collaborator_calls':dict(peer_counts),'outside_edits':outside_edits,'comment_resolved':resolved,'execution_requests':1,'elapsed_seconds':round(time.monotonic()-start,2)}
 Path('/private/tmp/word-acceptance-clean.json').write_text(json.dumps(evidence,indent=2))
 print(json.dumps({k:v for k,v in evidence.items() if k!='prepared'}),flush=True)
 assert result['status']=='complete',result
 assert result['applied_count']==37
 assert main_counts['execution_snapshot']==2
 assert main_counts['comments_list']==1
 assert main_counts['body_ooxml']==main_counts['describe']==0
 assert outside_edits==6 and resolved
finally:peer.close()
