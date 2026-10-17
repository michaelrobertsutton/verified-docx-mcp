"""Manual live conflict probe; run after approved_changes_probe.py.

Requires the separately opened disposable conflict fixture.
"""
import asyncio
import json
from pathlib import Path

from fastmcp import Client

from verified_docx_mcp import tables
from verified_docx_mcp.live import bridge
from verified_docx_mcp.live import write_mode as wm
from verified_docx_mcp.server import mcp

bridge.start_in_background()
path='/private/tmp/word-stream-conflict.docx'
async def main():
 async with Client(mcp,timeout=300) as client:
  ops=[{'op':'replace_cell_markdown','section':'target','args':{'table_id':1,'row_index':1,'cell_index':c,'markdown':text}} for c,text in ((1,'Must not overwrite collaborator'),(2,'Independent target succeeds'))]+[{'op':'live_save','section':'','args':{}}]
  prepared=(await client.call_tool('live_changes_prepare',{'path':path,'operations':ops})).data
  pane=wm.live_session_for(path)
  sections=pane.request_threadsafe('execution_snapshot',{'section_slugs':['target']})['sections']
  peer_payload={'executionSections':sections,'executionTargetSection':'target','table_index':1,'row_index':1,'cell_index':1,'expected_before_text':sections[0]['tables'][0]['rows'][0][0],'expected_before_paragraph_count':1,'paragraphs':tables._runs_to_wire(tables.markdown_to_ooxml.parse_paragraph_runs('Coauthor retained'))}
  peer=pane.request_threadsafe('cell_set',peer_payload)
  assert peer['applied'] is True
  report=(await client.call_tool('live_changes_execute',{'path':path,'run_id':prepared['run_id'],'manifest_sha256':prepared['manifest_sha256'],'approved_by':'User-approved live conflict acceptance'})).data
  assert report['issues'][0]['status']=='conflict' and report['issues'][0]['code'] in {'CELL_TEXT_MISMATCH','LIVE_STALE','TARGET_CHANGED'},report
  assert report['status']=='partial' and report['applied_count']==2 and report['remaining']==[0],report
  def read(c):return pane.request_threadsafe('cell_get',{'executionTargetSection':'target','table_index':1,'row_index':1,'cell_index':c})['text'].strip()
  assert read(1)=='Coauthor retained'
  assert read(2)=='Independent target succeeds'
  evidence={'prepared':prepared,'report':report,'coauthor_content_preserved':True,'independent_target_applied':True}
  Path('/private/tmp/word-stream-conflict-evidence.json').write_text(json.dumps(evidence,indent=2))
  print(json.dumps(report),flush=True)
asyncio.run(main())
