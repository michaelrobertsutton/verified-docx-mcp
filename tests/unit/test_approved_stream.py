import json
import subprocess
from pathlib import Path
from types import SimpleNamespace

from test_approved_changes import cell, execute, setup  # noqa: F401

from verified_docx_mcp.live import approved_changes as runs


def test_pane_stream_harness():
    root=Path(__file__).resolve().parents[2]
    subprocess.run(['node',str(root/'tests/unit/js/approved_stream_harness.mjs'),str(root/'addin/taskpane.js')],check=True,capture_output=True,text=True)


def test_stream_one_ack_independent_readback_per_target(setup):  # noqa: F811
    path,pane=setup
    pane.hello=SimpleNamespace(capabilities=[*pane.hello.capabilities,'approved_cell_stream_v1'])
    native=pane.request_threadsafe
    def streamed(op,payload=None,**kwargs):
        if op!='approved_cell_set': return native(op,payload,**kwargs)
        result=native('cell_set',payload,**kwargs)
        pane.calls['approved_cell_set']+=1
        result['readback']=native('cell_get',payload)
        result['metrics']={'sync_count':18,'phases':{'preflight':{'sync_count':10}}}
        return result
    pane.request_threadsafe=streamed
    prepared=runs.prepare(path,[cell(),{'op':'live_save','section':'','args':{}}])
    result=execute(path,prepared)
    assert result['status']=='complete'
    assert pane.calls['approved_cell_set']==1
    journal=json.loads(Path(prepared['manifest_path'].replace('.manifest.','.journal.')).read_text())
    assert journal['evidence'][0]['metrics']['sync_count']==18
    assert execute(path,prepared)['status']=='complete'
    assert pane.calls['approved_cell_set']==1


def test_only_proven_prewrite_capability_failure_falls_back(setup):  # noqa: F811
    from verified_docx_mcp.live.session import LiveOpFailed
    path,pane=setup
    pane.hello=SimpleNamespace(capabilities=[*pane.hello.capabilities,'approved_cell_stream_v1'])
    native=pane.request_threadsafe
    def unsupported(op,payload=None,**kw):
        if op=='approved_cell_set': raise LiveOpFailed('LIVE_CAPABILITY_MISSING','Merged table needs conservative path')
        return native(op,payload,**kw)
    pane.request_threadsafe=unsupported
    prepared=runs.prepare(path,[cell()])
    assert execute(path,prepared)['status']=='complete'
    assert pane.calls['cell_set']==1
    assert pane.calls['cell_get']==1


def test_uncertain_stream_is_never_replayed_or_fallen_back(setup):  # noqa: F811
    from verified_docx_mcp.live.session import LiveOpFailed
    path,pane=setup
    pane.hello=SimpleNamespace(capabilities=[*pane.hello.capabilities,'approved_cell_stream_v1'])
    native=pane.request_threadsafe
    attempted=[]
    def uncertain(op,payload=None,**kw):
        if op=='approved_cell_set':
            attempted.append(op)
            raise LiveOpFailed('VERIFICATION_FAILED','Write may have landed')
        return native(op,payload,**kw)
    pane.request_threadsafe=uncertain
    prepared=runs.prepare(path,[cell()])
    result=execute(path,prepared)
    assert result['issues'][0]['status']=='unknown'
    assert len(attempted)==1 and pane.calls['cell_set']==0
