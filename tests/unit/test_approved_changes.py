"""Real server mutators driven through a scoped pane double; no model loop."""
import copy
import hashlib
from collections import Counter
from types import SimpleNamespace

import pytest

from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import write_mode as wm
from verified_docx_mcp.live.session import LiveOpFailed


class Pane:
    document_name = "collaborative.docx"
    document_url = "https://example.test/collaborative.docx"
    hello = SimpleNamespace(capabilities=["approved_changes_v1", "cell_edit", "cell_multiline", "table_edit",
                                          "cell_font_inheritance_v1"])

    def __init__(self):
        self.sections = [{"slug": "target", "paragraphs": ["Target", "Proof is current.", "Delete the whole paragraph."],
                          "tables": [{"table_index": 1, "rows": [["", "same", "same", "keep"] for _ in range(12)]}]}]
        self.calls = Counter()
        self.outside = 0
        self.after = None
        self.fail = None
        self.font_requests = []
        self.cell_rpr = '<w:rPr><w:sz w:val="17"/></w:rPr>'  # an 8.5pt table
        self.cell_payloads = []

    def cell_ooxml(self, text):
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        pkg = "http://schemas.microsoft.com/office/2006/xmlPackage"
        # An empty cell has no text run: its font lives on the paragraph mark, as in real Word.
        para = (f"<w:p><w:r>{self.cell_rpr}<w:t>{text}</w:t></w:r></w:p>" if text
                else f"<w:p><w:pPr>{self.cell_rpr}</w:pPr></w:p>")
        return (f'<pkg:package xmlns:pkg="{pkg}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>'
                f'<w:document xmlns:w="{ns}"><w:body>{para}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>')

    def request_threadsafe(self, op, payload=None, **kwargs):
        self.calls[op] += 1
        payload = payload or {}
        if op == "approved_cell_set":
            result = self.request_threadsafe("cell_set", payload, **kwargs)
            result["readback"] = self.request_threadsafe("cell_get", payload)
            return result
        if op == "execution_snapshot":
            return {"sections": copy.deepcopy(self.sections)}
        if op == "comments_list":
            return {"comments": [], "source": "live"}
        if op == "describe":
            return {"bodySha256": hashlib.sha256(str(self.outside).encode()).hexdigest()}
        row = self.sections[0]["tables"][0]["rows"]
        if op == "cell_get":
            text = row[payload["row_index"]-1][payload["cell_index"]-1]
            reply = {"text": text, "paragraphs": text.split("\n")}
            if payload.get("include_ooxml"):
                self.font_requests.append(payload["include_ooxml"])
                reply["ooxml"] = self.cell_ooxml(text)
            return reply
        if op == "table_get":
            return {"rows": [[{"text": c} for c in r] for r in row], "merged": False,
                    "style": "Table Grid", "headerRowCount": 1}
        if op in {"cell_set", "cells_set", "save"}:
            assert kwargs.get("expected_body_sha256") is None
            assert "expectedBodySha256" not in payload
            if payload["executionSections"] != self.sections:
                raise LiveOpFailed("LIVE_STALE", "Protected section changed; nothing written")
            if self.fail:
                raise self.fail
            if op == "cell_set":
                self.cell_payloads.append(payload)
                row[payload["row_index"]-1][payload["cell_index"]-1] = "\n".join(
                    "".join("\n" if r.get("hard_break") else r["text"] for r in p) for p in payload["paragraphs"])
            if op == "cells_set":
                for c in payload["cells"]:
                    row[c["row_index"]-1][c["cell_index"]-1] = "\n".join(
                        "".join(r["text"] for r in p) for p in c["paragraphs"])
            self.outside += 1  # co-author text edits and comment resolutions elsewhere
            scope_after = runs.scope_sha(self.sections)
            if self.after:
                self.after(self)
            return {"applied": True, "saved": True, "pre": "a"*64, "post": scope_after,
                    "shapes_before": 0, "shapes_after": 0}
        raise AssertionError(op)


@pytest.fixture(params=["legacy", "stream"])
def setup(tmp_path, monkeypatch, request):
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path / "state"))
    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    pane = Pane()
    pane.hello = SimpleNamespace(capabilities=[*pane.hello.capabilities, *(["approved_cell_stream_v1"] if request.param == "stream" else [])])
    monkeypatch.setattr(wm, "live_session_for", lambda path: runs.scoped_session(pane))
    return str(tmp_path / pane.document_name), pane


def cell(row=1, col=1, text="new"):
    return {"op": "replace_cell_markdown", "section": "target",
            "args": {"table_id": 1, "row_index": row, "cell_index": col, "markdown": text}}


def execute(path, prepared):
    return runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")


def test_36_cells_one_prepare_one_execute_no_full_reads_with_coauthors(setup):
    path, pane = setup
    ops = [cell(r, c, f"**{r}:{c}**") for r in range(1, 13) for c in range(1, 4)]
    ops.append({"op": "live_save", "section": "", "args": {}})
    prepared = runs.prepare(path, ops)
    report = execute(path, prepared)
    assert report["status"] == "complete", report
    assert report["applied_count"] == 37
    assert pane.calls["execution_snapshot"] == 2  # one initial, one final
    assert pane.calls["comments_list"] == 1
    assert pane.calls["body_ooxml"] == 0
    assert pane.calls["describe"] == 0
    assert pane.calls["cell_set"] == 36
    assert pane.calls["save"] == 1
    assert pane.sections[0]["tables"][0]["rows"][11] == ["12:1", "12:2", "12:3", "keep"]
    # A duplicate execute is a compact report, not a replay.
    assert execute(path, prepared) == report
    assert pane.calls["cell_set"] == 36


def test_empty_repeated_and_multiline_cells_and_row(setup):
    path, pane = setup
    prepared = runs.prepare(path, [cell(1, 2, "one\n\ntwo"), cell(1, 3, ""),
        {"op": "replace_table_row", "section": "target", "args": {"table_id": 1, "row_index": 2,
         "cells": ["", "same", "**new**", "keep"]}}])
    result = execute(path, prepared)
    assert result["status"] == "complete", result
    assert pane.sections[0]["tables"][0]["rows"][0] == ["", "one\ntwo", "", "keep"]


def test_same_section_change_stops_before_next_write_and_resumes_suffix(setup):
    path, pane = setup
    prepared = runs.prepare(path, [cell(1), cell(2)])
    def edit(p):
        p.sections[0]["paragraphs"][1] = "Coauthor changed proof."
        p.after = None
    pane.after = edit
    stopped = execute(path, prepared)
    assert stopped["status"] == "partial"
    assert stopped["applied_count"] == 2
    pane.sections[0]["paragraphs"][1] = "Proof is current."
    resumed = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"])
    assert resumed["status"] == "complete"
    # Every cell write is preceded by exactly one first-run font read (issue #107); the
    # readbacks (two) are unchanged, and verified operations never replay.
    assert len(pane.font_requests) == pane.calls["cell_set"]
    assert pane.calls["cell_get"] - len(pane.font_requests) == 2


def test_unknown_after_dispatch_cannot_resume(setup):
    path, pane = setup
    prepared = runs.prepare(path, [cell()])
    def disconnect(p):
        raise RuntimeError("connection lost after mutation")
    pane.after = disconnect
    result = execute(path, prepared)
    assert result["status"] == "unknown"
    with pytest.raises(VerifyError, match="Unknown write outcome"):
        runs.execute(path, prepared["run_id"], prepared["manifest_sha256"])
    assert pane.calls["cell_set"] == 1


def test_approved_cell_write_carries_the_first_run_font_in_its_single_write(setup):
    path, pane = setup
    result = execute(path, runs.prepare(path, [cell()]))
    assert result["status"] == "complete", result
    assert pane.calls["cell_set"] == 1  # the font rides in the same write, never a second one
    assert pane.cell_payloads[0]["inherit_font"] == {"size": 8.5}


def test_approved_cell_without_a_direct_font_sends_no_font(setup):
    path, pane = setup
    pane.cell_rpr = ""
    assert execute(path, runs.prepare(path, [cell()]))["status"] == "complete"
    assert "inherit_font" not in pane.cell_payloads[0]


def test_approved_inherit_format_none_is_bound_into_the_manifest_and_skips_the_font_read(setup):
    path, pane = setup
    op = cell()
    op["args"]["inherit_format"] = "none"
    assert execute(path, runs.prepare(path, [op]))["status"] == "complete"
    assert pane.font_requests == [] and "inherit_font" not in pane.cell_payloads[0]
    op["args"]["inherit_format"] = "sometimes"
    with pytest.raises(VerifyError, match="inherit_format"):
        runs.prepare(path, [op])


def test_approved_cell_write_on_a_pane_without_font_inheritance_is_refused_before_writing(setup):
    path, pane = setup
    pane.hello = SimpleNamespace(capabilities=[c for c in pane.hello.capabilities if c != "cell_font_inheritance_v1"])
    result = execute(path, runs.prepare(path, [cell()]))
    assert result["status"] != "complete"
    assert pane.calls["cell_set"] == 0  # refused before dispatch: a font reset is never silent


def test_a_save_in_the_run_does_not_settle_an_uncertain_write(setup):
    path, pane = setup
    ops = [cell(), {"op": "live_save", "section": "", "args": {}}]
    prepared = runs.prepare(path, ops)
    def disconnect(p):
        if p.calls["cell_set"] and not p.calls["save"]:
            raise RuntimeError("connection lost after mutation")
    pane.after = disconnect
    result = execute(path, prepared)
    assert result["status"] in {"unknown", "partial"}
    assert result["applied"] is False
    # Whatever the save did, the uncertain cell write is still uncertain and is never replayed.
    with pytest.raises(VerifyError, match="Unknown write outcome|uncertain"):
        runs.execute(path, prepared["run_id"], prepared["manifest_sha256"])
    assert pane.calls["cell_set"] == 1


def test_approval_digest_document_and_structure_bound(setup):
    path, pane = setup
    prepared = runs.prepare(path, [cell()])
    with pytest.raises(VerifyError):
        runs.execute(path, prepared["run_id"], "0"*64, "Lead")
    with pytest.raises(VerifyError):
        runs.execute(path, prepared["run_id"], prepared["manifest_sha256"])
    with pytest.raises(VerifyError):
        runs.execute(path.replace("collaborative", "different"), prepared["run_id"], prepared["manifest_sha256"], "Lead")
    pane.sections[0]["tables"][0]["table_index"] = 2
    assert execute(path, prepared)["status"] == "partial"


def test_projection_scoped_text_and_whole_paragraph_delete():
    sections = [{"slug":"s", "paragraphs":["Heading", "whole paragraph.", "proof"], "tables":[]}]
    op = {"op":"delete_paragraph", "section":"s", "args":{"anchor_text":"whole paragraph."}}
    assert runs.project(sections, op)[0]["paragraphs"] == ["Heading", "proof"]
    op["args"]["anchor_text"] = "whole"
    with pytest.raises(VerifyError):
        runs.project(sections, op)


def test_manifest_tampering_and_crashed_dispatch(setup):
    import json
    path, _ = setup
    p = runs.prepare(path, [cell()])
    root = runs.folder(path)
    journal_path = root / f"{p['run_id']}.journal.json"
    journal = json.loads(journal_path.read_text())
    journal["status"] = "dispatching"
    runs.write(journal_path, journal)
    with pytest.raises(VerifyError, match="Unknown write outcome"):
        execute(path, p)
    manifest_path = root / f"{p['run_id']}.manifest.json"
    manifest = json.loads(manifest_path.read_text())
    manifest["operations"][0]["args"]["markdown"] = "tampered"
    runs.write(manifest_path, manifest)
    with pytest.raises(VerifyError, match="Manifest changed"):
        execute(path, p)


def test_retry_is_bounded_and_only_for_proven_prewrite_refusal(setup):
    path, pane = setup
    p = runs.prepare(path, [cell()])
    pane.fail = LiveOpFailed('stale', 'refused before writing')
    result = execute(path, p)
    assert result['status'] == 'partial'
    assert pane.calls['cell_set'] == 3
    assert result['applied_count'] == 0
    pane.fail = None
    assert runs.execute(path, p['run_id'], p['manifest_sha256'])['status'] == 'complete'


def test_uncertainty_blocks_new_run_until_explicit_supported_reconciliation(setup):
    path, pane = setup
    p = runs.prepare(path, [cell()])
    pane.after = lambda p: (_ for _ in ()).throw(RuntimeError('disconnected after write'))
    assert execute(path, p)['status'] == 'unknown'
    with pytest.raises(VerifyError, match='uncertain outcome'):
        runs.prepare(path, [cell(2)])
    with pytest.raises(VerifyError, match='does not support'):
        runs.reconcile(path, p['run_id'], p['manifest_sha256'], 'not-applied', 'Decision 1')
    reconciled = runs.reconcile(path, p['run_id'], p['manifest_sha256'], 'applied', 'Decision 1')
    assert reconciled['applied_count'] == 1
    pane.after = None
    assert runs.execute(path, p['run_id'], p['manifest_sha256'])['status'] == 'complete'
    assert pane.calls['cell_set'] == 1


def test_two_document_runs_have_independent_manifests_and_journals(setup):
    path, _ = setup
    a = runs.prepare(path, [cell()])
    second = path.replace('collaborative', 'second')
    b = runs.prepare(second, [cell(2)])
    assert runs.folder(path) != runs.folder(second)
    assert a['run_id'] != b['run_id']
    assert execute(path, a)['status'] == 'complete'
    assert execute(second, b)['status'] == 'partial'  # same fake pane changed, separate baseline is intact


def test_uncertain_run_does_not_block_disjoint_section(setup):
    path, pane = setup
    p = runs.prepare(path, [cell()])
    root = runs.folder(path)
    import json
    jf = root / f"{p['run_id']}.journal.json"
    journal=json.loads(jf.read_text()); journal['status']='unknown'; runs.write(jf,journal)
    # The same document has another independent section. Its preparation must
    # not erase the unresolved first run or require a lead override.
    pane.sections.append({'slug':'other','paragraphs':['Other'], 'tables':[{'table_index':1,'rows':[['old']]}]})
    original = pane.request_threadsafe
    def request(op,payload=None,**kwargs):
        if op == 'execution_snapshot':
            return {'sections':[copy.deepcopy(s) for s in pane.sections if s['slug'] in payload['section_slugs']]}
        return original(op,payload,**kwargs)
    pane.request_threadsafe=request
    other=cell(); other['section']='other'
    second=runs.prepare(path,[other])
    assert second['status']=='prepared'
    assert json.loads(jf.read_text())['status']=='unknown'


def test_formatting_only_unknown_cannot_be_certified_from_unchanged_text(setup):
    import json
    path,_pane=setup
    p=runs.prepare(path,[cell(1,2,'**same**')])
    jf=runs.folder(path)/f"{p['run_id']}.journal.json"
    j=json.loads(jf.read_text()); j['status']='unknown'; runs.write(jf,j)
    with pytest.raises(VerifyError,match='formatting-only'):
        runs.reconcile(path,p['run_id'],p['manifest_sha256'],'applied','Decision 2')


def test_uncertain_save_requires_words_saved_state(setup):
    import json
    path,_pane=setup
    p=runs.prepare(path,[cell(),{'op':'live_save','section':'','args':{}}])
    assert execute(path,p)['status']=='complete'
    jf=runs.folder(path)/f"{p['run_id']}.journal.json"
    j=json.loads(jf.read_text());j.update(status='unknown',next=1);runs.write(jf,j)
    with pytest.raises(VerifyError,match='Word must report saved=true'):
        runs.reconcile(path,p['run_id'],p['manifest_sha256'],'applied','Decision 3')


@pytest.mark.parametrize("refusal", ["WOULD_DELETE_COMMENTS", "MIXED_FORMATTING", "ANCHORED_SHAPES"])
def test_known_failure_defers_one_cell_and_finishes_the_rest(setup,refusal):
    path,pane=setup
    p=runs.prepare(path,[cell(r) for r in range(1,7)]+[{'op':'live_save','section':'','args':{}}])
    original=pane.request_threadsafe
    def request(op,payload=None,**kw):
        if op=='cell_set' and payload['row_index']==3:
            raise LiveOpFailed(refusal,'target cannot be safely replaced')
        return original(op,payload,**kw)
    pane.request_threadsafe=request
    report=execute(path,p)
    assert report['status']=='partial'
    assert report['applied_count']==6  # five independent cells and the save
    assert report['remaining']==[2]
    assert report['issues'][0]['code']==refusal
    assert report['issues'][0]['status']=='conflict'
    assert pane.sections[0]['tables'][0]['rows'][5][0]=='new'
    assert pane.calls['save']==1


def test_uncertain_third_cell_does_not_prevent_later_independent_cells(setup):
    path,pane=setup
    p=runs.prepare(path,[cell(r) for r in range(1,7)])
    def disconnect_once(pane):
        if pane.outside==3:
            pane.after=None
            raise RuntimeError('lost reply after third cell')
    pane.after=disconnect_once
    result=execute(path,p)
    assert result['status']=='partial'
    assert result['applied_count']==5
    assert result['remaining']==[2]
    assert result['issues'][0]['status']=='unknown'
    assert all(row[0]=='new' for row in pane.sections[0]['tables'][0]['rows'][:6])
    assert pane.calls['cell_set']==6  # uncertain operation was never repeated


def test_dependent_edit_is_deferred_while_other_targets_finish(setup):
    path,pane=setup
    p=runs.prepare(path,[cell(1,text='first'),cell(1,text='second'),cell(2)])
    original=pane.request_threadsafe
    def request(op,payload=None,**kw):
        if op=='cell_set' and payload['row_index']==1:
            raise LiveOpFailed('WOULD_DELETE_COMMENTS','comment')
        return original(op,payload,**kw)
    pane.request_threadsafe=request
    result=execute(path,p)
    assert result['applied_count']==1
    assert result['remaining']==[0,1]
    assert result['issues'][1]['status']=='dependency'
    assert pane.sections[0]['tables'][0]['rows'][1][0]=='new'


def test_save_occurs_after_successful_retry(setup):
    path,pane=setup
    p=runs.prepare(path,[cell(1),cell(2),{'op':'live_save','section':'','args':{}}])
    original=pane.request_threadsafe
    events=[]
    refused=False
    def request(op,payload=None,**kw):
        nonlocal refused
        if op=='cell_set' and payload['row_index']==1 and not refused:
            refused=True
            raise LiveOpFailed('LIVE_STALE','pre-write refusal')
        result=original(op,payload,**kw)
        if op in {'cell_set','save'}: events.append(op)
        return result
    pane.request_threadsafe=request
    assert execute(path,p)['status']=='complete'
    assert events==['cell_set','cell_set','save']


def test_final_read_disconnect_resumes_without_replaying_writes(setup):
    path,pane=setup
    p=runs.prepare(path,[cell(),{'op':'live_save','section':'','args':{}}])
    original=pane.request_threadsafe
    disconnected=False
    def request(op,payload=None,**kw):
        nonlocal disconnected
        if op=='execution_snapshot' and not disconnected:
            disconnected=True
            raise RuntimeError('final read disconnected')
        return original(op,payload,**kw)
    pane.request_threadsafe=request
    assert execute(path,p)['status']=='verification_pending'
    assert execute(path,p)['status']=='complete'
    assert pane.calls['cell_set']==1
    assert pane.calls['save']==1


def test_resume_after_partial_result_saves_newly_applied_edits(setup):
    path,pane=setup
    p=runs.prepare(path,[cell(1),cell(2),{'op':'live_save','section':'','args':{}}])
    original=pane.request_threadsafe
    def request(op,payload=None,**kw):
        if op=='cell_set' and payload['row_index']==1:
            raise LiveOpFailed('WOULD_DELETE_COMMENTS','comment')
        return original(op,payload,**kw)
    pane.request_threadsafe=request
    assert execute(path,p)['status']=='partial'
    assert pane.calls['save']==1
    pane.request_threadsafe=original
    assert execute(path,p)['status']=='complete'
    assert pane.calls['save']==2


def test_direct_scoped_readback_uses_same_deadline_as_write():
    calls=[]
    class Direct:
        def request_threadsafe(self,op,payload=None,**kw): calls.append((op,kw['timeout']))
    scoped=runs.ScopedSession(Direct(),[{'slug':'target'}])
    scoped.request_threadsafe('cell_set',{})
    scoped.request_threadsafe('cell_get',{'executionTargetSection':'target'})
    assert calls==[('cell_set',60),('cell_get',60)]


def test_refused_anchored_delete_keeps_independent_cell_eligible(setup):
    path,pane=setup
    native=pane.request_threadsafe
    def refused_delete(op,payload=None,**kw):
        if op=='paragraph_delete': raise LiveOpFailed('ANCHORED_SHAPES','Anchor paragraph cannot be removed')
        return native(op,payload,**kw)
    pane.hello.capabilities.append('delete_paragraph')
    pane.request_threadsafe=refused_delete
    ops=[{'op':'delete_paragraph','section':'target','args':{'anchor_text':'Delete the whole paragraph.'}},cell(),{'op':'live_save','section':'','args':{}}]
    prepared=runs.prepare(path,ops)
    result=execute(path,prepared)
    assert result['status']=='partial' and result['applied_count']==2,result
    assert result['remaining']==[0]
    assert result['issues'][0]['code']=='ANCHORED_SHAPES'
    assert pane.sections[0]['tables'][0]['rows'][0][0]=='new'
