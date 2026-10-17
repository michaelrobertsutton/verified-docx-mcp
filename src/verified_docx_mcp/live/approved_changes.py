"""Durable, digest-bound live runs. No model round trips between operations.

The pane validates a structured section baseline at the mutation boundary;
whole-body revision tokens remain evidence, not a concurrency precondition.
"""
from __future__ import annotations

import copy
import fcntl
import hashlib
import json
import os
import re
import time
import uuid
from contextlib import contextmanager
from contextvars import ContextVar
from difflib import unified_diff
from typing import Any

from .. import audit, paths, tables
from ..errors import ErrorCode, VerifyError, _make_error
from . import write_mode as wm
from .session import LiveOpFailed

_TARGET = ContextVar("approved_change_target", default=None)
_SCOPE: ContextVar[list | None] = ContextVar("approved_change_scope", default=None)
OPS = {
    "replace_cell_markdown": {"table_id", "row_index", "cell_index", "markdown"},
    "replace_table_row": {"table_id", "row_index", "cells"},
    "replace_text": {"find", "replace", "expected_matches"},
    "delete_paragraph": {"anchor_text"},
    "live_save": set(),
}
# Only explicit errors proven to occur before dispatch/mutation may be retried.
RETRY = {"LIVE_STALE", "LOCKED_BY_OTHER_CLIENT"}
REFUSED = RETRY | {"MATCH_COUNT_MISMATCH", "ZERO_MATCH", "INVALID_INPUT", "CELL_TEXT_MISMATCH",
                   "LIVE_CAPABILITY_MISSING", "OUTSIDE_LOCKED_SECTION", "ANCHORED_SHAPES",
                   "WOULD_DELETE_COMMENTS", "LOCK_SCOPE_UNRESOLVED", "MIXED_FORMATTING",
                   "STRUCTURAL_BOUNDARY", "TRACKED_CHANGES_PRESENT"}


def invalid(message):
    return _make_error(ErrorCode.INVALID_INPUT, message)


def sha(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def write(path, value):
    tmp = path.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as stream:
        json.dump(value, stream, ensure_ascii=False)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(tmp, path)


def folder(path):
    root = audit._state_dir() / "approved-changes" / hashlib.sha256(path.encode()).hexdigest()
    root.mkdir(mode=0o700, parents=True, exist_ok=True)
    return root


@contextmanager
def locked(root):
    fd = os.open(root / "document.lock", os.O_CREAT | os.O_RDWR, 0o600)
    with os.fdopen(fd, "a") as stream:
        fcntl.flock(stream, fcntl.LOCK_EX)
        yield


class ScopedSession:
    def __init__(self, session, sections):
        self.session, self.sections = session, sections
        self.request_count = 0

    def __getattr__(self, name):
        return getattr(self.session, name)

    def request_threadsafe(self, op, payload=None, **kwargs):
        from .shared import READ_OPS
        data = dict(payload or {})
        if op not in READ_OPS:
            data.pop("expectedBodySha256", None)
            data["executionSections"] = self.sections
            data["executionTargetSection"] = _TARGET.get()
            kwargs["expected_body_sha256"] = None
        from .session import resolve_timeout
        kwargs.setdefault("timeout", resolve_timeout(op, protected_sections=True))
        self.request_count += 1
        return self.session.request_threadsafe(op, data, **kwargs)


def scoped_session(session):
    sections = _SCOPE.get()
    return ScopedSession(session, sections) if sections is not None else session


def snapshot(session, slugs, details=False):
    wm.require_capability(session, "approved_changes_v1",
                          feature_description="approved runs; upgrade the server and reload the Word pane")
    result = session.request_threadsafe("execution_snapshot", {"section_slugs": slugs})
    return result if details else result["sections"]


def canonical(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def plain(markdown):
    return "\n".join(tables._intended_cell_paragraphs(tables.markdown_to_ooxml.parse_paragraph_runs(markdown)))


def project(sections, operation):
    """Resolve exactly one structured target and compute the approved next state."""
    state = copy.deepcopy(sections)
    op, args = operation["op"], operation["args"]
    if op == "live_save":
        return state
    scope = [s for s in state if s["slug"] == operation["section"]]
    if len(scope) != 1:
        raise invalid("Operation needs one prepared section")
    section = scope[0]
    if op in {"replace_cell_markdown", "replace_table_row"}:
        found = [t for t in section["tables"] if t["table_index"] == args["table_id"]]
        if len(found) != 1:
            raise invalid("Table is absent from the approved section")
        table = found[0]
        row = args["row_index"] - 1
        if not 0 <= row < len(table["rows"]):
            raise invalid("Row is outside the prepared table")
        if op == "replace_cell_markdown":
            cell = args["cell_index"] - 1
            if not 0 <= cell < len(table["rows"][row]):
                raise invalid("Cell is outside the prepared row")
            table["rows"][row][cell] = plain(args["markdown"])
            if "paragraph_counts" in table:
                table["paragraph_counts"][row][cell] = len(tables._intended_cell_paragraphs(tables.markdown_to_ooxml.parse_paragraph_runs(args["markdown"])))
        else:
            if len(args["cells"]) != len(table["rows"][row]):
                raise invalid("Row replacement changes the column count")
            table["rows"][row] = [plain(c) for c in args["cells"]]
            if "paragraph_counts" in table:
                table["paragraph_counts"][row] = [len(tables._intended_cell_paragraphs(tables.markdown_to_ooxml.parse_paragraph_runs(c))) for c in args["cells"]]
    elif op == "delete_paragraph":
        hits = [i for i, p in enumerate(section["paragraphs"]) if p == args["anchor_text"]]
        if len(hits) != 1 or hits[0] == 0:
            raise invalid("delete_paragraph needs the unique whole paragraph, excluding the heading")
        section["paragraphs"].pop(hits[0])
        for table in section["tables"]:
            if table.get("after_paragraph", -1) >= hits[0]:
                table["after_paragraph"] -= 1
    else:
        find = canonical(args["find"])
        locations = [(section["paragraphs"], i) for i, p in enumerate(section["paragraphs"]) if find in p]
        locations += [(row, i) for t in section["tables"] for row in t["rows"] for i, p in enumerate(row) if find in p]
        if len(locations) != 1 or locations[0][0][locations[0][1]].count(find) != 1:
            raise invalid("replace_text needs one exact target in its section")
        values, index = locations[0]
        if values is section["paragraphs"] and index == 0:
            raise invalid("Changing a protected heading requires a new plan")
        values[index] = values[index].replace(find, canonical(args["replace"]), 1)
    return state


def target_keys(sections, operation):
    op, args, slug = operation["op"], operation["args"], operation["section"]
    if op == "live_save": return []
    section = next(s for s in sections if s["slug"] == slug)
    if op == "replace_cell_markdown":
        return [f"{slug}|t|{args['table_id']}|{args['row_index']}|{args['cell_index']}"]
    if op == "replace_table_row":
        return [f"{slug}|t|{args['table_id']}|{args['row_index']}|{c+1}" for c in range(len(args['cells']))]
    needle = canonical(args.get("find", args.get("anchor_text")))
    keys = [f"{slug}|p|{i}" for i,text in enumerate(section['paragraphs']) if needle in text]
    keys += [f"{slug}|t|{t['table_index']}|{r+1}|{c+1}" for t in section['tables'] for r,row in enumerate(t['rows']) for c,text in enumerate(row) if needle in text]
    return keys


def target_value(sections, key):
    parts=key.split('|'); section=next(s for s in sections if s['slug']==parts[0])
    if parts[1]=='p': return section['paragraphs'][int(parts[2])]
    table=next(t for t in section['tables'] if t['table_index']==int(parts[2]))
    r,c=int(parts[3])-1,int(parts[4])-1
    return {'text':table['rows'][r][c], 'paragraphs':table.get('paragraph_counts', [[None]*len(row) for row in table['rows']])[r][c]}


def changed_targets(before, after):
    """Quarantine only changed leaves; structural changes quarantine their section."""
    changed=set()
    if [s['slug'] for s in before] != [s['slug'] for s in after]:
        return {s['slug']+'|*' for s in before}
    for old,new in zip(before,after):
        slug=old['slug']
        if len(old['paragraphs']) != len(new['paragraphs']) or len(old['tables']) != len(new['tables']):
            changed.add(slug+'|*'); continue
        if old['paragraphs'][0] != new['paragraphs'][0]:
            changed.add(slug+'|*'); continue
        for i,(a,b) in enumerate(zip(old['paragraphs'],new['paragraphs'])):
            if a != b: changed.add(f'{slug}|p|{i}')
        for a,b in zip(old['tables'],new['tables']):
            if (a['table_index'],a.get('after_paragraph'),[len(r) for r in a['rows']]) != (b['table_index'],b.get('after_paragraph'),[len(r) for r in b['rows']]):
                changed.add(slug+'|*');continue
            for r,(ar,br) in enumerate(zip(a['rows'],b['rows'])):
                for c,(at,bt) in enumerate(zip(ar,br)):
                    key=f"{slug}|t|{a['table_index']}|{r+1}|{c+1}"
                    if at != bt or target_value(before,key) != target_value(after,key): changed.add(key)
    return changed


def validate(operations):
    if not isinstance(operations, list) or not 1 <= len(operations) <= 500:
        raise invalid("Supply 1..500 ordered operations")
    for i, item in enumerate(operations):
        if not isinstance(item, dict) or set(item) != {"op", "section", "args"} or item.get("op") not in OPS:
            raise invalid(f"Operation {i} requires op, section and args from the supported operation set")
        op, args = item["op"], item["args"]
        if not isinstance(args, dict) or set(args) - ({"replacement_segments"} if op == "replace_text" else set()) != OPS[op]:
            raise invalid(f"Operation {i} has unexpected or missing arguments")
        if op != "live_save" and (not isinstance(item["section"], str) or not item["section"]):
            raise invalid("Every text operation must name its protected section slug")
        if op == "replace_text":
            from ..formatted_text import validate_segments
            validate_segments(args.get("replacement_segments"), args["replace"])
        for key, value in args.items():
            if key == "replacement_segments":
                continue
            if key in {"table_id", "row_index", "cell_index"}:
                if type(value) is not int or value < 1:
                    raise invalid(f"{key} must be a positive integer")
            elif key == "cells":
                if not isinstance(value, list) or not value or any(not isinstance(v, str) for v in value):
                    raise invalid("cells must be a nonempty string list")
            elif key == "expected_matches":
                if type(value) is not int or value != 1:
                    raise invalid("expected_matches must be integer 1")
            elif not isinstance(value, str):
                raise invalid(f"{key} must be text")
        if op == "replace_text" and (not args["find"] or len(args["find"]) > 230 or
                                     any(c in args["find"] + args["replace"] for c in "\r\n")):
            raise invalid("Text replacements need a nonempty find <=230 characters and no paragraph breaks")
        if op == "live_save" and i != len(operations) - 1:
            raise invalid("live_save must be the last operation")


def prepare(path: str, operations: list[dict[str, Any]], track_changes: bool = False) -> dict:
    path = str(paths.resolve_allowed_docx_path(path, must_exist=False))
    validate(operations)
    slugs = list(dict.fromkeys(o["section"] for o in operations if o["op"] != "live_save"))
    if not slugs:
        raise invalid("At least one protected text operation is required")
    session = wm.live_session_for(path)
    root = folder(path)
    with locked(root):
        ensure_settled(root, sections=slugs)
        if any(o["args"].get("replacement_segments") is not None for o in operations):
            wm.require_capability(session, "formatted_replacements_v1", feature_description="explicit formatted replacements")
        captured = snapshot(session, slugs, details=True)
        initial = captured["sections"]
        operations = copy.deepcopy(operations)
        for operation in operations:
            if operation["op"] in {"replace_cell_markdown", "replace_table_row"} and "table_addresses" in captured:
                matches = [a for a in captured["table_addresses"] if a["slug"] == operation["section"] and a["global_index"] == operation["args"]["table_id"]]
                if len(matches) != 1:
                    raise invalid("Table is absent or ambiguous in the prepared section")
                operation["args"]["table_id"] = matches[0]["local_index"]
        comments = session.request_threadsafe("comments_list", {"include_anchor": True})
        current = initial
        targets, preconditions = [], []
        for operation in operations:
            keys = target_keys(current, operation)
            targets.append(keys)
            preconditions.append({key:target_value(current,key) for key in keys})
            current = project(current, operation)
        run_id = uuid.uuid4().hex
        manifest = {"schema": 1, "run_id": run_id, "path": path, "document_url": session.document_url,
                    "sections": initial, "operations": operations, "comments": comments, "track_changes": track_changes, "targets": targets, "preconditions": preconditions}
        if any(o["args"].get("replacement_segments") is not None for o in operations):
            manifest["formatting_policy"] = "explicit-segments-v1"
        digest = sha(manifest)
        write(root / f"{run_id}.manifest.json", manifest)
        write(root / f"{run_id}.journal.json", {"digest": digest, "next": 0, "status": "prepared", "evidence": [], "verified_indices": [],
                                                     "working": initial, "deferred": {}, "attempts": {}, "blocked": []})
    return {"run_id": run_id, "manifest_sha256": digest, "operation_count": len(operations),
            "sections": slugs, "status": "prepared", "preview": [{"op": o["op"], "section": o["section"], **({"replacement_segments": o["args"]["replacement_segments"]} if "replacement_segments" in o["args"] else {})} for o in operations],
            "manifest_path": str(root / f"{run_id}.manifest.json")}


def scope_sha(sections):
    return hashlib.sha256(json.dumps(sections, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()


def apply(path, operation, track_changes=False):
    """Reuse pane mutations and independent target readbacks without describe/body reads."""
    session = wm.live_session_for(path)
    op, args = operation["op"], operation["args"]
    current = _SCOPE.get()
    scope = next((s for s in current if s["slug"] == operation["section"]), None)
    common = {"executionTargetSection": operation["section"], "track_changes": track_changes}
    before, after = "", ""
    try:
        if op in {"replace_cell_markdown", "replace_table_row"}:
            table = next(t for t in scope["tables"] if t["table_index"] == args["table_id"])
            row = table["rows"][args["row_index"]-1]
            address = {"table_index": args["table_id"], "row_index": args["row_index"], **common}
            if op == "replace_cell_markdown":
                wm.require_capability(session, "cell_edit", feature_description="scoped cell writes")
                before = row[args["cell_index"]-1]
                after = plain(args["markdown"])
                address["cell_index"] = args["cell_index"]
                paragraphs = tables.markdown_to_ooxml.parse_paragraph_runs(args["markdown"])
                if len(paragraphs) > 1 or any(r.hard_break for p in paragraphs for r in p):
                    wm.require_capability(session, "cell_multiline", feature_description="multiline cell writes")
                streamed = "approved_cell_stream_v1" in session.hello.capabilities and os.environ.get("VERIFIED_DOCX_DISABLE_APPROVED_CELL_STREAM") != "1"
                cell_payload = {**address, "expected_before_text": before,
                    "paragraphs": tables._runs_to_wire(paragraphs),
                    "expected_before_paragraph_count": table.get("paragraph_counts", [[None]*len(r) for r in table["rows"]])[args["row_index"]-1][args["cell_index"]-1]}
                try:
                    result = session.request_threadsafe("approved_cell_set" if streamed else "cell_set", cell_payload)
                except LiveOpFailed as exc:
                    # The stream converts every postwrite failure to VERIFICATION_FAILED.
                    # Only its explicit prewrite capability refusal may use the old path.
                    if not streamed or exc.code != "LIVE_CAPABILITY_MISSING":
                        raise
                    streamed = False
                    result = session.request_threadsafe("cell_set", cell_payload)
                readback = result["readback"] if streamed else session.request_threadsafe("cell_get", address)
                if [canonical(p) for p in readback["paragraphs"]] != tables._intended_cell_paragraphs(paragraphs):
                    raise RuntimeError("Cell readback differs from the approved replacement")
            else:
                wm.require_capability(session, "table_edit", feature_description="scoped row writes")
                before, after = "\t".join(row), "\t".join(plain(c) for c in args["cells"])
                live_table = session.request_threadsafe("table_get", {"table_index": args["table_id"], **common})
                if live_table["merged"]:
                    raise invalid("Row replacement refuses merged tables; use cell operations")
                cells = [{**address, "cell_index": i+1, "expected_before_text": row[i],
                          "paragraphs": tables._runs_to_wire(tables.markdown_to_ooxml.parse_paragraph_runs(md)),
                          "expected_before_paragraph_count": table.get("paragraph_counts", [[None]*len(r) for r in table["rows"]])[args["row_index"]-1][i]}
                         for i, md in enumerate(args["cells"])]
                result = session.request_threadsafe("cells_set", {**common, "cells": cells})
                got = session.request_threadsafe("table_get", {"table_index": args["table_id"], **common})
                if [canonical(c["text"]) for c in got["rows"][args["row_index"]-1]] != [plain(c) for c in args["cells"]]:
                    raise RuntimeError("Row readback differs from the approved replacement")
        elif op == "replace_text":
            if args.get("replacement_segments") is not None:
                wm.require_capability(session, "formatted_replacements_v1", feature_description="explicit formatted replacements")
            before, after = args["find"], args["replace"]
            result = session.request_threadsafe("replace", {**args, **common, "allow_comment_loss": False,
                                                           "inherit_format": "replaced", "scope": "body"})
            if args.get("replacement_segments") is not None:
                from ..formatted_text import validate_segments
                returned = result.get("matches", [{}])[0].get("segments")
                try:
                    verified_segments = validate_segments(returned, after) if returned is not None else None
                except VerifyError as exc:
                    raise RuntimeError("Malformed formatting readback after mutation") from exc
                if verified_segments is None or verified_segments != validate_segments(args["replacement_segments"], after):
                    raise RuntimeError("Independent formatting readback differs from approved segments")
            if len(result.get("matches", [])) != 1 or result["matches"][0]["after"] != after:
                raise RuntimeError("Replacement did not independently verify")
        elif op == "delete_paragraph":
            wm.require_capability(session, "delete_paragraph", feature_description="whole paragraph deletion")
            before = args["anchor_text"]
            result = session.request_threadsafe("paragraph_delete", {**args, **common})
        else:
            result = session.request_threadsafe("save", common)
            if result.get("saved") is not True:
                raise RuntimeError("Word did not confirm save")
            result["applied"] = True
    except LiveOpFailed as exc:
        # Keep the pane's exact refusal classification. Unknown codes stay uncertain.
        code = {"stale": "LIVE_STALE", "zero_match": "ZERO_MATCH", "match_count_mismatch": "MATCH_COUNT_MISMATCH"}.get(exc.code, exc.code)
        known = ErrorCode.__members__.get(code, ErrorCode.LIVE_OP_FAILED)
        raise _make_error(known, exc.message, {"pane_code": exc.code}) from exc
    if result.get("applied") is not True:
        raise RuntimeError("Pane returned no positive applied evidence")
    if op != "live_save" and result.get("post") != scope_sha(project(current, operation)):
        raise RuntimeError("Post-write protected section differs from the approved result")
    evidence = {"applied": True, "before": before, "after": after,
                "revision_scope": "protected-sections", "scope_before": scope_sha(current),
                "scope_after": result.get("post"), "verified_via": "word-addin", "write_mode": "live"}
    evidence["bridge_requests"] = getattr(session, "request_count", None)
    if "metrics" in result:
        evidence["metrics"] = result["metrics"]
    audit.append_audit(path=path, tool=op, evidence=evidence)
    return evidence


def report(manifest, journal):
    verified=set(journal.get('verified_indices',range(journal['next'])))
    performance = {"operation_elapsed_ms": round(sum(e.get("elapsed_ms", 0) for e in journal.get("evidence", [])), 2),
                   "office_sync_count": sum(e.get("metrics", {}).get("sync_count", 0) for e in journal.get("evidence", [])),
                   "bridge_requests": sum(e.get("bridge_requests") or 0 for e in journal.get("evidence", []))}
    return {"performance": performance, "run_id":manifest['run_id'], "status":journal['status'], "applied":journal['status']=='complete',
            "applied_count":len(verified), "operation_count":len(manifest['operations']),
            "remaining":[i for i in range(len(manifest['operations'])) if i not in verified],
            "issues":[dict(index=int(i),**problem) for i,problem in journal.get('deferred',{}).items()],
            "external_changes":journal.get('blocked',[]), "problem":journal.get('problem'), "manifest_sha256":journal['digest']}


def execute(path: str, run_id: str, manifest_sha256: str, approved_by: str | None = None) -> dict:
    path=str(paths.resolve_allowed_docx_path(path,must_exist=False))
    if not re.fullmatch(r'[0-9a-f]{32}',run_id): raise invalid('Invalid run_id')
    root=folder(path)
    with locked(root):
        mf,jf=root/f'{run_id}.manifest.json',root/f'{run_id}.journal.json'
        if not mf.is_file() or not jf.is_file(): raise invalid('No prepared run for this document')
        manifest,journal=json.loads(mf.read_text()),json.loads(jf.read_text())
        if manifest['path']!=path or sha(manifest)!=manifest_sha256 or journal['digest']!=manifest_sha256:
            raise invalid('Manifest changed, digest differs, or document identity differs')
        ensure_settled(root,run_id,sections=[s['slug'] for s in manifest['sections']])
        if journal['status']=='complete': return report(manifest,journal)
        if not journal.get('approved_by'):
            if not isinstance(approved_by,str) or not approved_by.strip(): raise invalid('Lead approval is required for this manifest digest')
            journal['approved_by']=approved_by.strip()
        session=wm.live_session_for(path)
        if session.document_url!=manifest['document_url']: raise invalid('Connected pane document identity differs from the prepared document')
        current=journal.get('working',manifest['sections'])
        verified=set(journal.get('verified_indices',range(journal['next'])))
        deferred=journal.setdefault('deferred',{})
        attempts: dict[str, int]={}
        journal['attempts']=attempts
        blocked=set(journal.get('blocked',[]))
        slugs=[s['slug'] for s in current]
        if journal['status']=='dispatching':
            active=journal.get('active_index',journal['next'])
            deferred[str(active)]={'status':'unknown','reason':'Process ended during dispatch'}
        unknown={int(i) for i,v in deferred.items() if v['status']=='unknown'}
        if journal['status']=='unknown' and not unknown:
            unknown.add(journal.get('active_index',journal['next']))
            deferred[str(next(iter(unknown)))]={'status':'unknown','reason':journal.get('problem')}
        if unknown and all(i in verified or i in unknown for i in range(len(manifest['operations']))):
            raise invalid('Unknown write outcome; reconcile the live target. It will not be retried')
        if journal['status']!='prepared':
            observed=snapshot(session,slugs)
            blocked.update(changed_targets(current,observed))
            current=observed
            uncertain_targets={t for i in unknown for t in manifest['targets'][i]}
            for blocked_key in list(blocked):
                if blocked_key in uncertain_targets or blocked_key.endswith('|*'): continue
                expected_values=[v[blocked_key] for i,v in enumerate(manifest['preconditions']) if i not in verified and blocked_key in v]
                try:
                    if expected_values and target_value(current,blocked_key)==expected_values[0] or not expected_values and target_value(current,blocked_key)==target_value(manifest['sections'],blocked_key): blocked.discard(blocked_key)
                except (KeyError,IndexError,StopIteration): pass
        # Each pass attempts all independent operations before revisiting transient failures.
        for pass_number in range(4):
            for index,operation in enumerate(manifest['operations']):
                if (operation['op']=='live_save') != (pass_number==3): continue
                key=str(index)
                if index in verified or index in unknown: continue
                previous=deferred.get(key,{})
                if previous.get('status')=='conflict' and previous.get('code') not in RETRY and pass_number>0: continue
                if attempts.get(key,0)>=3: continue
                own_targets=set(manifest['targets'][index])
                other_targets=set(blocked)
                for i,value in deferred.items():
                    if i!=key and value['status'] in {'unknown','conflict'}:
                        other_targets.update(manifest['targets'][int(i)])
                overlap=bool(own_targets & other_targets or any(t.split('|')[0]+'|*' in other_targets for t in own_targets))
                if operation['op']=='delete_paragraph' and any(t.startswith(operation['section']+'|') for t in other_targets):
                    overlap=True;blocked.add(operation['section']+'|*')
                if overlap:
                    if previous.get('status') not in {'conflict','unknown'}:
                        deferred[key]={'status':'dependency','reason':'A changed or uncertain target must be resolved first'}
                    continue
                try:
                    matches=all(target_value(current,t)==value for t,value in manifest['preconditions'][index].items())
                except (KeyError,IndexError,StopIteration): matches=False
                if not matches:
                    deferred[key]={'status':'conflict','code':'TARGET_CHANGED','reason':'Expected original target changed; preserved current content'}
                    continue
                attempts[key]=attempts.get(key,0)+1
                if operation['op']!='live_save':
                    # A resumed mutation makes an earlier partial-run save obsolete.
                    verified.difference_update(i for i,op in enumerate(manifest['operations']) if op['op']=='live_save')
                journal.update(status='dispatching',active_index=index,working=current,verified_indices=sorted(verified),blocked=sorted(blocked))
                write(jf,journal)
                token=_SCOPE.set(current);target_token=_TARGET.set(operation['section'])
                problem=None
                try:
                    operation_started=time.monotonic()
                    evidence=apply(path,operation,manifest['track_changes'])
                    evidence["elapsed_ms"]=round((time.monotonic()-operation_started)*1000,2)
                except VerifyError as exc:
                    code=exc.envelope.error_code.value
                    problem={'status':'conflict' if code in REFUSED else 'unknown','code':code,'reason':exc.envelope.message}
                except Exception as exc:  # noqa: BLE001 - persist uncertain outcomes, then attempt independent targets
                    problem={'status':'unknown','reason':f'{type(exc).__name__}: {exc}'}
                finally:
                    _SCOPE.reset(token);_TARGET.reset(target_token)
                if problem:
                    deferred[key]=problem
                    if problem['status']=='unknown': unknown.add(index)
                    try:
                        observed=snapshot(session,slugs)
                        changed=changed_targets(current,observed)
                        # Preserve newly observed content, quarantine its targets, and
                        # continue elsewhere. Preconditions still come from the immutable manifest.
                        blocked.update(changed)
                        if changed:
                            problem['section_diff']='\n'.join(unified_diff(json.dumps(current,ensure_ascii=False,indent=2).splitlines(),json.dumps(observed,ensure_ascii=False,indent=2).splitlines(),fromfile='approved',tofile='current',lineterm=''))[:4000]
                        current=observed
                    except Exception as exc:  # noqa: BLE001 - unavailable transport cannot establish independent safe targets
                        journal.update(status='read_unavailable',problem=f'Live read unavailable: {exc}',working=current,blocked=sorted(blocked))
                        write(jf,journal)
                        return report(manifest,journal)
                    if operation['op']=='delete_paragraph' and problem['status']=='unknown': blocked.add(operation['section']+'|*')
                else:
                    current=project(current,operation);verified.add(index);deferred.pop(key,None)
                    journal['evidence'].append(dict(index=index,**evidence))
                journal.update(status='running',next=len(verified),verified_indices=sorted(verified),working=current,blocked=sorted(blocked))
                write(jf,journal)
            # Three passes bound retries while giving independent edits first priority.
        try: final=snapshot(session,slugs)
        except Exception as exc:  # noqa: BLE001 - report final verification unavailable without replay
            journal.update(status='verification_pending',problem=f'Final verification unavailable: {exc}')
            write(jf,journal);return report(manifest,journal)
        if final!=current:
            journal.update(status='partial',problem='Protected section changed during final verification')
        elif deferred or blocked:
            journal.update(status='unknown' if not verified and any(v['status']=='unknown' for v in deferred.values()) else 'partial',problem='Some operations need attention; independent operations were attempted')
        else: journal.update(status='complete',problem=None)
        journal.update(next=len(verified),verified_indices=sorted(verified),working=current,blocked=sorted(blocked))
        write(jf,journal)
        return report(manifest,journal)


def ensure_settled(root, own_run=None, sections=None):
    for jf in root.glob("*.journal.json"):
        if own_run and jf.name == f"{own_run}.journal.json":
            continue
        previous = json.loads(jf.read_text())
        if previous["status"] in {"dispatching", "unknown"} or any(v['status']=='unknown' for v in previous.get('deferred',{}).values()):
            other_manifest = jf.with_name(jf.name.replace(".journal.json", ".manifest.json"))
            other = json.loads(other_manifest.read_text()) if other_manifest.exists() else {}
            other_sections = {s["slug"] for s in other.get("sections", [])}
            if sections and other_sections and not set(sections) & other_sections:
                continue
            raise invalid("Another run has an uncertain outcome in an overlapping or unknown section; reconcile it before changing that section")


def reconcile(path,run_id,manifest_sha256,disposition,decision_ref,operation_index=None):
    path=str(paths.resolve_allowed_docx_path(path,must_exist=False))
    if not re.fullmatch(r'[0-9a-f]{32}',run_id) or disposition not in {'applied','not-applied'} or not decision_ref.strip():
        raise invalid('Reconciliation requires run_id, applied/not-applied disposition and decision_ref')
    root=folder(path)
    with locked(root):
        manifest=json.loads((root/f'{run_id}.manifest.json').read_text());jf=root/f'{run_id}.journal.json';journal=json.loads(jf.read_text())
        if sha(manifest)!=manifest_sha256 or journal['digest']!=manifest_sha256 or manifest['path']!=path: raise invalid('Manifest digest or document identity differs')
        deferred=journal.setdefault('deferred',{})
        candidates=[int(i) for i,v in deferred.items() if v['status']=='unknown']
        if journal['status']=='dispatching' and not candidates: candidates=[journal.get('active_index',journal['next'])]
        if not candidates and journal['status']=='unknown': candidates=[journal.get('active_index',journal['next'])]
        if operation_index is None:
            if len(candidates)!=1: raise invalid('Select operation_index for one uncertain operation')
            operation_index=candidates[0]
        if operation_index not in candidates: raise invalid('Operation has no uncertain outcome to reconcile')
        operation=manifest['operations'][operation_index]
        before=manifest['sections']
        for prior in manifest['operations'][:operation_index]: before=project(before,prior)
        after=project(before,operation)
        if operation["args"].get("replacement_segments") is not None:
            raise invalid("Formatted outcomes require independent formatting evidence; text alone cannot reconcile them")
        if after==before and operation['op']!='live_save': raise invalid('Unchanged text cannot prove a formatting-only write; both dispositions require operation-specific evidence')
        session=wm.live_session_for(path)
        if session.document_url!=manifest['document_url']: raise invalid('Connected document identity differs')
        captured=snapshot(session,[s['slug'] for s in manifest['sections']],details=True)
        if operation['op']=='live_save' and disposition=='applied' and captured.get('saved') is not True:
            raise invalid('Section text cannot prove a save; Word must report saved=true')
        expected=after if disposition=='applied' else before
        if operation['op']=='delete_paragraph':
            supported=next(s for s in captured['sections'] if s['slug']==operation['section'])==next(s for s in expected if s['slug']==operation['section'])
        else: supported=all(target_value(captured['sections'],t)==target_value(expected,t) for t in manifest['targets'][operation_index])
        if not supported:
            raise invalid('Live target does not support the requested disposition; manual reconciliation required')
        verified=set(journal.get('verified_indices',range(journal['next'])))
        if disposition=='applied':
            verified.add(operation_index);journal['evidence'].append({'index':operation_index,'applied':True,'verified_via':'lead-target-readback','decision_ref':decision_ref})
        deferred.pop(str(operation_index),None)
        blocked=set(journal.get('blocked',[]));blocked.difference_update(manifest['targets'][operation_index])
        journal.setdefault('decisions',[]).append({'index':operation_index,'disposition':disposition,'decision_ref':decision_ref})
        journal.update(status='partial',problem=None,working=captured['sections'],verified_indices=sorted(verified),next=len(verified),blocked=sorted(blocked))
        write(jf,journal);return report(manifest,journal)
