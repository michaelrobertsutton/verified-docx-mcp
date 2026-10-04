import assert from 'node:assert/strict';
import {fixture} from './edit_harness.mjs';
const path=process.argv[2];
const f=fixture(path,'body text',false);
const box=fixture(path,'callout text',false);
const empty=fixture(path,'',false);
f.sandbox.Word.ShapeType={textBox:'TextBox',geometricShape:'GeometricShape',group:'Group',canvas:'Canvas'};
let items=[{id:1,name:'callout',type:'TextBox',body:box.body},{id:2,name:'empty',type:'TextBox',body:empty.body}];
f.body.shapes={get items(){return items;},load(){}};
const sync=f.context.sync;
f.context.sync=async()=>{await sync();await box.context.sync();await empty.context.sync();};
const listing=await f.sandbox.api.opTextboxesList();
assert.equal(listing.textboxes.length,2);assert.equal(listing.textboxes[1].text,'');
const handle=listing.textboxes[0].textbox_id;
const before=await f.sandbox.api.opScopeDescribe({scope:handle});
const r=await f.sandbox.api.opReplace({scope:handle,expectedScopeSha256:before.scopeSha256,
  find:'callout',replace:'updated',expected_matches:1});
assert.equal(box.state.text,'updated text');assert.notEqual(r.scope_post,before.scopeSha256);
await assert.rejects(()=>f.sandbox.api.opReplace({scope:handle,expectedScopeSha256:before.scopeSha256,
  find:'updated',replace:'stale',expected_matches:1}),e=>e.code==='LIVE_STALE');
assert.equal(box.state.text,'updated text');
await assert.rejects(()=>f.sandbox.api.opTextboxesRead({scope:'textbox:old:1'}),e=>e.code==='LIVE_STALE');
items.push({id:3,type:'Group'});
assert.equal((await f.sandbox.api.opTextboxesList()).coverage,'partial');
await assert.rejects(()=>f.sandbox.api.opReplace({scope:'all',find:'text',replace:'x',expected_matches:2}),e=>e.code==='LIVE_CAPABILITY_MISSING');
items=[{id:1,name:'broken',type:'TextBox',body:{load(){},text:undefined}}];
await assert.rejects(()=>f.sandbox.api.opTextboxesList(),e=>e.code==='HOST_SHAPE_READ_FAILED');
// A host that cannot search shape text must say so, not report "found 0".
const blind=fixture(path,'blind text',false);
blind.body.search=()=>({items:[],load(){}});
items=[{id:9,name:'blind',type:'TextBox',body:blind.body}];
const blindHandle=(await f.sandbox.api.opTextboxesList()).textboxes[0].textbox_id;
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'blind',replace:'x',expected_matches:1}),
  e=>e.code==='LIVE_CAPABILITY_MISSING'&&/cannot search inside blind/.test(e.message));
console.log('text boxes passed');
