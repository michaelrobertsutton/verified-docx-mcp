import assert from 'node:assert/strict';
import {fixture} from './edit_harness.mjs';
const path=process.argv[2];
const f=fixture(path,'body text',false);
const box=fixture(path,'callout text',false);
const empty=fixture(path,'',false);
// Shape bodies expose their comment marks through OOXML and their paragraphs, as real hosts do.
// Range.search on shape text freezes Word for Mac, so the pane must never call it.
const noSearch=()=>{throw new Error('Range.search must never run on shape text');};
const paragraphsOf=shape=>({get items(){let off=0;return shape.state.text.split('\r').map(t=>{const s=off;off+=t.length+1;
  return {text:t,getRange:()=>shape.ranges(s,s+t.length)};});},load(){}});
for(const shape of [box,empty]){
  shape.body.getOoxml=()=>({value:'<pkg:package>'+shape.state.comments.map(c=>`<w:commentRangeStart w:id="${c.id}"/>`).join('')+'</pkg:package>'});
  shape.body.search=noSearch;
  shape.body.paragraphs=paragraphsOf(shape);
}
f.sandbox.Word.ShapeType={textBox:'TextBox',geometricShape:'GeometricShape',group:'Group',canvas:'Canvas'};
let items=[{id:1,name:'callout',type:'TextBox',body:box.body},{id:2,name:'empty',type:'TextBox',body:empty.body}];
f.body.shapes={get items(){return items;},load(){}};
const sync=f.context.sync;
f.context.sync=async()=>{await sync();await box.context.sync();await empty.context.sync();};
const listing=await f.sandbox.api.opTextboxesList();
assert.equal(listing.textboxes.length,2);assert.equal(listing.textboxes[1].text,'');
const handle=listing.textboxes[0].textbox_id;
const before=await f.sandbox.api.opScopeDescribe({scope:handle});
// A whole paragraph with mixed formatting refuses under the default policy, with no mutation.
await assert.rejects(()=>f.sandbox.api.opReplace({scope:handle,expectedScopeSha256:before.scopeSha256,
  find:'callout text',replace:'updated text',expected_matches:1}),e=>e.code==='MIXED_FORMATTING');
assert.equal(box.state.text,'callout text');
const r=await f.sandbox.api.opReplace({scope:handle,expectedScopeSha256:before.scopeSha256,
  find:'callout text',replace:'updated text',expected_matches:1,inherit_format:'previous'});
assert.equal(box.state.text,'updated text');assert.notEqual(r.scope_post,before.scopeSha256);
await assert.rejects(()=>f.sandbox.api.opReplace({scope:handle,expectedScopeSha256:before.scopeSha256,
  find:'updated text',replace:'stale',expected_matches:1}),e=>e.code==='LIVE_STALE');
assert.equal(box.state.text,'updated text');
await assert.rejects(()=>f.sandbox.api.opTextboxesRead({scope:'textbox:old:1'}),e=>e.code==='LIVE_STALE');
items.push({id:3,type:'Group'});
assert.equal((await f.sandbox.api.opTextboxesList()).coverage,'partial');
await assert.rejects(()=>f.sandbox.api.opReplace({scope:'all',find:'text',replace:'x',expected_matches:2}),e=>e.code==='LIVE_CAPABILITY_MISSING');
items=[{id:1,name:'broken',type:'TextBox',body:{load(){},text:undefined}}];
await assert.rejects(()=>f.sandbox.api.opTextboxesList(),e=>e.code==='HOST_SHAPE_READ_FAILED');
// A host that cannot search shape text must say so, not report "found 0".
const blind=fixture(path,'blind text',false);
blind.body.search=noSearch;
// Word for Mac 16.113.3: every comment lookup inside shape text throws GeneralException, and the
// shape body can only be inspected for comments through its OOXML.
const ghost=()=>{throw Object.assign(new Error('GeneralException'),{code:'GeneralException'});};
const macRange=(s,e)=>{const r=blind.ranges(s,e);r.getComments=ghost;return r;};
blind.body.getComments=ghost;
let shapeOoxmlReadable=true;
blind.body.getOoxml=()=>{if(!shapeOoxmlReadable)throw new Error('getOoxml failed');
  return {value:'<pkg:package>'+blind.state.comments.map(c=>`<w:commentRangeStart w:id="${c.id}"/>`).join('')+'</pkg:package>'};};
blind.body.paragraphs={get items(){return [{text:blind.state.text,getRange:()=>macRange(0,blind.state.text.length)}];},load(){}};
const blindSync=f.context.sync;
f.context.sync=async()=>{await blindSync();await blind.context.sync();};
items=[{id:9,name:'blind',type:'TextBox',body:blind.body}];
const blindHandle=(await f.sandbox.api.opTextboxesList()).textboxes[0].textbox_id;
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'blind',replace:'x',expected_matches:1}),
  e=>e.code==='LIVE_CAPABILITY_MISSING'&&/cannot search inside blind/.test(e.message));
// Whole-paragraph fallback uses the normal guards and read-back path.
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'blind text',replace:'x',expected_matches:2}),
  e=>e.code==='LIVE_OP_FAILED'&&/expected 2/.test(e.message));
assert.equal(blind.state.text,'blind text');
const whole=await f.sandbox.api.opReplace({scope:blindHandle,find:'blind text',replace:'updated shape',expected_matches:1,inherit_format:'previous'});
assert.equal(whole.match_count,1);assert.equal(blind.state.text,'updated shape');
assert.equal(f.state.text,'body text');
const formatted=await f.sandbox.api.opFormat({scope:blindHandle,find:'updated shape',expected_matches:1,bold:true});
assert.equal(formatted.matches[0].boldAfter,true);
blind.state.comments=[{id:'inside',start:0,end:13,load(){},getRange(){return blind.ranges(0,13);},replies:{items:[],load(){}}}];
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'updated shape',replace:'no',expected_matches:1,inherit_format:'previous'}),
  e=>e.code==='WOULD_DELETE_COMMENTS');
await assert.rejects(()=>f.sandbox.api.opFormat({scope:blindHandle,find:'updated shape',expected_matches:1,bold:false}),
  e=>e.code==='WOULD_DELETE_COMMENTS');
assert.equal(blind.state.text,'updated shape');
// allow_comment_loss is honoured on shapes through the same OOXML source, and reports what went.
const allowed=await f.sandbox.api.opReplace({scope:blindHandle,find:'updated shape',replace:'allowed',expected_matches:1,inherit_format:'previous',allow_comment_loss:true});
assert.equal(JSON.stringify(allowed.comments_removed.map(c=>c.id)),'["inside"]');
assert.equal(blind.state.text,'allowed');
blind.state.text='updated shape';
blind.state.comments=[];
// A shape body whose OOXML cannot be read fails closed: no mutation, a typed reason.
shapeOoxmlReadable=false;
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'updated shape',replace:'no',expected_matches:1,inherit_format:'previous'}),
  e=>e.code==='LIVE_CAPABILITY_MISSING'&&/Cannot verify comments in shape text/.test(e.message));
await assert.rejects(()=>f.sandbox.api.opFormat({scope:blindHandle,find:'updated shape',expected_matches:1,bold:false}),
  e=>e.code==='LIVE_CAPABILITY_MISSING');
assert.equal(blind.state.text,'updated shape');
shapeOoxmlReadable=true;
const fresh=await f.sandbox.api.opScopeDescribe({scope:blindHandle});
blind.state.text='changed shape';
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,expectedScopeSha256:fresh.scopeSha256,
  find:'changed shape',replace:'no',expected_matches:1}),e=>e.code==='LIVE_STALE');
blind.state.text='updated shape';
// Do not address only the whole-paragraph occurrence while omitting a substring.
blind.state.text='updated\rupdated shape';
blind.body.paragraphs={items:[{text:'updated',getRange:()=>macRange(0,7)}],load(){}};
await assert.rejects(()=>f.sandbox.api.opFormat({scope:blindHandle,find:'updated',expected_matches:1,bold:false}),
  e=>e.code==='LIVE_CAPABILITY_MISSING');
blind.state.text='updated shape';
blind.body.paragraphs={get items(){return [{text:blind.state.text,getRange:()=>macRange(0,blind.state.text.length)}];},load(){}};
// A host that ignores a replacement must fail read-back verification.
blind.body.paragraphs={items:[{text:'updated shape',getRange(){const range=macRange(0,13);range.insertText=()=>macRange(0,13);return range;}}],load(){}};
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'updated shape',replace:'ignored',expected_matches:1,inherit_format:'previous'}),
  e=>e.code==='VERIFICATION_FAILED');
blind.body.paragraphs={get items(){return [{text:blind.state.text,getRange:()=>macRange(0,blind.state.text.length)}];},load(){}};
const getRange=blind.body.paragraphs.items[0].getRange;
blind.body.paragraphs={items:[{text:'updated shape',getRange(){const range=getRange();range.getTrackedChanges=()=>({items:[{}],load(){}});return range;}}],load(){}};
await assert.rejects(()=>f.sandbox.api.opReplace({scope:blindHandle,find:'updated shape',replace:'no',expected_matches:1}),
  e=>e.code==='TRACKED_CHANGES_PRESENT');
console.log('text boxes passed');
