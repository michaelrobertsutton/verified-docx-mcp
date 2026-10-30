// Issue #108: replacement_font (read-only first-run font) and replace's FORMAT_DRIFT check.
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {fixture} from './edit_harness.mjs';
const path=process.argv[2];
const CALIBRI={bold:false,italic:false,underline:'None',strikeThrough:false,color:'#000000',name:'Calibri',size:11,
  doubleStrikeThrough:false,subscript:false,superscript:false};

// The effective font of the FIRST character, even when the whole match is mixed.
const f=fixture(path,'Heading. plain sentence',false);
const plain=await f.sandbox.api.opReplacementFont({find:'plain',expected_matches:1});
assert.deepEqual({...plain.font},CALIBRI);
assert.equal(plain.match_count,1);
const mixed=await f.sandbox.api.opReplacementFont({find:'Heading. plain',expected_matches:1});
assert.equal(mixed.font.bold,true);
await assert.rejects(()=>f.sandbox.api.opReplacementFont({find:'absent',expected_matches:1}),e=>e.code==='zero_match');
await assert.rejects(()=>f.sandbox.api.opReplacementFont({find:'e',expected_matches:1}),e=>e.code==='match_count_mismatch');
assert.equal(f.state.text,'Heading. plain sentence');  // read-only

// Matches whose first runs differ refuse instead of picking one.
const two=fixture(path,'ab ab',false);
two.state.chars[3].bold=false;  // the fixture bolds the first 8 characters
await assert.rejects(()=>two.sandbox.api.opReplacementFont({find:'ab',expected_matches:2}),
  e=>e.code==='INVALID_INPUT'&&/specify every field/.test(e.message));
two.state.chars[3].bold=true;
assert.equal((await two.sandbox.api.opReplacementFont({find:'ab',expected_matches:2})).match_count,2);

// replace honours expected_first_font: unchanged -> writes; different -> FORMAT_DRIFT, nothing written.
const ok=fixture(path,'Heading. plain sentence',false);
await ok.sandbox.api.opReplace({find:'plain sentence',replace:'new',expected_matches:1,expected_first_font:CALIBRI});
assert.equal(ok.state.text,'Heading. new');
const drift=fixture(path,'Heading. plain sentence',false);
await assert.rejects(()=>drift.sandbox.api.opReplace({find:'plain sentence',replace:'new',expected_matches:1,
  expected_first_font:{...CALIBRI,size:14}}),e=>e.code==='FORMAT_DRIFT');
assert.equal(drift.state.text,'Heading. plain sentence');
await assert.rejects(()=>drift.sandbox.api.opReplace({find:'plain sentence',replace:'new',expected_matches:1,
  expected_first_font:{bold:false}}),e=>e.code==='INVALID_INPUT');

// A reformat after the preflight reads but before the write is caught by the last preflight sync.
const late=fixture(path,'Heading. plain sentence',false);
vm.runInContext(`let lateChecks=0; sectionGuard=async()=>({check:async()=>{if(++lateChecks===2) globalThis.reformat();}});`,late.sandbox);
late.sandbox.reformat=()=>{late.state.chars[9].italic=true;};
await assert.rejects(()=>late.sandbox.api.opReplace({find:'plain sentence',replace:'new',expected_matches:1,
  expected_first_font:CALIBRI}),e=>e.code==='FORMAT_DRIFT');
assert.equal(late.state.text,'Heading. plain sentence');
console.log('partial font baseline and drift refusal passed');
