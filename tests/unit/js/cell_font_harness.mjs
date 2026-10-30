// Pane-side checks for issue #107's pure pieces: the inherited cell font (apply + readback),
// the refusal of tracked merges, and the advertised capabilities. Word's own font, range and
// OOXML-import behaviour needs real Word (docs/acceptance/107-cell-editing.md).
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
const sandbox={crypto:webcrypto,TextEncoder,console,DOMParser:class{},document:{getElementById:()=>({addEventListener(){}})},
  Office:{onReady(){},context:{requirements:{isSetSupported:()=>true}}},Word:{run:fn=>fn({})}};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2],'utf8')+
  '\nglobalThis.t={applyInheritedFont,verifyInheritedFont,opCellParagraphMerge,PANE_CAPABILITIES};',sandbox);
const {applyInheritedFont,verifyInheritedFont,opCellParagraphMerge,PANE_CAPABILITIES}=sandbox.t;
const context={sync:async()=>{}};
const cell=(font={})=>({body:{font:{name:'Calibri',size:10.5,color:'#000000',load(){},...font}}});

// apply: only the keys the server sent; a bare hex colour gets its '#'.
let c=cell();
applyInheritedFont(c,{name:'Arial',size:8.5,color:'1F3864'});
assert.equal(c.body.font.name,'Arial');assert.equal(c.body.font.size,8.5);assert.equal(c.body.font.color,'#1F3864');
c=cell();applyInheritedFont(c,{size:8.5});
assert.equal(c.body.font.name,'Calibri');assert.equal(c.body.font.color,'#000000');assert.equal(c.body.font.size,8.5);
c=cell();applyInheritedFont(c,{name:null,size:'',color:undefined});
assert.equal(c.body.font.size,10.5);

// readback: matches (colour compared without '#' and case) or VERIFICATION_FAILED.
await verifyInheritedFont(context,cell({name:'Arial',size:8.5,color:'#1f3864'}),{name:'Arial',size:8.5,color:'1F3864'});
await verifyInheritedFont(context,cell(),{});
await assert.rejects(()=>verifyInheritedFont(context,cell({size:10.5}),{size:8.5}),
  e=>e.code==='VERIFICATION_FAILED'&&/size/.test(e.message));
await assert.rejects(()=>verifyInheritedFont(context,cell({name:'Calibri'}),{name:'Arial'}),e=>e.code==='VERIFICATION_FAILED');
await assert.rejects(()=>verifyInheritedFont(context,cell({color:'#000000'}),{color:'1F3864'}),e=>e.code==='VERIFICATION_FAILED');

// a tracked merge is refused before Word is touched.
await assert.rejects(()=>opCellParagraphMerge({track_changes:true}),e=>e.code==='INVALID_INPUT');

for(const cap of ['cell_font_inheritance_v1','cell_paragraph_merge_v1','paragraph_positions_v1'])
  assert.ok(PANE_CAPABILITIES.includes(cap),cap);
console.log('cell font harness passed');
