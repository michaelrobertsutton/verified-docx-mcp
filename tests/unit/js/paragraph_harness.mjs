import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
let texts=['previous','target','next',''];
let trailer=null; let mode='Off', table=false, anchored=false, section=false, malformed=false, comments=false, revisions=false, broken=false;
const collection={load(){},get items(){return texts.map((text,i)=>({text,tableNestingLevel:table&&i===1?1:0,
  getOoxml:()=>({value:'xml'}),getRange:()=>({getComments:()=>({items:comments?[comment]:[],load(){}}),
    getTrackedChanges:()=>({items:revisions?[{}]:[],load(){}})}),delete(){if(!broken)texts.splice(i,1);}}));}};
const comment={id:'c',content:'question',authorName:'A',resolved:false,
  getRange:()=>({text:'target',load(){},compareLocationWith:()=>({value:'Equal'})}),replies:{items:[],load(){}}};
const context={document:{body:{load(){},get text(){return texts.join('\n');},paragraphs:collection},
  load(){},get changeTrackingMode(){return mode;},set changeTrackingMode(v){mode=v;}},sync:async()=>{}};
class DOMParser {parseFromString(){const p={textContent:'target',getElementsByTagNameNS:(ns,tag)=>(tag==='anchor'&&anchored)||(tag==='sectPr'&&section)?[{}]:[]};
  return {getElementsByTagName:tag=>malformed?[{}]:[],getElementsByTagNameNS:(ns,tag)=>tag==='body'?[{getElementsByTagNameNS:(n2,t2)=>t2==='p'?(trailer?[p,trailer]:[p]):((t2==='anchor'&&anchored)?[{}]:[])}]:[]};}}
const sandbox={crypto:webcrypto,TextEncoder,console,DOMParser,document:{getElementById:()=>({addEventListener(){}})},
  Office:{onReady(){},context:{requirements:{isSetSupported:()=>true}}},Word:{run:fn=>fn(context),RangeLocation:{whole:'Whole'},ChangeTrackingMode:{trackAll:'TrackAll'}}};
vm.createContext(sandbox);vm.runInContext(fs.readFileSync(process.argv[2],'utf8')+'\nglobalThis.op=opParagraphDelete;',sandbox);
for(const flag of ['table','anchored','section','malformed','comments','revisions']){
  if(flag==='table')table=true;if(flag==='anchored')anchored=true;if(flag==='section')section=true;
  if(flag==='malformed')malformed=true;if(flag==='comments')comments=true;if(flag==='revisions')revisions=true;
  await assert.rejects(()=>sandbox.op({anchor_text:'target',track_changes:true}));
  assert.equal(texts.length,4);assert.equal(mode,'Off');
  table=anchored=section=malformed=comments=revisions=false;
}
await assert.rejects(()=>sandbox.op({anchor_text:''}),e=>e.code==='STRUCTURAL_BOUNDARY');
texts=['same','same',''];await assert.rejects(()=>sandbox.op({anchor_text:'same'}));
texts=['previous','target','next',''];broken=true;
await assert.rejects(()=>sandbox.op({anchor_text:'target',track_changes:true}),e=>e.code==='VERIFICATION_FAILED');assert.equal(mode,'Off');broken=false;
const r=await sandbox.op({anchor_text:'target',track_changes:true});assert.equal(r.after_count,3);assert.deepEqual(texts,['previous','next','']);assert.equal(mode,'Off');
// Word appends one empty paragraph to a paragraph's OOXML; only that is tolerated.
const empty={textContent:'',getElementsByTagNameNS:()=>[]};
trailer={textContent:'stray text',getElementsByTagNameNS:()=>[]};
await assert.rejects(()=>sandbox.op({anchor_text:'next'}),e=>e.code==='STRUCTURAL_BOUNDARY'&&/Incomplete paragraph inspection/.test(e.message));
trailer={textContent:'',getElementsByTagNameNS:(ns,tag)=>tag==='drawing'?[{}]:[]};
await assert.rejects(()=>sandbox.op({anchor_text:'next'}),e=>e.code==='STRUCTURAL_BOUNDARY');
trailer=empty;texts=['previous','target','next',''];
const withTrailer=await sandbox.op({anchor_text:'target'});assert.equal(withTrailer.after_count,3);trailer=null;
console.log('paragraph deletion passed');
