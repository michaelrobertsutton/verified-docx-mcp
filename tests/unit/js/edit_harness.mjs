// Behavioral model of ranges, formatting and comments. Mutations queue until sync.
import fs from 'node:fs';
import {DOMParser,EMPTY} from './shape_dom.mjs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
export function fixture(path, initial='Heading. plain sentence', commented=true) {
  const state={paragraphXml:EMPTY,text:initial, chars:Array.from(initial,(_,i)=>({bold:i<8,italic:false,
    underline:'None',strikeThrough:false,color:'#000000',name:'Calibri',size:11,doubleStrikeThrough:false,subscript:false,superscript:false})),comments:[],queue:[]};
  const ranges=(start,end)=>{
    const font={load() {},reset() {state.queue.push(()=>state.chars.slice(start,end).forEach(c=>Object.assign(c,{bold:false,italic:false,underline:"None"})));}};
    for(const key of ['bold','italic','underline','strikeThrough','color','name','size','doubleStrikeThrough','subscript','superscript'])
      Object.defineProperty(font,key,{get(){const vals=state.chars.slice(start,end).map(c=>c[key]);return vals.every(v=>v===vals[0])?vals[0]:null;},set(v){state.queue.push(()=>state.chars.slice(start,end).forEach(c=>c[key]=v));}});
    const range={load(){},font,getOoxml:()=>({value:state.paragraphXml}),get text(){return state.text.slice(start,end);},
      paragraphs:{items:[{text:state.text,getOoxml:()=>({value:state.paragraphXml})}],load(){}},getTrackedChanges:()=>({items:[],load(){}}),
      getComments:()=>({get items(){return state.comments;},load(){}}),
      compareLocationWith(other){return {value:end<=other.start?'Before':start>=other.end?'After':'Overlaps'};},
      start,end,
      search(find){const items=[],text=state.text.slice(start,end);for(let i=text.indexOf(find);i>=0;i=text.indexOf(find,i+find.length))items.push(ranges(start+i,start+i+find.length));return {items,load(){}};},
      insertText(text){const inherited={...state.chars[Math.max(0,start-1)]};
        state.queue.push(()=>{state.text=state.text.slice(0,start)+text+state.text.slice(end);
          state.chars.splice(start,end-start,...Array.from(text,()=>({...inherited})));
          state.comments=state.comments.filter(c=>!(c.start>=start&&c.end<=end));});
        return ranges(start,start+text.length);}};
    return range;
  };
  if(commented)state.comments.push({id:'c1',content:'question',authorName:'A',resolved:false,
    creationDate:'d',start:9,end:23,load(){},getRange(){return ranges(this.start,this.end);},
    replies:{items:[{id:'r1',content:'reply',authorName:'B'}],load(){}}});
  const body={load(){},get text(){return state.text;},getComments:()=>({get items(){return state.comments;},load(){}}),
    search(find){let items=[];for(let i=state.text.indexOf(find);i>=0;i=state.text.indexOf(find,i+find.length))items.push(ranges(i,i+find.length));return {items,load(){}};}};
  const context={document:{body,changeTrackingMode:'Off',load(){}},sync:async()=>{const q=state.queue.splice(0);q.forEach(fn=>fn());}};
  const sandbox={crypto:webcrypto,TextEncoder,console,DOMParser,
    document:{getElementById:()=>({addEventListener(){}})},
    Office:{onReady(){},context:{requirements:{isSetSupported:()=>true}}},
    Word:{run:(objects,fn)=>(fn||objects)(context),InsertLocation:{replace:'Replace'},
      ChangeTrackingMode:{trackAll:'TrackAll'},UnderlineType:{single:'Single',none:'None'}}};
  vm.createContext(sandbox);vm.runInContext(fs.readFileSync(path,'utf8')+'\nglobalThis.api={opReplace,opReplacementFont,opFormat,guardComments,opTextboxesList,opTextboxesRead,opScopeDescribe};',sandbox);
  return {state,context,sandbox,ranges,body};
}
if(process.argv[1].endsWith('edit_harness.mjs')) {
  const f=fixture(process.argv[2]);
  await assert.rejects(()=>f.sandbox.api.opReplace({find:'plain sentence',replace:'new text',expected_matches:1}),e=>e.code==='WOULD_DELETE_COMMENTS');
  assert.equal(f.state.text,'Heading. plain sentence');
  const result=await f.sandbox.api.opReplace({find:'plain sentence',replace:'new text',expected_matches:1,allow_comment_loss:true});
  assert.equal(result.comments_removed[0].replies[0].content,'reply');
  assert.equal(f.state.text,'Heading. new text');
  const partial=fixture(process.argv[2]);
  await assert.rejects(()=>partial.sandbox.api.opReplace({find:'plain',replace:'new',expected_matches:1}),e=>e.code==='WOULD_DELETE_COMMENTS');
  // issue #82: per-thread consent (allow_comment_loss_ids)
  const edit={find:'plain sentence',replace:'new text',expected_matches:1};
  const ids=fixture(process.argv[2]);
  await assert.rejects(()=>ids.sandbox.api.opReplace({...edit,allow_comment_loss_ids:[]}),e=>e.code==='WOULD_DELETE_COMMENTS'&&/not in allow_comment_loss_ids/.test(e.message));
  await assert.rejects(()=>ids.sandbox.api.opReplace({...edit,allow_comment_loss_ids:[{comment_id:'other',reply_count:1}]}),e=>e.code==='WOULD_DELETE_COMMENTS');
  await assert.rejects(()=>ids.sandbox.api.opReplace({...edit,allow_comment_loss_ids:[{comment_id:'c1',reply_count:1}]}),e=>e.code==='WOULD_DELETE_COMMENTS'&&/not resolved/.test(e.message));
  ids.state.comments[0].resolved=true;
  await assert.rejects(()=>ids.sandbox.api.opReplace({...edit,allow_comment_loss_ids:[{comment_id:'c1',reply_count:0}]}),e=>e.code==='WOULD_DELETE_COMMENTS'&&/reply/.test(e.message));
  await assert.rejects(()=>ids.sandbox.api.opReplace({...edit,allow_comment_loss:true,allow_comment_loss_ids:[{comment_id:'c1',reply_count:1}]}),e=>e.code==='INVALID_INPUT');
  assert.equal(ids.state.text,'Heading. plain sentence');
  const exact=await ids.sandbox.api.opReplace({...edit,allow_comment_loss_ids:[{comment_id:'c1',reply_count:1}]});
  assert.equal(exact.comments_removed[0].id,'c1');
  assert.equal(ids.state.text,'Heading. new text');
  console.log('comment preservation passed');
}
