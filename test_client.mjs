import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';

// Exercise the registered component through its public props and event handlers.
let registration, current, selectionValue=null;
const listeners=new Map();
const document={activeElement:null,addEventListener:(name,fn)=>listeners.set(name,fn),removeEventListener:(name,fn)=>{if(listeners.get(name)===fn)listeners.delete(name);},execCommand:()=>{}};
const window={innerWidth:1200,innerHeight:900,__ModuleLoader__:{load(value){registration=value;}},confirm:()=>false,getSelection:()=>selectionValue,addEventListener:(name,fn)=>listeners.set(name,fn),removeEventListener:(name,fn)=>{if(listeners.get(name)===fn)listeners.delete(name);}};
const React={Fragment:'fragment',createElement:(type,props,...children)=>({type,props:props??{},children:children.flat(Infinity).filter(x=>x!==null&&x!==undefined&&x!==false)}),useState(value){return current.state(value);},useRef(value){return current.ref(value);},useEffect(fn,deps){current.effect(fn,deps);}};
vm.runInNewContext(readFileSync(new URL('./client.js',import.meta.url),'utf8'),{window,document,fetch:()=>{throw Error('Unexpected fetch');},console});
let Canvas;
registration.factory(()=>React).apply({effect:fn=>fn(),locale:{register(){},bind:()=>key=>key},documentPreviews:{register(){}},slots:{inject:(_name,fn)=>fn(),register:(_spec,component)=>{Canvas=component;}}});
class Harness {
  constructor(component,props){this.component=component;this.props=props;this.hooks=[];this.pending=[];this.dirty=true;}
  state(value){const i=this.index++;if(!(i in this.hooks))this.hooks[i]=typeof value==='function'?value():value;return [this.hooks[i],next=>{const v=typeof next==='function'?next(this.hooks[i]):next;if(!Object.is(v,this.hooks[i])){this.hooks[i]=v;this.dirty=true;}}];}
  ref(value){const i=this.index++;return this.hooks[i]??(this.hooks[i]={current:value});}
  effect(fn,deps){const i=this.index++,old=this.hooks[i];if(!old||deps.some((v,n)=>!Object.is(v,old.deps[n]))){this.hooks[i]={deps,cleanup:old?.cleanup};this.pending.push(()=>{this.hooks[i].cleanup?.();this.hooks[i].cleanup=fn();});}}
  flush(){for(let n=0;this.dirty;n++){assert.ok(n<30,'render loop');this.dirty=false;this.index=0;current=this;this.tree=this.component(this.props);current=null;for(const effect of this.pending.splice(0))effect();}return this.tree;}
  async settle(){for(let i=0;i<10;i++){await Promise.resolve();this.flush();}return this.tree;}
  change(props){this.props={...this.props,...props};this.dirty=true;this.flush();}
  dispose(){for(const hook of this.hooks)hook?.cleanup?.();}
}
function nodes(tree,predicate){if(!tree||typeof tree!=='object')return [];return [...(predicate(tree)?[tree]:[]),...tree.children.flatMap(child=>nodes(child,predicate))];}
const clone=value=>JSON.parse(JSON.stringify(value));
const run={id:'r:0:0',text:'Original',bold:false};
const doc={kind:'docx',path:'C:/workspace/report.docx',hash:'hash-original',paragraphs:[{id:'p:0',text:run.text,runs:[run]}]};
const deck={kind:'pptx',path:'C:/workspace/deck.pptx',hash:'deck-hash',width:9144000,height:6858000,slides:[{id:'slide:0',elements:[{id:'shape:0:2',kind:'text',x:914400,y:914400,width:1828800,height:914400,editable:true,movable:true,paragraphs:[{id:'p:0:2:0',text:run.text,runs:[{...run,id:'r:0:2:0:0'}]}]}]},{id:'slide:1',elements:[]}]};
function deferred(){let resolve;const promise=new Promise(r=>{resolve=r;});return {promise,resolve};}
async function setup(model=doc,custom){const calls=[],drafts=[];let revision=1;const props={sessionId:'chat-session',resourceAddress:'dsh-resource://file/session/file-session/'+(model.kind==='pptx'?'deck.pptx':'report.docx'),content:{kind:'renderer',revision,loaded(){},failed(){}},invoke:async(sessionId,request)=>{calls.push({sessionId,request:clone(request)});if(custom){const value=custom(request);if(value!==undefined)return value;}if(request.action==='read')return {ok:true,value:{...clone(model),path:request.path.includes('-edited')?request.path:model.path}};if(request.action==='save')return {ok:true,value:{path:model.path.replace(/\.(docx|pptx)$/,'-edited.$1'),hash:'saved'}};return {ok:true,value:{pages:[{index:1,width:900,height:700,data:'data:image/png;base64,'}],pageCount:2}};},t:key=>key,useTabInfo:()=>({tab:{actions:{openResource:address=>drafts.push(address)}}}),inputActions:{captureInsertion:()=>({}),insertText:text=>{drafts.push(text);return true;},submit(){}},scrollportRef:{current:null}};const app=new Harness(Canvas,props);app.flush();await app.settle();return {app,calls,drafts,refresh(){app.change({content:{...props.content,revision:++revision}});}};}
function button(app,label){const result=nodes(app.tree,n=>n.type==='button'&&n.children.includes(label));assert.equal(result.length,1,'button '+label);return result[0];}
function click(app,label){const b=button(app,label);assert.ok(!b.props.disabled,label+' enabled');b.props.onClick();app.flush();}
function edit(app,text){const node=nodes(app.tree,n=>typeof n.type==='function')[0];node.props.onFocus();node.props.onCommit(text);app.flush();return node;}
function pageInput(app){return nodes(app.tree,n=>n.type==='input'&&n.props.min===1&&n.props.max!==200)[0];}
function positionInput(app,key){const label=nodes(app.tree,n=>n.type==='label'&&n.children[0]===key)[0];return label.children[1];}
let passed=0;
async function test(name,fn){await fn();passed++;console.log('ok '+passed+' - '+name);}
await test('focused input becomes dirty; text key stays stable and undo groups typing',async()=>{const {app}=await setup();const initial=nodes(app.tree,n=>typeof n.type==='function')[0];edit(app,'First');const next=nodes(app.tree,n=>typeof n.type==='function')[0];assert.equal(initial.props.key,next.props.key);assert.equal(button(app,'save').props.disabled,false);assert.ok(listeners.has('beforeunload'));next.props.onCommit('Second');app.flush();click(app,'undo');assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Original');assert.equal(button(app,'undo').props.disabled,true);click(app,'redo');assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Second');app.dispose();});
await test('automatic host refresh preserves unsaved operations',async()=>{const {app,calls,refresh}=await setup();edit(app,'Pending');const count=calls.length;refresh();await app.settle();assert.equal(calls.length,count);assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Pending');assert.equal(button(app,'save').props.disabled,false);app.dispose();});
await test('explicit reload cancellation preserves edits, acceptance clears',async()=>{const {app}=await setup();edit(app,'Pending');click(app,'reload');assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Pending');window.confirm=()=>true;click(app,'reload');await app.settle();assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Original');window.confirm=()=>false;app.dispose();});
await test('save and reload use file session and saved copy path',async()=>{const {app,calls,drafts}=await setup();edit(app,'Changed');click(app,'save');await app.settle();click(app,'reload');await app.settle();assert.ok(calls.every(c=>c.sessionId==='file-session'));assert.equal(calls.at(-1).request.path,'C:/workspace/report-edited.docx');edit(app,'Again');click(app,'save');await app.settle();click(app,'open');assert.ok(drafts[0].includes('/session/file-session/'));app.dispose();});
await test('undo and redo synchronize position inspector',async()=>{const {app}=await setup(deck);const obj=nodes(app.tree,n=>n.props.className?.startsWith('oc-object'))[0];obj.getBoundingClientRect=()=>({left:10,top:10,bottom:30});obj.props.onClick({target:obj,currentTarget:obj});app.flush();positionInput(app,'x').props.onChange({target:{value:'2'}});app.flush();click(app,'apply');assert.equal(positionInput(app,'x').props.value,2);click(app,'undo');assert.equal(positionInput(app,'x').props.value,1);click(app,'redo');assert.equal(positionInput(app,'x').props.value,2);app.dispose();});
await test('cancelled object drag restores committed position',async()=>{const {app}=await setup(deck);let handle=button(app,'↔');const parentElement={style:{}};const currentTarget={parentElement,setPointerCapture(){}};handle.props.onPointerDown({preventDefault(){},currentTarget,pointerId:1,clientX:0,clientY:0});app.flush();handle=button(app,'↔');handle.props.onPointerMove({currentTarget,clientX:90,clientY:90});assert.equal(parentElement.style.left,'180px');handle.props.onPointerCancel({currentTarget});app.flush();assert.equal(parentElement.style.left,'90px');assert.equal(positionInput(app,'x').props.value,1);assert.equal(button(app,'save').props.disabled,true);app.dispose();});
await test('old preview cannot overwrite refreshed resource',async()=>{const held=deferred();const {app,refresh}=await setup(doc,req=>req.action==='preview'?held.promise:undefined);click(app,'preview');refresh();await app.settle();held.resolve({ok:true,value:{pages:[{index:1,width:900,height:700,data:'stale'}],pageCount:99}});await app.settle();assert.equal(nodes(app.tree,n=>n.type==='img').length,0);app.dispose();});
await test('old save response cannot clear preserved edits after refresh',async()=>{const held=deferred();const {app,refresh}=await setup(doc,req=>req.action==='save'?held.promise:undefined);edit(app,'Keep me');click(app,'save');refresh();await app.settle();held.resolve({ok:true,value:{path:'old-edited.docx'}});await app.settle();click(app,'edit');assert.equal(nodes(app.tree,n=>typeof n.type==='function')[0].props.text,'Keep me');assert.equal(button(app,'save').props.disabled,false);app.dispose();});
await test('preview selection clears on edit switch',async()=>{const {app}=await setup();click(app,'preview');await app.settle();const frame=nodes(app.tree,n=>n.props.className==='oc-frame')[0];frame.props.ref.current={getBoundingClientRect:()=>({left:0,top:0,width:100,height:100})};frame.props.onPointerDown({button:0,preventDefault(){},currentTarget:{setPointerCapture(){}},pointerId:1,clientX:10,clientY:10});frame.props.onPointerMove({clientX:40,clientY:40});frame.props.onPointerUp({clientX:40,clientY:40});app.flush();assert.equal(nodes(app.tree,n=>n.props.className==='oc-floating').length,1);click(app,'edit');assert.equal(nodes(app.tree,n=>n.props.className==='oc-floating').length,0);app.dispose();});
await test('editable DOM receives input immediately and no React text children',async()=>{const {app}=await setup();const outer=nodes(app.tree,n=>typeof n.type==='function')[0];let changed;const child=new Harness(outer.type,{...outer.props,onCommit:value=>{changed=value;},onEnd(){}});child.flush();assert.equal(child.tree.children.length,0);child.tree.props.onInput({currentTarget:{innerText:'Typing'}});assert.equal(changed,'Typing');child.dispose();app.dispose();});
await test('keyboard selection is document-owned and cleared on reload',async()=>{
  const {app,drafts}=await setup();
  const root={contains:node=>node===runNode||node===paragraphNode};
  const paragraphNode={nodeType:1,parentElement:null,contains:node=>node===runNode||node===paragraphNode,getAttribute:()=> 'p:0',closest(){return this;}};
  const runNode={nodeType:1,parentElement:paragraphNode,contains:node=>node===runNode,getAttribute:()=> 'r:0:0',closest(){return this;}};
  app.tree.props.ref.current=root;
  selectionValue={anchorNode:runNode,focusNode:runNode,rangeCount:1,getRangeAt:()=>({getBoundingClientRect:()=>({left:10,top:10,bottom:30})}),toString:()=> 'Original passage'};
  listeners.get('selectionchange')();app.flush();
  click(app,'request');nodes(app.tree,n=>n.type==='textarea')[0].props.onChange({target:{value:'Rewrite selection'}});app.flush();
  const arrow=nodes(app.tree,n=>n.props.className==='oc-arrow')[0];arrow.props.onClick();app.flush();click(app,'draft');await app.settle();assert.ok(drafts[0].includes('Original passage'));assert.ok(drafts[0].includes('r:0:0'));
  const outside={nodeType:1,closest:()=>null};
  selectionValue={anchorNode:outside,focusNode:outside,toString:()=> 'Toolbar text'};
  listeners.get('selectionchange')();app.flush();
  assert.equal(nodes(app.tree,n=>n.props.className==='oc-floating').length,0);
  assert.equal(drafts.length,1,'outside selection cannot enqueue toolbar text');
  click(app,'reload');await app.settle();
  assert.equal(nodes(app.tree,n=>n.props.className==='oc-comment-dock').length,0);
  assert.equal(nodes(app.tree,n=>n.type==='textarea').length,0);
  selectionValue=null;app.dispose();
});
console.log(`${passed} client regression tests passed`);
