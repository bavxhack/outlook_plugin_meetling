const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
class Event {
  listeners = new Set();
  addListener(fn) {this.listeners.add(fn);}
  removeListener(fn) {this.listeners.delete(fn);}
  async emit(...args) {return Promise.all([...this.listeners].map(fn=>fn(...args)));}
}
async function setup() {
  const vault = new Map(), requests = [], inserts = [], filters = new Map();
  const configuration = {baseUrl:'https://host.example/meetling/'};
  const browser = {
    runtime:{id:'test-addon',getURL:p=>'moz-extension://test/'+p,onMessage:new Event()},
    storage:{managed:{get:async()=>({})},local:{get:async()=>({configuration}),set:async()=>{}}},
    browserAction:{onClicked:new Event()},
    windows:{create:async()=>({id:5,tabs:[{id:10}]}),remove:async()=>{},update:async()=>{},onRemoved:new Event()},
    tabs:{update:async()=>{},query:async()=>[{id:10}]},
    webRequest:{onBeforeRequest:new Event(),onHeadersReceived:new Event(),filterResponseData:id=>{const f={write(){},close(){},disconnect(){}};filters.set(id,f);return f;}},
    meetlingHost:{onCommand:new Event(),vaultGet:async base=>vault.get(base)||null,vaultSet:async(base,token,name)=>{vault.set(base,{token,name});},vaultDelete:async base=>vault.delete(base),
      openDraft:async()=>'captured-draft',checkTarget:async id=>{if(id==='closed')throw Error('closed');},insert:async(...args)=>{inserts.push(args);}}
  };
  let transport = async(url,init)=>{
    requests.push({url,init});
    return {ok:true,text:async()=>url.endsWith('/server')?'[{"id":42,"name":"A"}]':url.endsWith('/logout')?'{"token":null}':'[{"id":77,"title":"New room","url":"https://call.example/r"}]'};
  };
  const context = vm.createContext({browser,URL,URLSearchParams,AbortController,TextDecoder,setTimeout,clearTimeout,fetch:(...args)=>transport(...args)});
  for(const file of ['core.js','background.js']) vm.runInContext(fs.readFileSync(new URL('../extension/'+file,'file://'+__filename),'utf8'),context);
  const send=(action,data={})=>[...browser.runtime.onMessage.listeners][0]({action,operationId:'operation-1',...data},{id:'test-addon',url:'moz-extension://test/panel.html'});
  await send('state');
  return {browser,vault,requests,inserts,filters,send,setTransport:fn=>{transport=fn;}};
}
test('login only captures HTTP 200 at exact endpoint in the original main frame',async()=>{
  const f=await setup();await f.send('login');
  const url='https://host.example/meetling/room/outlook/plugin/outlooklogin';
  await f.browser.webRequest.onBeforeRequest.emit({tabId:99,frameId:0,url,requestId:'wrong'});
  assert.equal(f.filters.size,0);
  await f.browser.webRequest.onBeforeRequest.emit({tabId:10,frameId:1,url,requestId:'frame'});assert.equal(f.filters.size,0);
  const details={tabId:10,frameId:0,url,requestId:'right'};
  await f.browser.webRequest.onBeforeRequest.emit(details);
  await f.browser.webRequest.onHeadersReceived.emit({...details,statusCode:200});
  const filter=f.filters.get('right');filter.ondata({data:new TextEncoder().encode('{"token":"private-token","name":"User"}').buffer});await filter.onstop();
  assert.equal(f.vault.get('https://host.example/meetling/').token,'private-token');
  const state=await f.send('state');assert.equal(state.name,'User');assert.ok(!JSON.stringify(state).includes('private-token'));
  assert.equal(f.browser.webRequest.onBeforeRequest.listeners.size,0);
});
test('single HTTPS return upgrade, cancel cleanup and no token accepted from another tab',async()=>{
  const f=await setup();await f.send('login');
  const d={tabId:10,frameId:0,url:'http://host.example/meetling/room/outlook/plugin/outlooklogin',requestId:'redirect'};
  assert.equal((await f.browser.webRequest.onBeforeRequest.emit(d))[0].redirectUrl,d.url.replace('http:','https:'));
  assert.equal((await f.browser.webRequest.onBeforeRequest.emit(d))[0].cancel,true);
  await f.send('cancelLogin');assert.equal(f.browser.webRequest.onBeforeRequest.listeners.size,0);
});
test('create once per operation; target captured first, both links inserted and no token sent to panel',async()=>{
  const f=await setup();f.vault.set('https://host.example/meetling/',{token:'private-token',name:'User'});
  const r=await f.send('create',{serverId:'42'});assert.equal(r.ok,true);assert.equal(r.inserted,true);
  assert.equal(f.inserts[0][0],'captured-draft');assert.match(f.inserts[0][1],/editConference_id=77/);
  assert.equal((await f.send('create',{serverId:'42'})).ok,false);
  assert.equal(f.requests.filter(r=>r.init.method==='POST').length,1);
  assert.ok(!JSON.stringify(await f.send('state')).includes('private-token'));
  assert.equal((await f.send('create',{serverId:'42',operationId:'operation-2'})).ok,true);
});
test('closed draft blocks creation; insertion failure retains room links without retry',async()=>{
  const f=await setup();f.vault.set('https://host.example/meetling/',{token:'private-token',name:'User'});
  assert.equal((await f.send('create',{serverId:'42',target:'closed'})).ok,false);
  assert.equal(f.requests.filter(r=>r.init.method==='POST').length,0);
  f.browser.meetlingHost.insert=async()=>{throw Error('closed during response');};
  const r=await f.send('create',{serverId:'42',target:'draft'});assert.equal(r.ok,true);assert.equal(r.inserted,false);assert.ok(r.conference.editUrl);
  assert.equal((await f.send('create',{serverId:'42',target:'draft'})).ok,false);
});
