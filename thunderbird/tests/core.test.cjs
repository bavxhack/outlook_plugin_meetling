const { test } = require('node:test');
const assert = require('node:assert/strict');
const M = require('../extension/core.js');
const s = M.settings({baseUrl:'https://host.example/meetling/'});
test('installation paths and legacy migration', () => {
  assert.equal(M.endpoint(s,'login'),'https://host.example/meetling/room/outlook/plugin/outlooklogin');
  assert.equal(M.endpoint(s,'servers'),'https://host.example/meetling/api/v1/outlook/plugin/server');
  assert.equal(M.endpoint(M.settings({...s,apiBasePath:'room/outlook/plugin/',endpoints:{login:'login'}}),'login'),M.endpoint(s,'login'));
  for(const baseUrl of ['http://host.example/','https://u:p@host.example/','https://host.example/?token=x','file:///tmp/a']) assert.throws(()=>M.settings({baseUrl}));
  for(const apiBasePath of ['../api/','/api/','https://other.example/','api/?x=1']) assert.throws(()=>M.settings({...s,apiBasePath}));
});
test('exact large numeric IDs, duplicate IDs and errors', () => {
  assert.deepEqual(M.servers(M.parse('[{"id":9007199254740993,"name":"Same"},{"id":9007199254740992,"name":"Same"}]')).map(s=>s.id),['9007199254740993','9007199254740992']);
  assert.throws(()=>M.servers(M.parse('[{"id":1,"name":"A"},{"id":1,"name":"B"}]')));
  for(const body of ['', '<html>Sign in</html>', '{"error":true,"reason":"secret server response"}']) assert.throws(()=>M.parse(body));
  assert.throws(()=>M.parse('{"error":true,"reason":"User not found"}'),e=>e.invalidToken);
});
test('conference links preserve paths, exact ID and escape HTML', () => {
  const c=M.conference(s,M.parse('[{"id":9007199254740993,"url":"https://call.example/room?a=1&b=2","title":"A < B"}]'));
  assert.equal(c.editUrl,'https://host.example/meetling/room/dashboard?editConference_id=9007199254740993');
  assert.match(c.html,/a=1&amp;b=2/);assert.match(c.text,/nur für Moderatoren/);
  assert.ok(c.text.indexOf(c.url)<c.text.indexOf(c.editUrl));
  for(const id of [undefined,0,-1,1.5,'bad&x=1','',true,{}]) {
    const c=M.conference(s,M.parse(JSON.stringify([{id,url:'https://call.example/room',title:'A'}])));
    assert.equal(c.editUrl,null);assert.ok(c.url);
  }
  for(const url of ['http://call.example/','javascript:alert(1)','https://u:p@call.example/']) assert.throws(()=>M.conference(s,[{url,title:'A'}]));
});
test('Bearer, form server_id, no cookies, no redirects for every API action',async()=>{
  const calls=[];
  const transport=async(url,init)=>{calls.push({url,init});return{ok:true,text:async()=>url.endsWith('/server')?'[{"id":42,"name":"Server"}]':url.endsWith('/logout')?'{"token":null}':'[{"id":12,"url":"https://call.example/r","title":"Meetling"}]'};};
  for(const action of ['servers','create','logout']) await M.request(s,action,'synthetic-test-token','42',transport);
  for(const {init} of calls){assert.equal(init.headers.Authorization,'Bearer synthetic-test-token');assert.equal(init.credentials,'omit');assert.equal(init.redirect,'error');assert.equal(init.headers.token,undefined);}
  assert.equal(calls[1].init.body,'server_id=42');assert.equal(calls[1].init.method,'POST');
  assert.equal(calls[2].init.method,'GET');
});
test('failed create never retries and does not expose transport contents',async()=>{
  let count=0;await assert.rejects(M.request(s,'create','test','42',async()=>{count++;throw Error('secret-token');}),e=>!e.message.includes('secret-token')&&e.message.includes('unbekannt'));
  assert.equal(count,1);
  await assert.rejects(M.request(s,'servers','test',null,async()=>({ok:false,status:401})),e=>e.invalidToken);
});
test('token validation, login schema and logout acknowledgement',async()=>{
  assert.equal(M.login('{"token":"safe","name":"User"}').token,'safe');
  for(const token of ['','bad\r\nHeader:x','a b'])assert.throws(()=>M.token(token));
  await assert.rejects(M.request(s,'logout','safe',null,async()=>({ok:true,text:async()=>'{"token":"still-active"}'})));
});
