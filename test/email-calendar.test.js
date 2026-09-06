import test from 'node:test';
import assert from 'node:assert/strict';
import { OpenAIClient } from '../src/openai-client.js';

test('owner calendar request reads actual events and supplies them to the reply',async()=>{
 let step=0;
 const client=new OpenAIClient({apiKey:'test',fetchImpl:async(url,o)=>{
  const body=JSON.parse(o.body);
  if(++step===1){assert.ok(body.tools.some(t=>t.name==='read_principal_calendar'));return Response.json({output:[{type:'function_call',call_id:'calendar',name:'read_principal_calendar',arguments:JSON.stringify({startIso:'2026-09-07T00:00:00-04:00',endIso:'2026-09-08T00:00:00-04:00'})}]});}
  assert.match(JSON.stringify(body.input),/Verified meeting/);return Response.json({output_text:'Verified meeting is on your calendar.'});
 }});
 const reply=await client.analyzeDelegatedEmail({subject:'Calendar'},[],{graph:{listPrincipalCalendar:async q=>{assert.equal(q.limit,50);return [{subject:'Verified meeting'}];}}});
 assert.match(reply.text,/Verified meeting/);
});

import { MicrosoftGraphClient } from '../src/microsoft-graph.js';
function pagedGraph(page) {
 return new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'s',ramyMailbox:'owner@example.com',fetchImpl:async(url,o)=>String(url).includes('oauth2')?Response.json({access_token:'test',expires_in:3600}):page(new URL(url),o)});
}
const range={startIso:'2026-09-01T00:00:00Z',endIso:'2026-10-01T00:00:00Z'};
test('calendar follows nextLink and includes conflicts beyond the first 50 events',async()=>{
 let pages=0;
 const graph=pagedGraph(async(url,o)=>{
  pages++;assert.equal(o.headers.Prefer,'outlook.timezone="Eastern Standard Time"');
  if(pages===1){url.searchParams.set('$skiptoken','opaque');return Response.json({value:Array.from({length:50},(_,i)=>({id:String(i)})),'@odata.nextLink':url.href});}
  assert.equal(url.searchParams.get('$skiptoken'),'opaque');return Response.json({value:[{id:'late-conflict',showAs:'busy'}]});
 });
 const events=await graph.listPrincipalCalendar(range);assert.equal(events.length,51);assert.equal(events[50].id,'late-conflict');assert.equal(pages,2);
});
test('calendar never returns partial availability after a later-page failure',async()=>{
 let pages=0;const graph=pagedGraph(async(url)=>{if(++pages===2)throw new Error('second page failed');url.searchParams.set('$skiptoken','next');return Response.json({value:[{id:'first'}],'@odata.nextLink':url.href});});
 await assert.rejects(graph.listPrincipalCalendar(range),/second page failed/);
});
test('calendar rejects foreign pagination and repeated links',async()=>{
 for(const foreign of [true,false]){
  let calls=0;const graph=pagedGraph(async(url)=>{calls++;return Response.json({value:[],'@odata.nextLink':foreign?'https://example.com/collect':url.href});});
  await assert.rejects(graph.listPrincipalCalendar(range),/invalid|did not complete/);assert.equal(calls,1);
 }
});
