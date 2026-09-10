import test from 'node:test';
import assert from 'node:assert/strict';
import {UrgentAlerts,buildUrgentSms,urgencyInstructions} from '../src/urgent-alerts.js';
import {DeliveryGuard} from '../src/delivery-guard.js';
import {SmsClient} from '../src/sms-client.js';

function fixture(){
 const records=new Map(),sent=[];
 const store={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
 const graph={listPrincipalInbox:async()=>[{id:'new',receivedDateTime:new Date(Date.now()+1000).toISOString(),bodyPreview:'test'}]};
 const openai={respond:async({instructions})=>({text:instructions===urgencyInstructions?'URGENT':'Please call back now.'})};
 const sms={configured:true,send:async text=>sent.push(text)};
 const make=()=>new UrgentAlerts({graph,openai,sms,guard:new DeliveryGuard(store,'owner:urgent')});
 return {graph,openai,sms,sent,make};
}
test('urgent alert sends once across polling and restart',async()=>{
 const f=fixture(),a=f.make();await a.tick();await a.tick();await f.make().tick();assert.equal(f.sent.length,1);
});
test('historical mail and routine messages do not send alerts',async()=>{
 const f=fixture();f.graph.listPrincipalInbox=async()=>[{id:'old',receivedDateTime:'2020-01-01T00:00:00Z'}];await f.make().tick();
 f.graph.listPrincipalInbox=async()=>[{id:'normal',receivedDateTime:new Date(Date.now()+1000).toISOString()}];f.openai.respond=async()=>({text:'NORMAL'});await f.make().tick();assert.equal(f.sent.length,0);
});
test('classification failure holds paid analysis across restart',async()=>{
 const f=fixture(),a=f.make();f.openai.respond=async()=>{throw Error('unavailable');};await assert.rejects(a.tick());assert.equal(f.sent.length,0);
 f.openai.respond=async()=>assert.fail('must not repeat paid analysis');await a.tick();await f.make().tick();assert.equal(a.heldForReview,1);assert.equal(f.sent.length,0);
});
test('ambiguous SMS failure is held after restart',async()=>{
 const f=fixture();f.sms.send=async text=>{f.sent.push(text);throw Error('timeout');};await assert.rejects(f.make().tick());await f.make().tick();assert.equal(f.sent.length,1);
});
test('SMS recipient is fixed and queue acceptance is not delivery confirmation',async()=>{
 let form;const sms=new SmsClient({accountSid:'AC'+'a'.repeat(32),authToken:'test',from:'+15145550100',to:'+15145550101',fetchImpl:async(u,o)=>{form=new URLSearchParams(o.body);return Response.json({sid:'SMtest',status:'queued'});}});
 const result=await sms.send('Test');assert.equal(form.get('To'),'+15145550101');assert.equal(result.status,'queued');assert.equal(result.delivered,undefined);
});
test('invalid classification is held for review',async()=>{
 const f=fixture(),a=f.make();f.openai.respond=async()=>({text:'Maybe'});await assert.rejects(a.tick(),/invalid decision/);
 f.openai.respond=async()=>assert.fail('must not repeat paid analysis');await a.tick();await f.make().tick();assert.equal(a.heldForReview,1);assert.equal(f.sent.length,0);
});

test('alert includes actual sender and a summary grounded in the email preview',async()=>{
 const message={from:{emailAddress:{name:'Mina Capital',address:'sender@example.com'}},subject:'Callback',bodyPreview:'Ramy, please call me now.'};
 const text=await buildUrgentSms(message,{respond:async request=>{
  assert.deepEqual(JSON.parse(request.input),{subject:'Callback',preview:'Ramy, please call me now.'});
  return {text:'Please call back now.'};
 }});
 assert.equal(text,'London | From: Mina Capital\nPlease call back now.');
});

test('summary failure holds the entire analysis without dispatch',async()=>{
 const f=fixture(),a=f.make();
 f.openai.respond=async({instructions})=>{if(instructions!==urgencyInstructions)throw Error('summary unavailable');return {text:'URGENT'};};
 await assert.rejects(a.tick(),/summary unavailable/);assert.equal(f.sent.length,0);
 f.openai.respond=async()=>assert.fail('must not repeat paid analysis');
 await a.tick();await f.make().tick();assert.equal(f.sent.length,0);
});

test('sender fallback and message size remain bounded',async()=>{
 const text=await buildUrgentSms({from:{emailAddress:{address:'owner@example.com'}}},{respond:async()=>({text:'x'.repeat(500)})});
 assert.match(text,/From: owner@example.com/);assert.ok(text.length<250);
});


test('routine classification survives restart without another paid request',async()=>{
 const f=fixture();let calls=0;f.openai.respond=async()=>{calls++;return {text:'NORMAL'};};
 await Promise.all([f.make().tick(),f.make().tick()]);await f.make().tick();assert.equal(calls,1);assert.equal(f.sent.length,0);
});
