import test from 'node:test';
import assert from 'node:assert/strict';
import {UrgentAlerts} from '../src/urgent-alerts.js';
import {DeliveryGuard} from '../src/delivery-guard.js';
import {SmsClient} from '../src/sms-client.js';

function fixture(){
 const records=new Map(),sent=[];
 const store={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
 const graph={listPrincipalInbox:async()=>[{id:'new',receivedDateTime:new Date(Date.now()+1000).toISOString(),bodyPreview:'test'}]};
 const openai={respond:async()=>({text:'URGENT'})};
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
test('classification failure remains retryable without premature dispatch',async()=>{
 const f=fixture(),a=f.make();f.openai.respond=async()=>{throw Error('unavailable');};await assert.rejects(a.tick());assert.equal(f.sent.length,0);
 f.openai.respond=async()=>({text:'URGENT'});await a.tick();assert.equal(f.sent.length,1);
});
test('ambiguous SMS failure is held after restart',async()=>{
 const f=fixture();f.sms.send=async text=>{f.sent.push(text);throw Error('timeout');};await assert.rejects(f.make().tick());await f.make().tick();assert.equal(f.sent.length,1);
});
test('SMS recipient is fixed and queue acceptance is not delivery confirmation',async()=>{
 let form;const sms=new SmsClient({accountSid:'AC'+'a'.repeat(32),authToken:'test',from:'+15145550100',to:'+15145550101',fetchImpl:async(u,o)=>{form=new URLSearchParams(o.body);return Response.json({sid:'SMtest',status:'queued'});}});
 const result=await sms.send('Test');assert.equal(form.get('To'),'+15145550101');assert.equal(result.status,'queued');assert.equal(result.delivered,undefined);
});
test('invalid classification remains retryable rather than silently skipped',async()=>{
 const f=fixture(),a=f.make();f.openai.respond=async()=>({text:'Maybe'});await assert.rejects(a.tick(),/invalid decision/);
 f.openai.respond=async()=>({text:'URGENT'});await a.tick();assert.equal(f.sent.length,1);
});
