import test from 'node:test';
import assert from 'node:assert/strict';
import { TextReminders } from '../src/text-reminders.js';
import { SmsConversation } from '../src/sms-conversation.js';
import { StateStore } from '../src/state-store.js';
import { DeliveryGuard } from '../src/delivery-guard.js';

const owner='owner@example.com';
const now=()=>new Date('2030-07-10T12:00:00Z');
const text='Put a reminder for me tomorrow at 9 am to mention to Remi that I want to pass the procedure on the RMQ';
const args={title:'Mention the RMQ procedure to Remi',startIso:'2030-07-11T09:00:00-04:00',timeText:'9 am',clarification:''};
function fixture(){
  const records=new Map(),writes=[],inputs=[];
  const dropbox={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
  const graph={principalMailbox:owner,createPersonalReminder:async p=>{writes.push(p);return {id:'event',created:true,reminderOn:true,title:p.title,startLocal:'2030-07-11T09:00:00',timezone:'Eastern time'};}};
  const openai={respond:async r=>{inputs.push(r);return {text:JSON.stringify(args)};}};
  const make=()=>new TextReminders({graph,dropbox,openai,now});
  const request={text,owner,requestKey:'SM-owner-1',receivedAt:now().toISOString()};
  return {records,writes,inputs,dropbox,graph,openai,make,request};
}
test('explicit SMS reminder routes to existing Graph action and confirms only its receipt',async()=>{
  const f=fixture();const reply=await f.make().handle(f.request);assert.match(reply,/Reminder saved/);assert.equal(f.writes.length,1);assert.equal(f.writes[0].startIso,args.startIso);assert.match(f.writes[0].notes,/Remi/);assert.ok(f.writes[0].taskKey.startsWith('sms-reminder-'));
  await f.make().handle(f.request);assert.equal(f.writes.length,1);
});
test('screenshot request without a time asks for time instead of refusing SMS',async()=>{
  const f=fixture();f.openai.respond=async()=>({text:JSON.stringify({...args,timeText:'tomorrow',clarification:''})});
  const reply=await f.make().handle({...f.request,text:text.replace(' at 9 am','')});assert.match(reply,/What time/);assert.equal(f.writes.length,0);
});
test('short time clarification carries original owner purpose, date and durable request key',async()=>{
  const f=fixture();const history=[{role:'user',content:text.replace(' at 9 am',''),receivedAt:'2030-07-10T12:00:00Z',requestKey:'original'},{role:'assistant',content:'Reminder not saved yet: What time tomorrow?'}];
  await f.make().handle({...f.request,text:'9 am',requestKey:'followup',receivedAt:'2030-07-11T01:00:00Z',history});
  const input=JSON.parse(f.inputs[0].input);assert.match(input.request,/Remi/);assert.equal(input.receivedAt,'2030-07-10T12:00:00Z');assert.equal(f.writes.length,1);
  await f.make().handle({...f.request,text:'9 am',requestKey:'another-followup',history});assert.equal(f.writes.length,1);
});
test('unrelated and quoted reminder text cannot authorize creation',async()=>{
  const f=fixture();for(const value of ['Read my inbox','Explain this: put a reminder tomorrow at 9 am','Do not remind me tomorrow at 9 am'])assert.equal(await f.make().handle({...f.request,text:value}),null);
  assert.equal(f.inputs.length,0);assert.equal(f.writes.length,0);
});
test('another sender is rejected before AI or Graph access',async()=>{
  const f=fixture();await assert.rejects(f.make().handle({...f.request,owner:'stranger@example.com'}),/Authenticated owner/);assert.equal(f.inputs.length,0);assert.equal(f.writes.length,0);
});
test('failed or ambiguous write never reports success or retries after restart',async()=>{
  const f=fixture();f.graph.createPersonalReminder=async p=>{f.writes.push(p);throw Error('timeout');};const reply=await f.make().handle(f.request);assert.match(reply,/did not confirm/);await f.make().handle(f.request);assert.equal(f.writes.length,1);
});
test('past dates and unverified Graph responses cannot report success',async()=>{
  const f=fixture();f.openai.respond=async()=>({text:JSON.stringify({...args,startIso:'2020-07-11T09:00:00-04:00'})});assert.match(await f.make().handle(f.request),/future date/);assert.equal(f.writes.length,0);
  f.openai.respond=async()=>({text:JSON.stringify(args)});f.graph.createPersonalReminder=async()=>({id:'event',created:true,reminderOn:false});assert.doesNotMatch(await f.make().handle(f.request),/Reminder saved/);
});
test('malformed model output and invented time ask for clarification without writing',async()=>{
  const f=fixture();f.openai.respond=async()=>({text:'not-json'});assert.match(await f.make().handle(f.request),/not saved yet/);
  f.openai.respond=async()=>({text:JSON.stringify({...args,timeText:'10 am'})});assert.match(await f.make().handle(f.request),/What time/);assert.equal(f.writes.length,0);
});
test('SMS worker invokes reminders before meeting or read-only handlers and sends verified result',async()=>{
  const f=fixture(),sent=[];const state=new StateStore(null),guard=new DeliveryGuard(f.dropbox,owner);
  const sms={configured:true,listIncoming:async()=>[{sid:'SM'+'1'.repeat(32),body:text,receivedAt:'2030-07-10T12:00:00Z'}],send:async body=>{sent.push(body);return {id:'reply',status:'delivered'};}};
  const worker=new SmsConversation({sms,openai:f.openai,graph:f.graph,dropbox:f.dropbox,state,guard,reminders:f.make(),meetings:{handle:async()=>{throw Error('must not route reminder to meetings');}}});
  await worker.tick();await worker.tick();assert.equal(f.writes.length,1);assert.equal(sent.length,1);assert.match(sent[0],/Reminder saved/);assert.ok(state.state.smsConversation.history[0].requestKey);
});
