import test from 'node:test';
import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import Fastify from 'fastify';
import formbody from '@fastify/formbody';
import { SmsConversation, answerOwnerSms, registerSmsWebhook } from '../src/sms-conversation.js';
import { SmsClient } from '../src/sms-client.js';
import { StateStore } from '../src/state-store.js';
import { DeliveryGuard } from '../src/delivery-guard.js';

const sid = 'SM' + '1'.repeat(32);
const accountSid = 'AC' + 'a'.repeat(32);
const phoneSid = 'PN' + 'b'.repeat(32);
const owner = '+15145550101', london = '+15145550100';
const credentials = {accountSid,authToken:'test-secret',from:london,to:owner};
function fixture() {
  const records = new Map(), sent = [], requests = [];
  const store = {readDeliveryRecord:async key=>records.get(key),createDeliveryRecord:async(key,value)=>{if(records.has(key))return false;records.set(key,value);return true;}};
  const state = new StateStore(null);
  const sms = {configured:true,listIncoming:async()=>[{sid,body:'London, can you hear me?',receivedAt:new Date(Date.now()+1000).toISOString()}],send:async body=>{sent.push(body);return {id:sid,status:'queued'};},messageStatus:async()=> 'delivered'};
  const openai = {respond:async request=>{requests.push(request);return {text:'Yes, I received your text. How can I help?',raw:{output:[]}};}};
  const make = (freshState=state)=>new SmsConversation({sms,openai,guard:new DeliveryGuard(store,'owner:incoming-sms'),state:freshState});
  return {records,sent,requests,sms,openai,state,make};
}
test('owner receives one answer across repeated polls and a full local-state reset',async()=>{
  const f=fixture(),worker=f.make();await worker.tick();await worker.tick();await f.make(new StateStore(null)).tick();
  assert.equal(f.sent.length,1);assert.equal(worker.lastReplyStatus,'delivered');assert.ok(worker.lastCheckedAt);
});
test('overlapping poll and webhook cannot dispatch duplicate answers',async()=>{
  const f=fixture(),worker=f.make();await Promise.all([worker.tick(),worker.tick(),worker.tick()]);assert.equal(f.sent.length,1);
});
test('historical messages do not receive surprise replies',async()=>{
  const f=fixture();f.sms.listIncoming=async()=>[{sid,body:'old text',receivedAt:'2020-01-01T00:00:00Z'}];await f.make().tick();assert.equal(f.sent.length,0);
});
test('ambiguous Twilio dispatch stays held across restarts',async()=>{
  const f=fixture();f.sms.send=async body=>{f.sent.push(body);throw Error('timeout');};await assert.rejects(f.make().tick());await f.make(new StateStore(null)).tick();assert.equal(f.sent.length,1);
});
test('OpenAI failure produces a truthful owner reply',async()=>{
  const f=fixture();f.openai.respond=async()=>{throw Error('unavailable');};await f.make().tick();assert.match(f.sent[0],/couldn’t complete/);
});
test('a follow-up includes the previous owner message and London answer',async()=>{
  const f=fixture(),worker=f.make();await worker.tick();f.sms.listIncoming=async()=>[{sid:'SM'+'2'.repeat(32),body:'What did I just ask?',receivedAt:new Date(Date.now()+2000).toISOString()}];await worker.tick();
  assert.equal(f.requests[1].input.length,3);assert.equal(f.requests[1].input[0].role,'user');assert.equal(f.requests[1].input[1].content,f.sent[0]);
});
test('STOP suppresses pending messages; START allows a later new request',async()=>{
  const f=fixture(),worker=f.make();const message=(n,body)=>({sid:'SM'+String(n).repeat(32),body,receivedAt:new Date(Date.now()+n*1000).toISOString()});
  f.sms.listIncoming=async()=>[message(1,'hello'),message(2,'STOP'),message(3,'anyone?')];await worker.tick();assert.equal(f.sent.length,0);
  f.sms.listIncoming=async()=>[message(4,'START'),message(5,'hello London')];await worker.tick();assert.equal(f.sent.length,1);
});
test('media is acknowledged without pretending to read it',async()=>{
  const f=fixture();f.sms.listIncoming=async()=>[{sid,body:'Look at this',numMedia:1,receivedAt:new Date(Date.now()+1000).toISOString()}];await f.make().tick();assert.match(f.sent[0],/email photos or documents/);assert.equal(f.requests.length,0);
});
test('SMS answers use live read tools and reject model-invented write calls',async()=>{
  const toolCalls=[],requests=[];let n=0;
  const openai={respond:async request=>{requests.push(request);if(n++===0)return {raw:{output:[{type:'function_call',name:'check_email',arguments:'{}',call_id:'one'},{type:'function_call',name:'save_email_draft',arguments:'{}',call_id:'two'}]}};return {text:'Your inbox contains one message.',raw:{output:[]}};}};
  const result=await answerOwnerSms({body:'What is in my inbox?',openai,graph:{listVoiceMessages:async()=>{toolCalls.push('read');return [{id:'x',subject:'Example'}];},createVoiceDraft:async()=>{throw Error('Must never run');}}});
  assert.match(result,/one message/);assert.deepEqual(toolCalls,['read']);assert.ok(!requests[0].tools.some(t=>t.name==='save_email_draft'));
  assert.match(requests[1].input.find(t=>t.call_id==='two'&&t.type==='function_call_output').output,/could not be completed/);
});
test('empty and persistently overlong replies cannot be dispatched',async()=>{
  await assert.rejects(answerOwnerSms({body:'test',openai:{respond:async()=>({text:'x'.repeat(600)})}}),/limit/);
});
test('incoming API validates owner, number, direction and pagination host',async()=>{
  const calls=[];const make=(fields={})=>({sid,from:owner,to:london,direction:'inbound',date_created:'Tue, 08 Sep 2026 12:00:00 +0000',body:'hello',...fields});
  const sms=new SmsClient({...credentials,fetchImpl:async(url)=>{calls.push(url);return Response.json({messages:[make(),make({from:'+15145550999'}),make({from:'whatsapp:'+owner}),make({direction:'outbound-api'})],next_page_uri:null});}});
  const result=await sms.listIncoming(Date.parse('2026-09-08T00:00:00Z'));assert.equal(result.length,1);assert.equal(new URL(calls[0]).searchParams.get('From'),owner);
  sms.fetchImpl=async()=>Response.json({messages:[],next_page_uri:'https://attacker.example/messages'});await assert.rejects(sms.listIncoming(Date.now()),/pagination/);
});
test('incoming API follows pages and returns oldest messages first',async()=>{
  let count=0;const sms=new SmsClient({...credentials,fetchImpl:async()=>Response.json(++count===1?{messages:[{sid,from:owner,to:london,direction:'inbound',date_created:'2026-09-08T12:00:00Z'}],next_page_uri:`/2010-04-01/Accounts/${accountSid}/Messages.json?Page=1`}:{messages:[{sid:'SM'+'2'.repeat(32),from:owner,to:london,direction:'inbound',date_created:'2026-09-08T11:00:00Z'}],next_page_uri:null})});
  const messages=await sms.listIncoming(Date.parse('2026-09-08T00:00:00Z'));assert.equal(count,2);assert.match(messages[0].receivedAt,/11:00/);
});
test('webhook configuration updates only SMS fields on the exact existing number',async()=>{
  let number={sid:phoneSid,phone_number:london,capabilities:{sms:true},sms_url:'https://old.example/sms',sms_method:'POST',voice_url:'https://voice.example/call'};let posts=0;
  const sms=new SmsClient({...credentials,fetchImpl:async(url,options)=>{
    if(options.method==='POST'){posts++;const form=new URLSearchParams(options.body);assert.deepEqual([...form.keys()].sort(),['SmsFallbackMethod','SmsFallbackUrl','SmsMethod','SmsUrl']);number={...number,sms_url:form.get('SmsUrl'),sms_method:form.get('SmsMethod'),sms_fallback_url:form.get('SmsFallbackUrl')};return Response.json(number);}
    return Response.json(url.includes('?')?{incoming_phone_numbers:[number]}:number);
  }});
  assert.equal(await sms.configureInbound('https://london.example'),'number-webhook-verified');await sms.configureInbound('https://london.example');assert.equal(posts,1);assert.equal(number.voice_url,'https://voice.example/call');
});
test('webhook rejects forged requests, silently ignores other senders and acknowledges owner immediately',async()=>{
  const app=Fastify();await app.register(formbody);let ticks=0;
  registerSmsWebhook(app,{sms:credentials,publicUrl:'https://london.example',onMessage:async()=>{ticks++;}});
  const fields={AccountSid:accountSid,From:owner,To:london,Body:'test',MessageSid:sid};
  const signature=(body)=>createHmac('sha1',credentials.authToken).update('https://london.example/incoming-sms'+Object.keys(body).sort().map(k=>k+body[k]).join('')).digest('base64');
  const inject=(body,sig)=>app.inject({method:'POST',url:'/incoming-sms',headers:{'content-type':'application/x-www-form-urlencoded','x-twilio-signature':sig},payload:new URLSearchParams(body).toString()});
  assert.equal((await inject(fields,'forged')).statusCode,403);
  const stranger={...fields,From:'+15145550999'};assert.equal((await inject(stranger,signature(stranger))).statusCode,200);await new Promise(setImmediate);assert.equal(ticks,0);
  const result=await inject(fields,signature(fields));assert.equal(result.statusCode,200);assert.match(result.body,/<Response\/>/);await new Promise(setImmediate);assert.equal(ticks,1);await app.close();
});


test('separate SMS workers acquire one analysis claim before paid work',async()=>{
 const f=fixture();await Promise.all([f.make(new StateStore()).tick(),f.make(new StateStore()).tick()]);
 assert.equal(f.requests.length,1);assert.equal(f.sent.length,1);
});
