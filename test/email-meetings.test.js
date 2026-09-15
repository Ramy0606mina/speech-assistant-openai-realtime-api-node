import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtempSync, rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {EmailMeetings} from '../src/email-meetings.js';
import {TextMeetings} from '../src/text-meetings.js';
import {StateStore} from '../src/state-store.js';
import {LondonCore} from '../src/london-core.js';
import {EmailDraftCreation} from '../src/email-draft-creation.js';
import {directOwnerRequestText} from '../src/email-reminder.js';

function fixture(file=null) {
  const records=new Map(), owner='owner@example.com'; let writes=0;
  const dropbox={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
  const graph={principalMailbox:owner,previewVoiceMeeting:p=>({startIso:p.startIso,endIso:'2030-09-17T17:30:00-04:00',localStart:'2030-09-17T16:30:00',localEnd:'2030-09-17T17:30:00'}),listPrincipalCalendar:async()=>[],createVoiceMeeting:async()=>{writes++;return {id:'event',invitationsSubmitted:true,joinLinkCreated:true};}};
  const openai={respond:async r=>{assert.equal(JSON.parse(r.input).receivedAt,'2030-09-15T16:21:20Z');return {text:JSON.stringify({title:'Mina Haus',startIso:'2030-09-17T16:30:00-04:00',durationMinutes:60,contacts:['a@example.com'],clarification:''})};}};
  const state=new StateStore(file), now=()=>new Date('2030-09-15T16:22:00Z');
  const meetings=new TextMeetings({graph,dropbox,openai,now});
  const handler=new EmailMeetings({graph,dropbox,state,meetings,now});
  const request={owner,conversationId:'thread-a',requestKey:'request',receivedAt:'2030-09-15T16:21:20Z',text:'Prepare and send a Teams meeting on Thursday 17 at 4:30 pm for one hour with a@example.com'};
  const confirm={...request,requestKey:'confirm',receivedAt:'2030-09-15T16:23:00Z',text:'confirm'};
  return {records,dropbox,graph,openai,state,now,meetings,handler,request,confirm,writes:()=>writes};
}
async function present(f) {const reply=await f.handler.handle(f.request);f.handler.recordReply(f.request);return reply;}

for(const text of ['confirm','CONFIRM',' Confirm! ','confirm.']) test('email accepts '+JSON.stringify(text)+' without visible code',async()=>{
  const f=fixture();const reply=await present(f);assert.match(reply,/Reply confirm/);assert.doesNotMatch(reply,/CONFIRM MEETING|[a-f0-9]{12}/);
  assert.match(await f.handler.handle({...f.confirm,text}),/invitation submitted/);assert.equal(f.writes(),1);
});
test('email duplicate and concurrent confirmations create only one invitation',async()=>{
 const f=fixture();await present(f);await Promise.all([f.handler.handle(f.confirm),f.handler.handle({...f.confirm,requestKey:'second'})]);await f.handler.handle({...f.confirm,requestKey:'third'});assert.equal(f.writes(),1);
});
test('email confirmation requires the same owner and conversation, a sent proposal and a later source',async()=>{
 const f=fixture();await f.handler.handle(f.request);
 assert.doesNotMatch(await f.handler.handle(f.confirm),/invitation submitted/);
 f.handler.recordReply(f.request);
 for(const patch of [{conversationId:'other'},{conversationId:null},{requestKey:'request'},{receivedAt:'2030-09-15T16:21:00Z'},{receivedAt:'invalid'}])assert.doesNotMatch(await f.handler.handle({...f.confirm,...patch}),/invitation submitted/);
 await assert.rejects(f.handler.handle({...f.confirm,owner:'stranger@example.com'}),/Authenticated/);assert.equal(f.writes(),0);
});
test('pending email survives state reload and completed confirmation remains idempotent',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'london-email-'));try {
  const f=fixture(join(dir,'state.json'));await present(f);
  const restored=()=>new EmailMeetings({...f,state:new StateStore(join(dir,'state.json'))});
  assert.match(await restored().handle(f.confirm),/invitation submitted/);
  assert.match(await restored().handle({...f.confirm,requestKey:'again'}),/invitation submitted/);assert.equal(f.writes(),1);
 }finally{rmSync(dir,{recursive:true,force:true});}
});
test('a missing persisted reference fails closed',async()=>{
 const f=fixture();await present(f);f.state.state.emailMeetings={};assert.match(await f.handler.handle(f.confirm),/no current/);assert.equal(f.writes(),0);
});
test('replacement proposal supersedes the old one and needs successful reply delivery',async()=>{
 const f=fixture();await present(f);const old=f.handler.pending(f.handler.key(f.request.owner,f.request.conversationId)).code;
 await f.handler.handle({...f.request,requestKey:'new',text:f.request.text+' instead'});
 assert.ok(f.records.has('text-meeting-cancelled-'+old));assert.doesNotMatch(await f.handler.handle(f.confirm),/invitation submitted/);assert.equal(f.writes(),0);
});
test('cancelled, expired and conflicting proposals cannot be confirmed',async()=>{
 for(const mode of ['cancel','expire','conflict']){
  const f=fixture();await present(f);
  if(mode==='cancel')await f.handler.handle({...f.confirm,text:'cancel',requestKey:'cancel'});
  if(mode==='expire')f.meetings.now=()=>new Date('2030-09-18T00:00:00Z');
  if(mode==='conflict')f.graph.listPrincipalCalendar=async()=>[{showAs:'busy'}];
  assert.doesNotMatch(await f.handler.handle(f.confirm),/invitation submitted/);assert.equal(f.writes(),0);
 }
});
test('uncertain invitation submission cannot be retried by another confirm',async()=>{
 const f=fixture();let attempts=0;f.graph.createVoiceMeeting=async()=>{attempts++;throw Error('timeout');};await present(f);
 assert.match(await f.handler.handle(f.confirm),/could not be verified/);await f.handler.handle({...f.confirm,requestKey:'again'});assert.equal(attempts,1);
});
test('sender signature is removed without accepting corrections as a bare confirm',()=>{
 const base={from:{emailAddress:{name:'RAMY MINA'}},body:{content:'confirm\r\n\r\nRAMY MINA\r\nFounder\r\nEmail: owner@example.com'}};
 assert.equal(directOwnerRequestText(base),'confirm');
 assert.equal(directOwnerRequestText({...base,body:{content:'confirm\nBut change it to Friday\n\nRAMY MINA\nEmail: owner@example.com'}}),'confirm\nBut change it to Friday');
});
test('original signed meeting email routes to meetings and sends an in-thread plain-confirm proposal',async()=>{
 const f=fixture(),sent=[];
 const full={id:'original',internetMessageId:'request',conversationId:'thread-a',receivedDateTime:f.request.receivedAt,subject:'Mina Haus meeting',from:{emailAddress:{address:f.request.owner,name:'RAMY MINA'}},body:{content:f.request.text+'\n\nRAMY MINA\nEmail: owner@example.com'}};
 f.graph.getLondonMessage=async()=>full;f.graph.readMailbox='london@example.com';
 f.graph.replyToLondonMessage=async q=>{sent.push(q);return {sent:true};};f.graph.sendMail=async()=>assert.fail('Meeting proposal must stay in-thread');
 const draft=new EmailDraftCreation({...f,openai:{respond:async()=>assert.fail('Meeting was misrouted to draft')}});
 const core=new LondonCore({...f,emailDraftCreation:draft,emailMeetings:f.handler,deliveryGuard:{check:async()=>null,claimAnalysis:async()=>true,claim:async()=>true,complete:async()=>{}}});
 const result=await core.processMessage(full);assert.equal(result.completionSent,true);assert.equal(sent[0].conversationId,'thread-a');assert.match(sent[0].body,/Reply confirm/);assert.equal(f.handler.pending(f.handler.key(f.request.owner,'thread-a')).status,'pending');assert.equal(f.writes(),0);
});
