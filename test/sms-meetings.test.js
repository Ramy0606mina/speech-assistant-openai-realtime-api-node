import test from 'node:test';
import assert from 'node:assert/strict';
import {SmsMeetings} from '../src/sms-meetings.js';
import {TextMeetings} from '../src/text-meetings.js';
import {SmsConversation} from '../src/sms-conversation.js';
import {SmsClient} from '../src/sms-client.js';
import {StateStore} from '../src/state-store.js';

function fixture(){
 const records=new Map(),state=new StateStore(null);let writes=0,outgoing=null;
 const owner='owner@example.com',at='2030-09-10T12:00:00Z';
 const request={text:'Create Teams meeting with admin@example.com tomorrow at 2pm for one hour',owner,requestKey:'proposal',receivedAt:at,maxReplyLength:480};
 const dropbox={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
 const graph={principalMailbox:owner,previewVoiceMeeting:p=>({startIso:p.startIso,endIso:'2030-09-11T15:00:00-04:00',localStart:'2030-09-11T14:00:00',localEnd:'2030-09-11T15:00:00'}),listPrincipalCalendar:async()=>[],createVoiceMeeting:async p=>{writes++;assert.deepEqual(p.attendees,['admin@example.com']);return{id:'created',invitationsSubmitted:true,joinLinkCreated:true};}};
 const openai={respond:async()=>({text:JSON.stringify({title:'Discussion',startIso:'2030-09-11T14:00:00-04:00',durationMinutes:60,contacts:['admin@example.com']})})};
 const meetings=new TextMeetings({graph,dropbox,openai,now:()=>new Date(at)});
 const sms={latestOutgoing:async()=>outgoing};
 const make=(saved=state)=>new SmsMeetings({meetings,state:saved,graph,dropbox,sms});
 const wrapper=make();
 const present=async(key='proposal')=>{
   const reply=await wrapper.handle({...request,requestKey:key});
   outgoing={id:'SM'+key.padEnd(32,'a'),body:reply,status:'delivered',sentAt:'2030-09-10T12:00:05Z'};
   await wrapper.recordReply({requestKey:key,replyId:outgoing.id,text:reply,sentAt:outgoing.sentAt});
   return reply;
 };
 const confirm=(text='confirm',extra={})=>({...request,text,requestKey:'confirmation',receivedAt:'2030-09-10T12:00:10Z',...extra});
 return{records,state,owner,request,graph,openai,meetings,sms,make,wrapper,present,confirm,writes:()=>writes,outgoing:()=>outgoing};
}

for(const spelling of ['confirm','CONFIRM','ConFiRm','  confirm \n',' Confirm! ']){
 test('plain confirmation accepts '+JSON.stringify(spelling),async()=>{
   const f=fixture(),reply=await f.present();
   assert.match(reply,/Discussion/);assert.match(reply,/2030-09-11 14:00:00/);assert.match(reply,/To: admin@example.com/);
   assert.match(reply,/Reply confirm to send the invitation/);assert.doesNotMatch(reply,/CONFIRM MEETING|[a-f0-9]{12}/);
   assert.equal(f.writes(),0);assert.match(await f.wrapper.handle(f.confirm(spelling)),/invitation submitted/);assert.equal(f.writes(),1);
 });
}

test('a second or concurrent confirmation cannot send twice',async()=>{
 const f=fixture();await f.present();
 const replies=await Promise.all([f.wrapper.handle(f.confirm()),f.wrapper.handle(f.confirm('CONFIRM',{requestKey:'other'}))]);
 assert.ok(replies.some(r=>/invitation submitted/.test(r)));assert.equal(f.writes(),1);
 assert.match(await f.make().handle(f.confirm('confirm',{requestKey:'third'})),/invitation submitted/);assert.equal(f.writes(),1);
});

test('no proposal, wrong owner, unsent summary, queued confirmation and same source cannot authorize',async()=>{
 const f=fixture();assert.match(await f.wrapper.handle(f.confirm()),/no current meeting proposal/);
 await f.wrapper.handle(f.request);
 assert.match(await f.wrapper.handle(f.confirm()),/no current meeting proposal/);
 await f.present();
 await assert.rejects(f.wrapper.handle(f.confirm('confirm',{owner:'stranger@example.com'})),/Authenticated owner/);
 assert.match(await f.wrapper.handle(f.confirm('confirm',{receivedAt:'2030-09-10T12:00:04Z'})),/no current meeting proposal/);
 assert.match(await f.wrapper.handle(f.confirm('confirm',{requestKey:'proposal'})),/no current meeting proposal/);
 assert.equal(f.writes(),0);
});

test('restart recovers only the latest actually sent summary from durable reference',async()=>{
 const f=fixture();await f.present();
 assert.match(await f.make(new StateStore(null)).handle(f.confirm()),/invitation submitted/);assert.equal(f.writes(),1);
 assert.match(await f.make(new StateStore(null)).handle(f.confirm('confirm',{requestKey:'again'})),/invitation submitted/);assert.equal(f.writes(),1);
});

test('restart can migrate a verified legacy proposal without typing its code',async()=>{
 const f=fixture();await f.present();const p=[...f.records.entries()].find(([k])=>k.startsWith('text-meeting-proposal-'))[1];
 f.records.forEach((v,k)=>{if(k.startsWith('sms-meeting-reference-'))f.records.delete(k);});
 f.outgoing().body=p.reply;
 assert.match(await f.make(new StateStore(null)).handle(f.confirm()),/invitation submitted/);assert.equal(f.writes(),1);
});

test('modified, failed or newer unrelated outgoing SMS cannot restore an old reference',async()=>{
 for(const fields of [{body:'confirm'},{status:'failed'},{body:'Meeting request cancelled. No invitation was sent.'}]){
   const f=fixture();await f.present();Object.assign(f.outgoing(),fields);
   assert.match(await f.make(new StateStore(null)).handle(f.confirm()),/no current meeting proposal/);assert.equal(f.writes(),0);
 }
});

test('a replacement proposal invalidates the older proposal and binds confirm to the new one',async()=>{
 const f=fixture();await f.present();const old=f.wrapper.pending.code;
 f.openai.respond=async()=>({text:JSON.stringify({title:'Updated discussion',startIso:'2030-09-11T14:00:00-04:00',durationMinutes:60,contacts:['admin@example.com']})});
 await f.present('replacement');assert.notEqual(f.wrapper.pending.code,old);
 assert.match(await f.meetings.confirm(old,f.owner,'old-confirm',480),/cancelled/);
 assert.match(await f.wrapper.handle(f.confirm()),/Updated discussion/);assert.equal(f.writes(),1);
});

test('clarification and cancellation invalidate pending proposals durably',async()=>{
 for(const cancel of [false,true]){
   const f=fixture(),summary=await f.present(),old=f.wrapper.pending.code;
   f.openai.respond=async()=>({text:JSON.stringify({clarification:'Which day?'})});
   const reply=await f.wrapper.handle({...f.request,text:cancel?'never mind':'Create Teams meeting with admin@example.com another day',requestKey:'change',history:[{role:'user',content:f.request.text},{role:'assistant',content:summary}]});
   assert.match(reply,cancel?/cancelled/:/Which day/);
   assert.match(await f.wrapper.handle(f.confirm()),/proposal changed/);
   assert.match(await f.meetings.confirm(old,f.owner,'old',480),/cancelled/);assert.equal(f.writes(),0);
 }
});

test('expired proposal, late conflict and uncertain provider result never retry',async()=>{
 for(const reason of ['expired','conflict','uncertain']){
   const f=fixture();await f.present();
   if(reason==='expired')f.records.get('text-meeting-proposal-'+f.wrapper.pending.code).expiresAt='2020-01-01T00:00:00Z';
   if(reason==='conflict')f.graph.listPrincipalCalendar=async()=>[{showAs:'busy'}];
   let attempts=0;
   if(reason==='uncertain')f.graph.createVoiceMeeting=async()=>{attempts++;throw Error('timeout');};
   assert.match(await f.wrapper.handle(f.confirm()),reason==='expired'?/expired/:reason==='conflict'?/overlaps/:/could not be verified/);
   await f.wrapper.handle(f.confirm('confirm',{requestKey:'retry'}));assert.equal(f.writes(),0);assert.equal(attempts,reason==='uncertain'?1:0);
 }
});

test('SMS conflict corrections prepare a summary and never implicitly send',async()=>{
 const f=fixture();f.graph.listPrincipalCalendar=async()=>[{showAs:'busy'}];
 const conflict=await f.wrapper.handle(f.request);f.graph.listPrincipalCalendar=async()=>[];
 const summary=await f.wrapper.handle({...f.request,text:'Set it for 3pm instead',requestKey:'change',history:[{role:'user',content:f.request.text},{role:'assistant',content:conflict}]});
 assert.match(summary,/Reply confirm/);assert.equal(f.writes(),0);
});

test('SMS conversation retains proposal state after sending and routes the next confirm',async()=>{
 const f=fixture(),sent=[];let body=f.request.text,n=0;
 Object.assign(f.sms,{configured:true,listIncoming:async()=>[{sid:'SM'+String(++n).repeat(32),body,receivedAt:new Date(Date.now()+1000).toISOString()}],send:async text=>{sent.push(text);return{id:'SM'+String(n+5).repeat(32),status:'queued'};},messageStatus:async()=> 'delivered'});
 const guard={notBefore:0,initialize:async()=>{},check:async()=>false,claimAnalysis:async()=>true,claim:async()=>true,complete:async()=>{}};
 const worker=new SmsConversation({sms:f.sms,graph:f.graph,dropbox:f.dropbox,state:f.state,meetings:f.wrapper,guard,openai:{respond:async()=>assert.fail('confirmation must bypass model')}});
 await worker.tick();assert.ok(f.wrapper.pending.presentedAt);assert.match(sent[0],/Reply confirm/);
 body='  CONFIRM  ';await worker.tick();assert.match(sent[1],/invitation submitted/);assert.equal(f.writes(),1);
});

test('control cancellation prevents restoring the former proposal',async()=>{
 const f=fixture();await f.present();await f.wrapper.cancelPending();
 assert.match(await f.make(new StateStore(null)).handle(f.confirm()),/cancelled/);assert.equal(f.writes(),0);
});

test('outgoing lookup is read-only and verifies the configured sender and recipient',async()=>{
 const from='+15145550100',to='+15145550101',sid='SM'+'a'.repeat(32);
 let row={sid,from,to,direction:'outbound-api',body:'summary',status:'delivered',date_sent:'2030-09-10T12:00:05Z'};
 const client=new SmsClient({accountSid:'AC'+'a'.repeat(32),authToken:'test',from,to,fetchImpl:async(url,options)=>{
   assert.equal(options.method,undefined);const query=new URL(url).searchParams;assert.equal(query.get('From'),from);assert.equal(query.get('To'),to);assert.equal(query.get('PageSize'),'1');
   return Response.json({messages:[row]});
 }});
 assert.equal((await client.latestOutgoing()).id,sid);
 row={...row,to:'+15145550999'};assert.equal(await client.latestOutgoing(),null);
});
