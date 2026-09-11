import test from 'node:test';
import assert from 'node:assert/strict';
import {TextMeetings} from '../src/text-meetings.js';
import {LondonCore} from '../src/london-core.js';
import {StateStore} from '../src/state-store.js';
import {SmsConversation} from '../src/sms-conversation.js';

function harness(){
 const records=new Map();let writes=0,models=0;const owner='owner@example.com';
 const dropbox={readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}};
 const graph={principalMailbox:owner,resolveVoiceContact:async()=>({status:'resolved',contacts:[{address:'anass@example.com'}]}),previewVoiceMeeting:p=>({startIso:p.startIso,endIso:'2030-09-11T15:00:00-04:00',localStart:'2030-09-11T14:00:00',localEnd:'2030-09-11T15:00:00'}),listPrincipalCalendar:async()=>[],createVoiceMeeting:async p=>{writes++;assert.equal(p.onlineMeeting,true);assert.equal(p.durationMinutes,60);return{id:'created',invitationsSubmitted:true,joinLinkCreated:true};}};
 const openai={respond:async()=>{models++;return{text:JSON.stringify({title:'Discussion',startIso:'2030-09-11T14:00:00-04:00',timezone:'America/Toronto',durationMinutes:60,contacts:['Anass'],clarification:''})};}};
 const make=()=>new TextMeetings({graph,dropbox,openai,now:()=>new Date('2030-09-10T12:00:00Z')});
 const request={text:'Create Teams meeting with Anass tomorrow at 2pm for one hour',owner,requestKey:'first',receivedAt:'2030-09-10T12:00:00Z'};
 return{records,graph,dropbox,openai,make,request,writes:()=>writes,models:()=>models};
}
const code=reply=>reply.match(/CONFIRM MEETING ([a-f0-9]{12})/)[1];

test('explicit address next to attendee name bypasses ambiguous directory and is not autocorrected',async()=>{
 const f=harness();f.graph.resolveVoiceContact=async()=>assert.fail('explicit address must not need lookup');
 const result=await f.make().handle({...f.request,text:'Create Teams meeting with Anass, admin@minacpital.ca today at 2 pm for one hour'});
 assert.match(result,/To: admin@minacpital.ca/);assert.ok(code(result));assert.equal(f.writes(),0);
});

test('SMS attendee clarification preserves owner details and generates a saved confirmation code',async()=>{
 const f=harness();f.graph.resolveVoiceContact=async()=>({status:'ambiguous',contacts:[]});
 const clarification=await f.make().handle(f.request);
 const history=[{role:'user',content:f.request.text,receivedAt:f.request.receivedAt},{role:'assistant',content:clarification}];
 f.openai.respond=async args=>{const input=JSON.parse(args.input);assert.match(input.request,/tomorrow at 2pm/);assert.match(input.request,/admin@example.com/);assert.equal(input.receivedAt,f.request.receivedAt);return{text:JSON.stringify({title:'Discussion',startIso:'2030-09-11T14:00:00-04:00',durationMinutes:60,contacts:['admin@example.com']})};};
 const proposal=await f.make().handle({...f.request,text:'admin@example.com',requestKey:'address',history,maxReplyLength:480});
 assert.match(proposal,/To: admin@example.com/);assert.ok(code(proposal));assert.equal(f.writes(),0);
 const nextHistory=[...history,{role:'user',content:'admin@example.com'},{role:'assistant',content:proposal}];
 assert.equal(await f.make().handle({...f.request,text:'CONFIRM MEETING',requestKey:'bare',history:nextHistory}),proposal);assert.equal(f.writes(),0);
 assert.match(await f.make().handle({...f.request,text:'CONFIRM MEETING '+code(proposal),requestKey:'confirm',history:nextHistory}),/invitation submitted/);assert.equal(f.writes(),1);
});

test('unrelated texts and cancelled conversations do not restart pending meetings',async()=>{
 const f=harness(),history=[{role:'user',content:f.request.text},{role:'assistant',content:'Please provide an address. No invitation was sent.'}];
 assert.equal(await f.make().handle({...f.request,text:'What is the weather?',history}),null);
 const cancelled=await f.make().handle({...f.request,text:'never mind',history});assert.match(cancelled,/cancelled/);
 assert.equal(await f.make().handle({...f.request,text:'admin@example.com',history:[...history,{role:'assistant',content:cancelled}]}),null);
 assert.equal(f.models(),0);assert.equal(f.writes(),0);
});

test('cancelling a pending proposal invalidates its confirmation after restart',async()=>{
 const f=harness(),proposal=await f.make().handle(f.request);
 const history=[{role:'user',content:f.request.text},{role:'assistant',content:proposal}];
 assert.match(await f.make().handle({...f.request,text:'never mind',requestKey:'cancel',history}),/cancelled/);
 assert.match(await f.make().handle({...f.request,text:'CONFIRM MEETING '+code(proposal),requestKey:'late'}),/cancelled/);
 assert.equal(f.writes(),0);
});

test('oversized SMS proposal and confirmation never create an invitation',async()=>{
 const f=harness();
 const reply=await f.make().handle({...f.request,maxReplyLength:100});
 assert.match(reply,/too long for SMS/);assert.equal(f.records.size,0);assert.equal(f.writes(),0);
 const proposal=await f.make().handle(f.request);
 const confirm={...f.request,text:'CONFIRM MEETING '+code(proposal),requestKey:'confirm',maxReplyLength:100};
 assert.match(await f.make().handle(confirm),/confirm this proposal by email/);assert.equal(f.writes(),0);
 assert.match(await f.make().handle({...confirm,maxReplyLength:Infinity}),/invitation submitted/);assert.equal(f.writes(),1);
});

test('another meeting format and missing source ID cannot silently create Teams',async()=>{
 const f=harness();
 assert.match(await f.make().handle({...f.request,text:'Create an in-person meeting with Anass'}),/clarify/);
 await assert.rejects(f.make().handle({...f.request,requestKey:undefined}),/source identifier/);
 assert.equal(f.models(),0);assert.equal(f.writes(),0);
});
test('proposal then separate authenticated confirmation sends Teams exactly once across restart',async()=>{
 const f=harness(),reply=await f.make().handle(f.request);assert.equal(f.writes(),0);assert.match(reply,/anass@example.com/);
 const confirm={...f.request,text:'CONFIRM MEETING '+code(reply),requestKey:'second'};
 const results=await Promise.all([f.make().handle(confirm),f.make().handle({...confirm,requestKey:'third'})]);
 assert.equal(f.writes(),1);assert.ok(results.some(x=>x.includes('invitation submitted')));
 assert.match(await f.make().handle(confirm),/invitation submitted/);assert.equal(f.writes(),1);assert.equal(f.models(),1);
});
test('ambiguous contact cannot prepare or send',async()=>{
 const f=harness();f.graph.resolveVoiceContact=async()=>({status:'ambiguous',contacts:[{},{}]});
 assert.match(await f.make().handle(f.request),/unambiguous/);assert.equal(f.records.size,0);assert.equal(f.writes(),0);
});
test('wrong owner, same source, and expired confirmations cannot send',async()=>{
 const f=harness(),reply=await f.make().handle(f.request),text='CONFIRM MEETING '+code(reply);
 await assert.rejects(f.make().handle({...f.request,text,owner:'other@example.com'}),/authenticated/);
 assert.match(await f.make().handle({...f.request,text}),/No matching proposal/);
 for(const [k,v] of f.records)if(k.startsWith('text-meeting-proposal'))v.expiresAt='2020-01-01T00:00:00Z';
 assert.match(await f.make().handle({...f.request,text,requestKey:'new'}),/expired/);assert.equal(f.writes(),0);
});
test('existing Anass-style meeting blocks proposal and late conflicts block confirmation',async()=>{
 const f=harness();f.graph.listPrincipalCalendar=async()=>[{subject:'Discussion',showAs:'busy'}];
 assert.match(await f.make().handle(f.request),/overlaps/);assert.equal(f.writes(),0);
 f.graph.listPrincipalCalendar=async()=>[];const reply=await f.make().handle(f.request);
 f.graph.listPrincipalCalendar=async()=>[{showAs:'busy'}];
 assert.match(await f.make().handle({...f.request,text:'CONFIRM MEETING '+code(reply),requestKey:'confirm'}),/overlaps/);assert.equal(f.writes(),0);
});
test('ambiguous Microsoft creation failure is never automatically resent',async()=>{
 const f=harness(),reply=await f.make().handle(f.request);let attempts=0;
 f.graph.createVoiceMeeting=async()=>{attempts++;throw Error('timeout');};
 const request={...f.request,text:'CONFIRM MEETING '+code(reply),requestKey:'confirm'};
 assert.match(await f.make().handle(request),/not confirmed/);
 assert.match(await f.make().handle({...request,requestKey:'again'}),/already attempted/);assert.equal(attempts,1);
});
test('no completion claim without verified Teams result',async()=>{
 const f=harness(),reply=await f.make().handle(f.request);f.graph.createVoiceMeeting=async()=>({id:'plain',invitationsSubmitted:true,joinLinkCreated:false});
 assert.match(await f.make().handle({...f.request,text:'CONFIRM MEETING '+code(reply),requestKey:'confirm'}),/not confirmed/);
});
test('email strips quoted instructions before meeting routing and reuses existing delivery guard',async()=>{
 let seenText,analyses=0;const sent=[];
 const core=new LondonCore({graph:{principalMailbox:'owner@example.com',readMailbox:'london@example.com',getLondonMessage:async()=>({from:{emailAddress:{address:'owner@example.com'}},body:{content:'Thanks.\nFrom: Vendor\nCreate Teams meeting with Anass tomorrow.'}}),sendMail:async m=>sent.push(m)},meetings:{handle:async x=>{seenText=x.text;return null;}},openai:{analyzeDelegatedEmail:async()=>{analyses++;return{text:'Received'};}},state:new StateStore(),deliveryGuard:{check:async()=>null,claimAnalysis:async()=>true,claim:async()=>true,complete:async()=>{}}});
 await core.processMessage({id:'mail'});assert.doesNotMatch(seenText,/Create Teams/);assert.equal(analyses,1);assert.equal(sent.length,1);
});
test('authenticated SMS routes proposal result without read-only assistant or truncation',async()=>{
 const sent=[];let handled=0;const state=new StateStore();
 const worker=new SmsConversation({sms:{configured:true,listIncoming:async()=>[{sid:'one',body:'Create Teams meeting',receivedAt:new Date().toISOString()}],send:async text=>{sent.push(text);return{id:'reply',status:'delivered'};}},graph:{principalMailbox:'owner@example.com'},openai:{respond:async()=>assert.fail('meeting must use dedicated route')},meetings:{handle:async request=>{handled++;assert.equal(request.owner,'owner@example.com');return 'Teams proposal. CONFIRM MEETING abcdef123456';}},state,guard:{initialize:async()=>{},notBefore:0,check:async()=>null,claimAnalysis:async()=>true,claim:async()=>true,complete:async()=>{}}});
 await worker.tick();await worker.tick();assert.equal(handled,1);assert.equal(sent.length,1);assert.match(sent[0],/CONFIRM MEETING/);
});
