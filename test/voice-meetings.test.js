import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {runVoiceTool,voiceTools} from '../src/voice-gateway.js';

function client(handler){return new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'s',actionTenantId:'t',actionClientId:'c',actionClientSecret:'s',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async(url,options)=>String(url).includes('oauth2')?Response.json({access_token:'token',expires_in:3600}):handler(String(url),options)});}

test('Microsoft meeting includes attendee invitations and a stable transaction',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});return Response.json({id:'event-1'});});
  const result=await graph.createVoiceMeeting({title:'Project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['JACK@example.com'],body:'Discuss next steps.',location:'Microsoft Teams',transactionId:'12345678-1234-4123-8123-123456789abc'});
  assert.equal(result.invitationsSubmitted,true);assert.equal(calls.length,1);assert.match(calls[0].url,/owner%40example.com\/events$/);
  const body=JSON.parse(calls[0].options.body);assert.equal(body.attendees[0].emailAddress.address,'jack@example.com');assert.deepEqual(body.start,{dateTime:'2030-07-11T10:00:00',timeZone:'Eastern Standard Time'});assert.deepEqual(body.end,{dateTime:'2030-07-11T10:30:00',timeZone:'Eastern Standard Time'});assert.equal(body.transactionId,'12345678-1234-4123-8123-123456789abc');
  assert.equal(result.timezone,'America/Toronto');assert.equal(result.microsoftTimeZone,'Eastern Standard Time');
});

test('Toronto meetings keep ten a.m. through winter and summer daylight-saving offsets',async()=>{
  const payloads=[];const graph=client((url,options)=>{payloads.push(JSON.parse(options.body));return Response.json({id:`event-${payloads.length}`});});
  const base={title:'Toronto call',durationMinutes:60,timezone:'America/Toronto',attendees:['person@example.com'],transactionId:'12345678-1234-4123-8123-123456789abc'};
  await graph.createVoiceMeeting({...base,startIso:'2030-01-15T10:00:00-05:00'});
  await graph.createVoiceMeeting({...base,startIso:'2030-07-15T10:00:00-04:00',transactionId:'22345678-1234-4123-8123-123456789abc'});
  assert.deepEqual(payloads.map(p=>p.start),[{dateTime:'2030-01-15T10:00:00',timeZone:'Eastern Standard Time'},{dateTime:'2030-07-15T10:00:00',timeZone:'Eastern Standard Time'}]);
});

test('meeting rejects an offset that disagrees with Toronto daylight saving and preserves an explicit alternative timezone',async()=>{
  const payloads=[];const graph=client((url,options)=>{payloads.push(JSON.parse(options.body));return Response.json({id:'event'});});
  const base={title:'Call',durationMinutes:30,attendees:['person@example.com'],transactionId:'12345678-1234-4123-8123-123456789abc'};
  await assert.rejects(()=>graph.createVoiceMeeting({...base,startIso:'2030-07-15T10:00:00-05:00',timezone:'America/Toronto'}),/daylight-saving/);
  const result=await graph.createVoiceMeeting({...base,startIso:'2030-07-15T10:00:00+01:00',timezone:'Europe/London'});
  assert.deepEqual(payloads[0].start,{dateTime:'2030-07-15T10:00:00',timeZone:'Europe/London'});assert.equal(result.timezone,'Europe/London');
});

test('meeting requires preparation then explicit confirmation before Microsoft write',async()=>{
  let writes=0;const graph={listPrincipalCalendar:async()=>[],createVoiceMeeting:async proposal=>{writes++;return {id:'event',attendees:proposal.attendees};}};
  const context={graph,dropbox:{createDeliveryRecord:async()=>true},meetingProposals:new Map(),meetingRequests:new Set(),callKey:'call'};
  const prepared=await runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['jack@example.com']},context);
  assert.equal(prepared.created,false);assert.equal(prepared.requiresConfirmation,true);assert.equal(writes,0);
  await assert.rejects(()=>runVoiceTool('confirm_calendar_meeting',{proposal_id:prepared.proposal.proposalId,confirmed:false},context),/not explicitly confirmed/);
  assert.equal(writes,0);
  const created=await runVoiceTool('confirm_calendar_meeting',{proposal_id:prepared.proposal.proposalId,confirmed:true},context);
  assert.equal(created.created,true);assert.equal(created.invitationsSubmitted,true);assert.equal(writes,1);
});

test('meeting preparation reports conflicts and refuses guessed attendee names',async()=>{
  const graph={listPrincipalCalendar:async()=>[{id:'busy',subject:'Existing',showAs:'busy',start:{dateTime:'2030-09-11T14:00:00Z'},end:{dateTime:'2030-09-11T14:30:00Z'}}]};
  const context={graph,meetingProposals:new Map()};
  await assert.rejects(()=>runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['Jack Rawdon']},context),/exact attendee/);
  const result=await runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['jack@example.com']},context);
  assert.equal(result.conflicts.length,1);assert.equal(result.conflicts[0].subject,'Existing');
});

test('meeting attempt is durable and cannot be replayed',async()=>{
  let writes=0;const proposalId='12345678-1234-4123-8123-123456789abc';const proposal={proposalId,title:'Call',startIso:'2030-09-11T14:00:00.000Z',durationMinutes:30,timezone:'America/Toronto',attendees:['jack@example.com'],body:'',location:''};
  const context={graph:{createVoiceMeeting:async()=>{writes++;return {id:'event'};}},dropbox:{createDeliveryRecord:async()=>false},meetingProposals:new Map([[proposalId,proposal]]),meetingRequests:new Set(),callKey:'call'};
  await assert.rejects(()=>runVoiceTool('confirm_calendar_meeting',{proposal_id:proposalId,confirmed:true},context),/already attempted/);assert.equal(writes,0);
  assert.ok(voiceTools().some(tool=>tool.name==='prepare_calendar_meeting'));assert.ok(voiceTools().some(tool=>tool.name==='confirm_calendar_meeting'));
});

