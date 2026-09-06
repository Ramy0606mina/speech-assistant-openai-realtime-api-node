import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {runVoiceTool,voiceTools} from '../src/voice-gateway.js';

const actionToken=`x.${Buffer.from(JSON.stringify({roles:['Calendars.ReadWrite','OnlineMeetings.ReadWrite.All']})).toString('base64url')}.x`;
function client(handler){return new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'s',actionTenantId:'t',actionClientId:'c',actionClientSecret:'s',ramyMailbox:'owner@example.com',ramyUserId:'11111111-1111-4111-8111-111111111111',londonMailbox:'london@example.com',fetchImpl:async(url,options)=>String(url).includes('oauth2')?Response.json({access_token:actionToken,expires_in:3600}):handler(String(url),options)});}

test('Microsoft meeting includes attendee invitations and a stable transaction',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});return Response.json({id:'event-1'});});
  const result=await graph.createVoiceMeeting({title:'Project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['JACK@example.com'],body:'Discuss next steps.',location:'Office',transactionId:'12345678-1234-4123-8123-123456789abc'});
  assert.equal(result.invitationsSubmitted,true);assert.equal(calls.length,1);assert.match(calls[0].url,/owner%40example.com\/events$/);
  const body=JSON.parse(calls[0].options.body);assert.equal(body.attendees[0].emailAddress.address,'jack@example.com');assert.deepEqual(body.start,{dateTime:'2030-07-11T10:00:00',timeZone:'Eastern Standard Time'});assert.deepEqual(body.end,{dateTime:'2030-07-11T10:30:00',timeZone:'Eastern Standard Time'});assert.equal(body.transactionId,'12345678-1234-4123-8123-123456789abc');
  assert.equal('isOnlineMeeting' in body,false);assert.equal('onlineMeetingProvider' in body,false);assert.equal(result.onlineMeeting,false);assert.equal(result.joinLinkCreated,false);
  assert.equal(result.timezone,'America/Toronto');assert.equal(result.microsoftTimeZone,'Eastern Standard Time');
});

test('Microsoft virtual meeting requests Teams and verifies the returned joining link without exposing it',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});return url.includes('/calendar?')?Response.json({allowedOnlineMeetingProviders:['teamsForBusiness']}):Response.json({id:'event-online',isOnlineMeeting:true,onlineMeetingProvider:'teamsForBusiness',onlineMeeting:{joinUrl:'https://teams.microsoft.com/meet/example'}});});
  const result=await graph.createVoiceMeeting({title:'Virtual project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['person@example.com'],onlineMeeting:true,transactionId:'32345678-1234-4123-8123-123456789abc'});
  const body=JSON.parse(calls[1].options.body);
  assert.equal(body.isOnlineMeeting,true);assert.equal(body.onlineMeetingProvider,'teamsForBusiness');
  assert.equal(result.onlineMeeting,true);assert.equal(result.onlineMeetingProvider,'teamsForBusiness');assert.equal(result.joinLinkCreated,true);assert.equal('joinUrl' in result,false);
});

test('Microsoft re-reads an online event when creation response omits joining information',async()=>{
  let call=0;const graph=client((url,options)=>{call++;if(url.includes('/calendar?'))return Response.json({allowedOnlineMeetingProviders:['teamsForBusiness']});return call===2?Response.json({id:'event-online'}):Response.json({id:'event-online',isOnlineMeeting:true,onlineMeetingProvider:'teamsForBusiness',onlineMeeting:{joinUrl:'https://teams.microsoft.com/meet/example'}});});
  const result=await graph.createVoiceMeeting({title:'Virtual project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['person@example.com'],onlineMeeting:true,transactionId:'42345678-1234-4123-8123-123456789abc'});
  assert.equal(call,3);assert.equal(result.joinLinkCreated,true);
});

test('Microsoft upgrades and verifies an event when Teams data remains absent after creation',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});if(url.includes('/calendar?'))return Response.json({allowedOnlineMeetingProviders:['teamsForBusiness']});if(calls.length<5)return Response.json({id:'event-online',isOnlineMeeting:false,onlineMeetingProvider:'unknown'});return Response.json({id:'event-online',isOnlineMeeting:true,onlineMeetingProvider:'teamsForBusiness',onlineMeeting:{joinUrl:'https://teams.microsoft.com/meet/example'}});});
  const result=await graph.createVoiceMeeting({title:'Virtual project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['person@example.com'],onlineMeeting:true,transactionId:'52345678-1234-4123-8123-123456789abc'});
  assert.equal(calls.length,5);assert.equal(calls[3].options.method,'PATCH');assert.deepEqual(JSON.parse(calls[3].options.body),{isOnlineMeeting:true,onlineMeetingProvider:'teamsForBusiness'});assert.equal(result.joinLinkCreated,true);
});

test('Microsoft refuses to report success when a Teams link cannot be verified',async()=>{
  const graph=client(url=>url.includes('/calendar?')?Response.json({allowedOnlineMeetingProviders:['teamsForBusiness']}):Response.json({id:'event-online',isOnlineMeeting:false,onlineMeetingProvider:'unknown'}));
  await assert.rejects(()=>graph.createVoiceMeeting({title:'Virtual project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['person@example.com'],onlineMeeting:true,transactionId:'62345678-1234-4123-8123-123456789abc'}),/did not create a Teams joining link/);
});

test('Microsoft creates an idempotent standalone Teams meeting when the calendar provider is unavailable',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});if(url.includes('/calendar?'))return Response.json({allowedOnlineMeetingProviders:[]});if(url.includes('/onlineMeetings/createOrGet'))return Response.json({joinWebUrl:'https://teams.microsoft.com/l/meetup-join/example'});return Response.json({id:'event-fallback'});});
  const result=await graph.createVoiceMeeting({title:'Virtual project call',startIso:'2030-07-11T10:00:00-04:00',durationMinutes:30,timezone:'America/Toronto',attendees:['person@example.com'],body:'Discuss.',onlineMeeting:true,transactionId:'72345678-1234-4123-8123-123456789abc'});
  const onlineCall=calls.find(call=>call.url.includes('/onlineMeetings/createOrGet'));const onlineBody=JSON.parse(onlineCall.options.body);assert.equal(onlineBody.externalId,'72345678-1234-4123-8123-123456789abc');
  assert.match(onlineCall.url,/users\/11111111-1111-4111-8111-111111111111\/onlineMeetings/);
  const eventBody=JSON.parse(calls.at(-1).options.body);assert.equal(eventBody.body.contentType,'HTML');assert.match(eventBody.body.content,/Join Microsoft Teams Meeting/);assert.equal(result.joinLinkCreated,true);
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
  const prepared=await runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['jack@example.com'],online_meeting:true},context);
  assert.equal(prepared.created,false);assert.equal(prepared.requiresConfirmation,true);assert.equal(writes,0);
  assert.equal(prepared.proposal.onlineMeeting,true);assert.equal(prepared.proposal.onlineMeetingProvider,'Microsoft Teams');
  await assert.rejects(()=>runVoiceTool('confirm_calendar_meeting',{proposal_id:prepared.proposal.proposalId,confirmed:false},context),/not explicitly confirmed/);
  assert.equal(writes,0);
  const created=await runVoiceTool('confirm_calendar_meeting',{proposal_id:prepared.proposal.proposalId,confirmed:true},context);
  assert.equal(created.created,true);assert.equal(created.invitationsSubmitted,true);assert.equal(writes,1);
});

test('meeting preparation reports conflicts and refuses guessed attendee names',async()=>{
  const graph={listPrincipalCalendar:async()=>[{id:'busy',subject:'Existing',showAs:'busy',start:{dateTime:'2030-09-11T14:00:00Z'},end:{dateTime:'2030-09-11T14:30:00Z'}}]};
  const context={graph,meetingProposals:new Map()};
  await assert.rejects(()=>runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['Jack Rawdon'],online_meeting:false},context),/exact attendee/);
  const result=await runVoiceTool('prepare_calendar_meeting',{title:'Call Jack',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['jack@example.com'],online_meeting:false},context);
  assert.equal(result.conflicts.length,1);assert.equal(result.conflicts[0].subject,'Existing');
});

test('meeting preparation treats spoken virtual mode as Teams even if the boolean is wrong',async()=>{
  const context={graph:{listPrincipalCalendar:async()=>[],previewVoiceMeeting:()=>({localStart:'2030-09-11T10:00:00',localEnd:'2030-09-11T10:30:00',timezone:'America/Toronto'})},meetingProposals:new Map()};
  const result=await runVoiceTool('prepare_calendar_meeting',{title:'Virtual Meeting',start_iso:'2030-09-11T10:00:00-04:00',timezone:'America/Toronto',duration_minutes:30,attendees:['person@example.com'],online_meeting:false},context);
  assert.equal(result.proposal.onlineMeeting,true);assert.equal(result.proposal.onlineMeetingProvider,'Microsoft Teams');
});

test('meeting attempt is durable and cannot be replayed',async()=>{
  let writes=0;const proposalId='12345678-1234-4123-8123-123456789abc';const proposal={proposalId,title:'Call',startIso:'2030-09-11T14:00:00.000Z',durationMinutes:30,timezone:'America/Toronto',attendees:['jack@example.com'],body:'',location:''};
  const context={graph:{createVoiceMeeting:async()=>{writes++;return {id:'event'};}},dropbox:{createDeliveryRecord:async()=>false},meetingProposals:new Map([[proposalId,proposal]]),meetingRequests:new Set(),callKey:'call'};
  await assert.rejects(()=>runVoiceTool('confirm_calendar_meeting',{proposal_id:proposalId,confirmed:true},context),/already attempted/);assert.equal(writes,0);
  assert.ok(voiceTools().some(tool=>tool.name==='prepare_calendar_meeting'));assert.ok(voiceTools().some(tool=>tool.name==='confirm_calendar_meeting'));
});

test('Microsoft cancellation sends the organizer note to attendees',async()=>{
  const calls=[];const graph=client((url,options)=>{calls.push({url,options});return new Response(null,{status:202});});
  const result=await graph.cancelVoiceMeeting({eventId:'event-1',comment:'Schedule changed.'});
  assert.equal(result.cancelled,true);assert.match(calls[0].url,/events\/event-1\/cancel$/);assert.equal(calls[0].options.method,'POST');assert.deepEqual(JSON.parse(calls[0].options.body),{comment:'Schedule changed.'});
});

test('calendar cancellation requires selection and a later explicit confirmation',async()=>{
  let cancellations=0;const event={id:'event-1',subject:'Project review',start:{dateTime:'2030-09-11T10:00:00'},end:{dateTime:'2030-09-11T10:30:00'},organizer:{emailAddress:{address:'owner@example.com'}},attendees:[{emailAddress:{name:'Jack',address:'jack@example.com'}}],isOrganizer:true,isCancelled:false};
  const context={graph:{listPrincipalCalendar:async()=>[event],cancelVoiceMeeting:async()=>{cancellations++;return {cancelled:true,cancellationSent:true};}},dropbox:{createDeliveryRecord:async()=>true},calendarEvents:new Map(),cancellationProposals:new Map(),cancellationRequests:new Set(),callKey:'call'};
  await assert.rejects(()=>runVoiceTool('prepare_calendar_cancellation',{event_id:'event-1'},context),/check_calendar/);
  const checked=await runVoiceTool('check_calendar',{start_iso:'2030-09-11T00:00:00-04:00',end_iso:'2030-09-12T00:00:00-04:00'},context);assert.equal(checked.events[0].isOrganizer,true);
  const prepared=await runVoiceTool('prepare_calendar_cancellation',{event_id:'event-1',comment:'Schedule changed.'},context);assert.equal(prepared.cancelled,false);assert.equal(cancellations,0);
  await assert.rejects(()=>runVoiceTool('confirm_calendar_cancellation',{proposal_id:prepared.proposal.proposalId,confirmed:false},context),/not explicitly confirmed/);
  const cancelled=await runVoiceTool('confirm_calendar_cancellation',{proposal_id:prepared.proposal.proposalId,confirmed:true},context);assert.equal(cancelled.cancelled,true);assert.equal(cancellations,1);
});

test('calendar cancellation refuses meetings Ramy does not organize',async()=>{
  const context={calendarEvents:new Map([['event-2',{id:'event-2',subject:'External meeting',isOrganizer:false,isCancelled:false}]]),cancellationProposals:new Map()};
  await assert.rejects(()=>runVoiceTool('prepare_calendar_cancellation',{event_id:'event-2'},context),/not the organizer/);
  assert.ok(voiceTools().some(tool=>tool.name==='prepare_calendar_cancellation'));assert.ok(voiceTools().some(tool=>tool.name==='confirm_calendar_cancellation'));
});

