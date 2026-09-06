import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {OpenAIClient} from '../src/openai-client.js';
test('follow-up uses only the existing principal register, private free time, no attendees, stable transaction',async()=>{
 const created=[];
 const g=new MicrosoftGraphClient({readTenantId:'test',readClientId:'test',readClientSecret:'test',ramyMailbox:'owner@example.com',fetchImpl:async(url,o)=>{
  if(String(url).includes('oauth2'))return Response.json({access_token:'test',expires_in:3600});
  assert.match(String(url),/users\/owner%40example.com/);
  if(String(url).includes('/calendars?'))return Response.json({value:[{id:'register',name:'London Action Register',owner:{address:'owner@example.com'}}]});
  assert.match(String(url),/calendars\/register\/events$/);created.push(JSON.parse(o.body));return Response.json({id:'saved'});
 }});
 const task={title:'Review test',date:'2026-09-07',taskKey:'source',notes:'Test only'};
 await g.createFollowUp(task);await g.createFollowUp(task);
 assert.equal(created[0].transactionId,created[1].transactionId);assert.deepEqual(created[0].attendees,[]);
 assert.equal(created[0].isReminderOn,false);assert.equal(created[0].showAs,'free');assert.equal(created[0].sensitivity,'private');
 await assert.rejects(g.createFollowUp({...task,date:'2026-02-30'}),/valid date/);
});
test('model prepares a task without changing calendar while reasoning',async()=>{
 let step=0,created=0;
 const ai=new OpenAIClient({apiKey:'test',fetchImpl:async()=>++step===1?Response.json({output:[{type:'function_call',call_id:'task',name:'prepare_follow_up',arguments:JSON.stringify({title:'Review test',date:'2026-09-07',notes:''})}]}):Response.json({output_text:'Follow-up details prepared.'})});
 const r=await ai.analyzeDelegatedEmail({body:{content:'Create follow-up on September 7'}},[],{graph:{createFollowUp:async()=>created++}});
 assert.equal(created,0);assert.equal(r.followUps.length,1);assert.equal(r.followUps[0].date,'2026-09-07');
});

test('calendar reads use the action app that owns calendar access',async()=>{
 let identity;
 const g=new MicrosoftGraphClient({readTenantId:'test',readClientId:'mail-reader',readClientSecret:'test',actionTenantId:'test',actionClientId:'calendar-actions',actionClientSecret:'test',ramyMailbox:'owner@example.com',fetchImpl:async(url,o)=>{
  if(String(url).includes('oauth2')){identity=new URLSearchParams(o.body).get('client_id');return Response.json({access_token:identity,expires_in:3600});}
  assert.equal(o.headers.Authorization,'Bearer calendar-actions');return Response.json({value:[]});
 }});
 assert.deepEqual(await g.listPrincipalCalendar({startIso:'2026-09-07T00:00:00Z',endIso:'2026-09-08T00:00:00Z'}),[]);
 assert.equal(identity,'calendar-actions');
});
