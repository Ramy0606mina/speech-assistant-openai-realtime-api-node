import test from 'node:test';
import assert from 'node:assert/strict';
import { OpenAIClient } from '../src/openai-client.js';

test('owner calendar request reads actual events and supplies them to the reply',async()=>{
 let step=0;
 const client=new OpenAIClient({apiKey:'test',fetchImpl:async(url,o)=>{
  const body=JSON.parse(o.body);
  if(++step===1){assert.ok(body.tools.some(t=>t.name==='read_principal_calendar'));return Response.json({output:[{type:'function_call',call_id:'calendar',name:'read_principal_calendar',arguments:JSON.stringify({startIso:'2026-09-07T00:00:00-04:00',endIso:'2026-09-08T00:00:00-04:00'})}]});}
  assert.match(JSON.stringify(body.input),/Verified meeting/);return Response.json({output_text:'Verified meeting is on your calendar.'});
 }});
 const reply=await client.analyzeDelegatedEmail({subject:'Calendar'},[],{graph:{listPrincipalCalendar:async q=>{assert.equal(q.limit,50);return [{subject:'Verified meeting'}];}}});
 assert.match(reply.text,/Verified meeting/);
});
