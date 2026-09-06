import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {runVoiceTool,voiceTools,validTwilioRequest} from '../src/voice-gateway.js';
import {createHmac} from 'node:crypto';

function client(handler) {
  return new MicrosoftGraphClient({readTenantId:'tenant',readClientId:'client',readClientSecret:'secret',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async(url,opts)=>{
    const value=String(url).includes('login.microsoftonline.com') ? {access_token:'test',expires_in:3600} : await handler(String(url),opts);
    return {ok:true,text:async()=>JSON.stringify(value)};
  }});
}
test('new email goes into owner Drafts and never calls send',async()=>{
  const calls=[];const graph=client((url,opts)=>{calls.push({url,opts});return {id:'draft',isDraft:true,subject:'Review'};});
  const result=await graph.createVoiceDraft({to:['recipient@example.com'],subject:'Review',body:'A draft for review.'});
  assert.equal(result.sent,false);assert.equal(result.folder,'Drafts');
  assert.equal(calls.length,1);assert.match(calls[0].url,/users\/owner%40example.com\/messages$/);
  assert.equal(calls[0].opts.method,'POST');assert.equal(JSON.parse(calls[0].opts.body).body.content,'A draft for review.');
});
test('reply uses native createReply preserving thread and Microsoft reply recipients',async()=>{
  const calls=[];const graph=client((url,opts)=>{calls.push({url,opts});return opts.method==='POST'?{id:'reply',isDraft:true,subject:'Re: Test'}:{id:'source',isDraft:false};});
  await graph.createVoiceDraft({mailbox:'principal',messageId:'source',body:'Thanks, I will review.'});
  assert.equal(calls.length,2);assert.match(calls[1].url,/messages\/source\/createReply$/);
  assert.deepEqual(JSON.parse(calls[1].opts.body),{comment:'Thanks, I will review.'});
});
test('rejects unconnected mailbox, guessed recipient, and unconfirmed draft',async()=>{
  const graph=client(()=>({id:'missing-draft-flag'}));
  await assert.rejects(()=>graph.createVoiceDraft({mailbox:'other',to:['x@example.com'],subject:'A',body:'B'}),/connected mailbox/);
  await assert.rejects(()=>graph.createVoiceDraft({to:['Anass'],subject:'A',body:'B'}),/explicit recipient/);
  await assert.rejects(()=>graph.createVoiceDraft({to:['x@example.com'],subject:'A',body:'B'}),/did not confirm/);
});
test('reply requires original read, and repeated draft attempts cannot duplicate',async()=>{
  let created=0;const graph={getVoiceMessage:async()=>({id:'source'}),createVoiceDraft:async()=>{created++;return {isDraft:true,sent:false};}};
  const ctx={graph,dropbox:{createDeliveryRecord:async()=>true},readMessages:new Set(),draftRequests:new Set(),callKey:'call'};
  const args={message_id:'source',body:'Reply'};
  await assert.rejects(()=>runVoiceTool('save_email_draft',args,ctx),/Read the selected/);
  await runVoiceTool('read_email',{message_id:'source'},ctx);
  await runVoiceTool('save_email_draft',args,ctx);
  await assert.rejects(()=>runVoiceTool('save_email_draft',args,ctx),/already attempted/);
  assert.equal(created,1);
});
test('failed durable claim blocks writes and no send tools exist',async()=>{
  let created=false;
  await assert.rejects(()=>runVoiceTool('save_email_draft',{to:['x@example.com'],body:'B',subject:'S'},{graph:{createVoiceDraft:()=>created=true},dropbox:{createDeliveryRecord:async()=>false},callKey:'call'}),/already attempted/);
  assert.equal(created,false);
  assert.equal(voiceTools().some(t=>/send|approve|create_event/.test(t.name)),false);
});
test('Twilio signature rejects spoofed callers and altered fields',()=>{
  const request={method:'POST',url:'/incoming-call',body:{From:'+15145550000'},headers:{}};
  assert.equal(validTwilioRequest(request,'key','https://example.com'),false);
  request.headers['x-twilio-signature']=createHmac('sha1','key').update('https://example.com/incoming-callFrom+15145550000').digest('base64');
  assert.equal(validTwilioRequest(request,'key','https://example.com'),true);
  request.body.From='+15145550001';assert.equal(validTwilioRequest(request,'key','https://example.com'),false);
});
