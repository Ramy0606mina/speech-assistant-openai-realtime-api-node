import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {runVoiceTool,voiceTools,validTwilioRequest} from '../src/voice-gateway.js';
import {createHmac} from 'node:crypto';

const signature='<p><strong>RAMY MINA</strong><br>Founder &amp; Managing Principal - Mina Group</p>';
function client(handler,signatureHtml=signature) {
  return new MicrosoftGraphClient({readTenantId:'tenant',readClientId:'client',readClientSecret:'secret',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',signatureHtml,fetchImpl:async(url,opts)=>{
    const value=String(url).includes('login.microsoftonline.com') ? {access_token:'test',expires_in:3600} : await handler(String(url),opts);
    return {ok:true,text:async()=>JSON.stringify(value)};
  }});
}
test('new email goes into owner Drafts and never calls send',async()=>{
  const calls=[];const graph=client((url,opts)=>{calls.push({url,opts});return {id:'draft',isDraft:true,subject:'Review'};});
  const result=await graph.createVoiceDraft({to:['recipient@example.com'],subject:'Review',body:'A draft for review.'});
  assert.equal(result.sent,false);assert.equal(result.folder,'Drafts');
  assert.equal(calls.length,1);assert.match(calls[0].url,/users\/owner%40example.com\/messages$/);
  assert.equal(calls[0].opts.method,'POST');assert.deepEqual(JSON.parse(calls[0].opts.body).body,{contentType:'HTML',content:`<p>A draft for review.</p><div class="london-outlook-signature">${signature}</div>`});
});
test('reply uses native createReply preserving thread and Microsoft reply recipients',async()=>{
  const calls=[];const graph=client((url,opts)=>{calls.push({url,opts});return opts.method==='POST'?{id:'reply',isDraft:true,subject:'Re: Test'}:{id:'source',isDraft:false};});
  await graph.createVoiceDraft({mailbox:'principal',messageId:'source',body:'Thanks, I will review.'});
  assert.equal(calls.length,2);assert.match(calls[1].url,/messages\/source\/createReply$/);
  assert.deepEqual(JSON.parse(calls[1].opts.body),{message:{body:{contentType:'HTML',content:`<p>Thanks, I will review.</p><div class="london-outlook-signature">${signature}</div>`}}});
});
test('principal drafts fail closed when the stored signature is unavailable',async()=>{
  let called=false;const graph=client(()=>{called=true;return {id:'draft',isDraft:true};},'');
  await assert.rejects(()=>graph.createVoiceDraft({to:['recipient@example.com'],subject:'Review',body:'A draft for review.'}),/signature is not configured/);
  assert.equal(called,false);
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
  const args={message_id:'source',body:'Hello Julie,\n\nThank you for your message. I will review it shortly.\n\nKind regards,'};
  await assert.rejects(()=>runVoiceTool('save_email_draft',args,ctx),/Read the selected/);
  await runVoiceTool('read_email',{message_id:'source'},ctx);
  await runVoiceTool('save_email_draft',args,ctx);
  await assert.rejects(()=>runVoiceTool('save_email_draft',args,ctx),/already attempted/);
  assert.equal(created,1);
});
test('failed durable claim blocks email writes and no email-send tool exists',async()=>{
  let created=false;
  await assert.rejects(()=>runVoiceTool('save_email_draft',{to:['x@example.com'],body:'Hello,\n\nThank you for your message.\n\nRegards,',subject:'S'},{graph:{createVoiceDraft:()=>created=true},dropbox:{createDeliveryRecord:async()=>false},callKey:'call'}),/already attempted/);
  assert.equal(created,false);
  assert.equal(voiceTools().some(t=>/send_email|approve_email/.test(t.name)),false);
});
test('voice refuses an unformatted email block before saving or claiming it',async()=>{
  let created=false,claimed=false;
  await assert.rejects(()=>runVoiceTool('save_email_draft',{to:['x@example.com'],subject:'Update',body:'hello here is the update please review it regards ramy'},{graph:{createVoiceDraft:()=>created=true},dropbox:{createDeliveryRecord:async()=>{claimed=true;return true;}},callKey:'call'}),/polished plain text/);
  assert.equal(created,false);assert.equal(claimed,false);
});
test('voice leaves the sender name to London’s stored Outlook signature',async()=>{
  let created=false;
  await assert.rejects(()=>runVoiceTool('save_email_draft',{to:['x@example.com'],subject:'Update',body:'Hello,\n\nThank you for your message.\n\nBest regards,\nRamy Mina'},{graph:{createVoiceDraft:()=>created=true},dropbox:{createDeliveryRecord:async()=>true},callKey:'call'}),/Do not add a sender name/);
  assert.equal(created,false);
});
test('Twilio signature rejects spoofed callers and altered fields',()=>{
  const request={method:'POST',url:'/incoming-call',body:{From:'+15145550000'},headers:{}};
  assert.equal(validTwilioRequest(request,'key','https://example.com'),false);
  request.headers['x-twilio-signature']=createHmac('sha1','key').update('https://example.com/incoming-callFrom+15145550000').digest('base64');
  assert.equal(validTwilioRequest(request,'key','https://example.com'),true);
  request.body.From='+15145550001';assert.equal(validTwilioRequest(request,'key','https://example.com'),false);
});


