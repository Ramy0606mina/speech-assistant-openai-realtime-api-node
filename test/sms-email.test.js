import test from 'node:test';
import assert from 'node:assert/strict';
import {SmsEmail,requestsEmailDraft} from '../src/sms-email.js';
import {answerOwnerSms,SmsConversation} from '../src/sms-conversation.js';
import {StateStore} from '../src/state-store.js';
import {DeliveryGuard} from '../src/delivery-guard.js';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';

const owner='owner@example.com';
const body='Hello Julie,\n\nPlease review the revised budget.\n\nBest regards,';
const call=(name,args)=>({raw:{output:[{type:'function_call',call_id:'call',name,arguments:JSON.stringify(args)}]}});
function fixture() {
  const records=new Map(),state=new StateStore(null),writes=[];
  let message={id:'saved-draft',isDraft:true,subject:'Budget',body:{contentType:'text',content:body},toRecipients:[{emailAddress:{address:'julie@example.com'}}],changeKey:'v1'};
  const graph={principalMailbox:owner,voiceMailbox:()=>owner,getVoiceMessage:async()=>structuredClone(message),
    createVoiceDraft:async args=>{writes.push(['create',args]);return {id:message.id,isDraft:true,mailbox:owner};},
    updateSmsDraft:async args=>{writes.push(['update',args]);message={...message,body:{contentType:'text',content:args.body},changeKey:'v2'};return {id:message.id,isDraft:true};},
    sendSmsDraft:async args=>{writes.push(['send',args]);message.isDraft=false;return {accepted:true};}};
  const dropbox={readDeliveryRecord:async key=>records.get(key),createDeliveryRecord:async(key,value)=>{if(records.has(key))return false;records.set(key,value);return true;}};
  const emails=new SmsEmail({graph,dropbox,state});
  const request=(text,key='second',history=[])=>({body:text,owner,requestKey:key,receivedAt:new Date(Date.now()+5000).toISOString(),history});
  async function saved() {
    const source=request('Draft the budget email','first');
    const reply=emails.remember(message,source);
    return [{role:'user',content:source.body,requestKey:'first'},{role:'assistant',content:reply}];
  }
  return {records,state,writes,graph,dropbox,emails,request,saved,get message(){return message;}};
}

test('email intent accepts normal drafting wording but excludes read-only and negated requests',()=>{
  for (const text of ['I want you to draft it in my email box as usual','Prepare an email to Julie','Send an email to Julie','Reply to Julie','I want you to respond to Julie','Save this in Outlook Drafts']) assert.equal(requestsEmailDraft(text),true,text);
  for (const text of ['Where is my draft?','Show my drafts','Do not draft an email','What is in my inbox?']) assert.equal(requestsEmailDraft(text),false,text);
});

test('SMS creates a real draft, reads it back, and remembers its exact ID without sending',async()=>{
  const f=fixture();
  const reply=await answerOwnerSms({...f.request('Prepare an email to julie@example.com','first'),...f,openai:{respond:async()=>call('save_email_draft',{to:['julie@example.com'],subject:'Budget',body})}});
  assert.match(reply,/Saved in Outlook Drafts/);assert.equal(f.emails.pending.id,'saved-draft');
  assert.doesNotMatch(reply,/\bsend\b|when ready/i);assert.deepEqual(f.writes.map(x=>x[0]),['create']);
});

test('recipient clarification and save-it follow-ups continue the authorized draft request',async()=>{
  const f=fixture(),first=f.request('Draft the budget email','first');
  assert.equal((await f.emails.prepare(first)).create,true);
  f.emails.recordAnswer(first,'Who should receive the email?');
  const history=[{role:'user',content:first.body,requestKey:'first'},{role:'assistant',content:'Who should receive the email?'}];
  assert.equal((await f.emails.prepare(f.request('julie@example.com','address',history))).create,true);
  assert.equal((await f.emails.prepare(f.request('Julie Martin','name',history))).create,true);
  assert.equal((await f.emails.prepare(f.request('save it','save',history))).create,true);
  assert.deepEqual(await f.emails.prepare(f.request('Show my calendar','calendar',history)),{});
});

test('recipient confirmation from the reported SMS screenshot retains the Outlook draft tool',async()=>{
  const f=fixture();
  const original=f.request('Draft an email to admin@minacapital.ca. Subject: Urgent Meeting. The meeting is tomorrow at 2 p.m. at the club.','original');
  const question='Should I address this email to admin@minacapital.ca?';
  const clarification=await answerOwnerSms({...original,...f,openai:{respond:async()=>({text:question})}});
  const history=[{role:'user',content:original.body,requestKey:original.requestKey},{role:'assistant',content:clarification}];
  const reply=await answerOwnerSms({...f.request("Yes, that's correct",'confirmation',history),...f,openai:{respond:async request=>{
    assert.ok(request.tools.some(tool=>tool.name==='save_email_draft'));
    return call('save_email_draft',{to:['admin@minacapital.ca'],subject:'Urgent Meeting',body:'Hello,\n\nPlease note that the meeting is tomorrow at 2:00 p.m. at the club.\n\nBest regards,'});
  }}});
  assert.match(reply,/Saved in Outlook Drafts/);
  assert.deepEqual(f.writes.map(item=>item[0]),['create']);
  assert.deepEqual(f.writes[0][1].to,['admin@minacapital.ca']);
  assert.equal(f.writes[0][1].subject,'Urgent Meeting');
});

test('keep it and acknowledgments never send; draft remains available in Outlook',async()=>{
  const f=fixture(),history=await f.saved();
  assert.match((await f.emails.prepare(f.request('Keep it in my draft box','keep',history))).reply,/Nothing was sent/);
  assert.deepEqual(f.writes,[]);
  for(const text of ['yes','confirm','okay','thanks','send it after changing the amount','do not send it']) {
    await f.emails.prepare(f.request(text,'noop',history));assert.equal(f.writes.length,0,text);
  }
});

test('SMS revises the saved body in place and preserves the original draft ID',async()=>{
  const f=fixture(),history=await f.saved();
  const revised='Hello Julie,\n\nPlease review the revised budget by Friday.\n\nBest regards,';
  const result=await answerOwnerSms({...f.request('Add that we need it by Friday','revise',history),...f,openai:{respond:async req=>{
    assert.ok(req.tools.some(t=>t.name==='update_saved_email_draft'));
    assert.ok(!req.tools.some(t=>t.name==='save_email_draft'));
    return call('update_saved_email_draft',{body:revised});
  }}});
  assert.match(result,/Saved in Outlook Drafts/);assert.equal(f.writes[0][0],'update');
  assert.equal(f.writes[0][1].id,'saved-draft');assert.equal(f.emails.pending.id,'saved-draft');
  assert.doesNotMatch(result,/\bsend\b|when ready/i);
});

test('explicit send commands are refused even with a saved draft and an old send invitation',async()=>{
  const f=fixture(),history=await f.saved();
  const oldReply=history.at(-1).content+' Review it in Outlook; reply send it when ready.';
  history.at(-1).content=oldReply;
  f.emails.save({...f.emails.pending,reply:oldReply,presentedAt:new Date(Date.now()-1000).toISOString()});
  for(const text of ['Please send it.','send it','Yes, send the email','go ahead send the draft now']) {
    const result=await f.emails.prepare(f.request(text,'send',history));
    assert.match(result.reply,/only saves email drafts/);assert.doesNotMatch(result.reply,/\bsend\b/);
  }
  assert.deepEqual(f.writes,[]);assert.equal(f.message.isDraft,true);
});

test('missing, stale, unrelated, pre-summary and foreign-owner confirmations never send',async()=>{
  const f=fixture(),history=await f.saved();
  const unrelated=[...history,{role:'user',content:'Book a meeting',requestKey:'meeting'},{role:'assistant',content:'Reply confirm.'}];
  for(const req of [f.request('send it'),f.request('send it','other',unrelated),{...f.request('send it','early',history),receivedAt:'2020-01-01T00:00:00Z'}]) {
    await f.emails.prepare(req);assert.equal(f.writes.length,0);
  }
  await assert.rejects(f.emails.prepare({...f.request('send it','foreign',history),owner:'other@example.com'}),/Authenticated/);
  f.emails.save({...f.emails.pending,updatedAt:'2020-01-01T00:00:00Z'});
  await f.emails.prepare(f.request('send it','old',history));assert.equal(f.writes.length,0);
});

test('an edited Outlook draft cannot enable sending',async()=>{
  const f=fixture(),history=await f.saved();f.message.bccRecipients=[{emailAddress:{address:'new@example.com'}}];
  const result=await f.emails.prepare(f.request('send it','changed',history));
  assert.match(result.reply,/only saves email drafts/);assert.equal(f.writes.length,0);
});

test('concurrent send commands and local-state loss never reach a send provider',async()=>{
  const f=fixture(),history=await f.saved();let attempts=0;
  f.graph.sendSmsDraft=async()=>{attempts++;throw Error('timeout');};
  const req=f.request('send it','send',history);
  await Promise.all([f.emails.prepare(req),f.emails.prepare(req)]);
  assert.equal(attempts,0);
  f.emails.cancelPending();const recoveredHistory=await f.saved();
  const result=await f.emails.prepare(f.request('send it','retry',recoveredHistory));
  assert.match(result.reply,/only saves email drafts/);assert.equal(attempts,0);
});

test('starting a new draft clears the previous draft reference before clarification',async()=>{
  const f=fixture(),history=await f.saved();
  await f.emails.prepare(f.request('Draft another email','new',history));
  assert.equal(f.emails.pending.id,undefined);
  await f.emails.prepare(f.request('send it','late',history));assert.equal(f.writes.length,0);
});

test('provider omission of a real saved ID never produces a successful draft receipt',async()=>{
  const f=fixture();f.graph.createVoiceDraft=async()=>({isDraft:true,mailbox:owner});
  const result=await answerOwnerSms({...f.request('Draft to julie@example.com'),...f,openai:{respond:async()=>call('save_email_draft',{to:['julie@example.com'],subject:'Budget',body})}});
  assert.match(result,/could not verify/);assert.equal(f.emails.pending.id,undefined);
});

test('a model-only saved claim is corrected and cannot be reported as an Outlook write',async()=>{
  const f=fixture();let calls=0;
  const result=await answerOwnerSms({...f.request('Draft an email to julie@example.com'),...f,openai:{respond:async()=>{calls++;return {text:'I saved it in Outlook Drafts. Anything else?'};}}});
  assert.equal(calls,2);assert.match(result,/No Outlook draft was saved/);assert.deepEqual(f.writes,[]);
});

test('SMS worker saves the draft and refuses its send follow-up before calendar/reminder parsers',async()=>{
  const f=fixture(),sent=[];let turn=0;
  const sms={configured:true,listIncoming:async()=>[{sid:'SM'+turn,body:turn?'send it':'Draft an email to julie@example.com about tomorrow',receivedAt:new Date(Date.now()+5000).toISOString()}],
    send:async text=>{sent.push(text);return {id:'reply'+turn,status:'queued'};},messageStatus:async()=>'delivered'};
  const worker=new SmsConversation({...f,sms,guard:new DeliveryGuard(f.dropbox,'sms'),
    meetings:{handle:async()=>assert.fail('Email reached meeting handler')},reminders:{handle:async()=>assert.fail('Email reached reminder handler')},
    openai:{respond:async()=>call('save_email_draft',{to:['julie@example.com'],subject:'Budget',body})}});
  await worker.tick();turn++;await worker.tick();
  assert.match(sent[0],/Saved in Outlook Drafts/);assert.doesNotMatch(sent[0],/\bsend\b/);
  assert.match(sent[1],/only saves email drafts/);
  assert.deepEqual(f.writes.map(x=>x[0]),['create']);
});

test('SMS and Graph expose no draft sending method',()=>{
  assert.equal(typeof SmsEmail.prototype.send,'undefined');
  assert.equal(typeof MicrosoftGraphClient.prototype.sendSmsDraft,'undefined');
});

test('model send invitations cannot reach SMS during drafting or after an old saved receipt',async()=>{
  for(const saved of [false,true]) {
    const f=fixture(),history=saved?await f.saved():[];
    const result=await answerOwnerSms({...f.request(saved?'Is it ready?':'Draft an email to Julie','request',history),...f,
      openai:{respond:async()=>({text:'Would you like me to send it?'})}});
    assert.match(result,/only saves email drafts/);assert.doesNotMatch(result,/\bsend\b/);assert.deepEqual(f.writes,[]);
  }
});
