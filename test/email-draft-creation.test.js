import test from 'node:test';
import assert from 'node:assert/strict';
import {EmailDraftCreation} from '../src/email-draft-creation.js';
import {StateStore} from '../src/state-store.js';

function fixture(){
 const records=new Set(),writes=[];
 const graph={principalMailbox:'owner@example.com',voiceMailbox:()=> 'owner@example.com',
  createVoiceDraft:async args=>{writes.push(args);return {id:'saved',isDraft:true,sent:false};},
  getVoiceMessage:async()=>({id:'saved',isDraft:true,subject:'Review',body:{content:'Hello Alex'},toRecipients:[{emailAddress:{address:'alex@example.com'}}]}),
  sendMail:async()=>assert.fail('Composed email must not be sent'),
  resolveVoiceContact:async()=>({status:'resolved',contacts:[{address:'alex@example.com'}]})};
 const queue=[],openai={respond:async()=>{assert.ok(queue.length,'Unexpected model call');return queue.shift();}};
 const dropbox={createDeliveryRecord:async key=>{if(records.has(key))return false;records.add(key);return true;}};
 const state=new StateStore(null),handler=new EmailDraftCreation({graph,openai,dropbox,state});
 const request={owner:graph.principalMailbox,requestKey:'mail-1',text:'Draft an email to alex@example.com asking for a review.'};
 const call=(name,args)=>({raw:{output:[{type:'function_call',call_id:'c'+queue.length,name,arguments:JSON.stringify(args)}]}});
 const save=()=>call('save_email_draft',{to:['alex@example.com'],subject:'Review',body:'Hello Alex,\n\nPlease review this.\n\nBest regards,'});
 return {handler,graph,queue,state,writes,request,call,save};
}
test('owner email saves an actual draft and establishes current revision context',async()=>{
 const f=fixture();f.queue.push(f.save());
 assert.match(await f.handler.handle(f.request),/Saved in Outlook Drafts/);
 assert.equal(f.writes.length,1);assert.equal(f.state.state.emailDraftRevision.current.id,'saved');
});
test('same owner email cannot save a duplicate on replay',async()=>{
 const f=fixture();f.queue.push(f.save(),f.save());await f.handler.handle(f.request);
 assert.match(await f.handler.handle(f.request),/not verified/);assert.equal(f.writes.length,1);
});
test('draft creation rejects unauthenticated input before any lookup or write',async()=>{
 const f=fixture();await assert.rejects(f.handler.handle({...f.request,owner:'other@example.com'}),/Authenticated/);assert.equal(f.writes.length,0);
});
test('existing draft revisions remain with the revision handler',async()=>{
 const f=fixture();assert.equal(await f.handler.handle({...f.request,text:'Shorten my existing Outlook draft'}),null);
 assert.equal(await f.handler.handle({...f.request,text:'Add a greeting to the current draft'}),null);
 assert.equal(await f.handler.handle({...f.request,text:'Draft a report of the costs'}),null);
 assert.equal(await f.handler.handle({...f.request,subject:'FW: Draft an email to Alex',text:'For your information.'}),null);
});
test('unverified model recipient never becomes a draft recipient',async()=>{
 const f=fixture();f.queue.push(f.call('save_email_draft',{to:['stranger@example.com'],subject:'Review',body:'Review please'}),{text:'What is the recipient address?'});
 assert.match(await f.handler.handle(f.request),/No Outlook draft/);assert.equal(f.writes.length,0);
});
test('unambiguous contact resolution authorizes a named recipient',async()=>{
 const f=fixture();f.queue.push(f.call('find_contact',{query:'Alex'}),f.save());
 assert.match(await f.handler.handle({...f.request,text:'Draft an email to Alex asking for review.'}),/Saved in Outlook/);
 assert.equal(f.writes.length,1);
});
test('ambiguous contact does not authorize any guessed address',async()=>{
 const f=fixture();f.graph.resolveVoiceContact=async()=>({status:'ambiguous',contacts:[{address:'alex@example.com'},{address:'alex2@example.com'}]});
 f.queue.push(f.call('find_contact',{query:'Alex'}),f.save(),{text:'Which Alex do you mean?'});
 assert.match(await f.handler.handle({...f.request,text:'Draft an email to Alex.'}),/Which Alex/);assert.equal(f.writes.length,0);
});
test('provider timeout yields uncertain result and no automatic duplicate',async()=>{
 const f=fixture();f.graph.createVoiceDraft=async args=>{f.writes.push(args);throw Error('timeout');};f.queue.push(f.save(),f.save());
 assert.match(await f.handler.handle(f.request),/not verified/);await f.handler.handle(f.request);assert.equal(f.writes.length,1);
});
test('save requires successful readback of a still-existing draft',async()=>{
 const f=fixture();f.graph.getVoiceMessage=async()=>({id:'saved',isDraft:false});f.queue.push(f.save());
 assert.match(await f.handler.handle(f.request),/not verified/);assert.equal(f.state.state.emailDraftRevision,null);
});
test('unsupported sending tool cannot send composed email',async()=>{
 const f=fixture();f.queue.push(f.call('send_mail',{to:'stranger@example.com'}),{text:'Provide the email details?'});
 assert.match(await f.handler.handle(f.request),/No Outlook draft/);assert.equal(f.writes.length,0);
});
