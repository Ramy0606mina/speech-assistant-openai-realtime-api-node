import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';

function fixture({wrongOwner=false,wrongThread=false,wrongRecipients=false,sendFails=false}={}) {
 const calls=[];
 const graph=new MicrosoftGraphClient({readTenantId:'tenant',readClientId:'client',readClientSecret:'secret',londonMailbox:'london@example.com',ramyMailbox:'owner@example.com',fetchImpl:async(url,options={})=>{
  const path=new URL(url).pathname;calls.push({path,...options});
  if(path.includes('oauth2'))return Response.json({access_token:'test',expires_in:3600});
  if(path.endsWith('/source')&&!options.method)return Response.json({id:'source',conversationId:'thread',from:{emailAddress:{address:wrongOwner?'other@example.com':'owner@example.com'}},replyTo:[{emailAddress:{address:'untrusted@example.com'}}]});
  if(path.endsWith('/createReply'))return Response.json({id:'draft',isDraft:true,conversationId:wrongThread?'other':'thread'});
  if(options.method==='PATCH'){
   const body=JSON.parse(options.body);assert.deepEqual(body.ccRecipients,[]);assert.deepEqual(body.bccRecipients,[]);assert.deepEqual(body.toRecipients,[{emailAddress:{address:'owner@example.com'}}]);assert.equal(body.body.content,'Reply confirm to send the invitation.');
   return Response.json({id:'draft',isDraft:true,conversationId:'thread',...body,...(wrongRecipients?{toRecipients:[{emailAddress:{address:'untrusted@example.com'}}]}:{})});
  }
  if(path.endsWith('/send')){if(sendFails)throw Error('send timeout');return new Response(null,{status:202});}
  assert.fail('Unexpected Graph call '+path);
 }});
 return {graph,calls};
}
const request={messageId:'source',conversationId:'thread',body:'Reply confirm to send the invitation.'};
test('meeting proposal stays in source conversation and overrides inherited recipients before send',async()=>{
 const f=fixture();assert.deepEqual(await f.graph.replyToLondonMessage(request),{sent:true,conversationId:'thread'});assert.equal(f.calls.filter(c=>c.path.endsWith('/send')).length,1);
});
for(const option of ['wrongOwner','wrongThread','wrongRecipients'])test('reply rejects '+option+' without sending',async()=>{
 const f=fixture({[option]:true});await assert.rejects(f.graph.replyToLondonMessage(request));assert.equal(f.calls.filter(c=>c.path.endsWith('/send')).length,0);
});
test('ambiguous reply send is surfaced without an automatic second send',async()=>{
 const f=fixture({sendFails:true});await assert.rejects(f.graph.replyToLondonMessage(request),/timeout/);assert.equal(f.calls.filter(c=>c.path.endsWith('/send')).length,1);
});
