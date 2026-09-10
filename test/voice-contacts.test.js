import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {realtimeInstructions,runVoiceTool,voiceTools} from '../src/voice-gateway.js';

function graphWith(handler){return new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'s',actionTenantId:'t',actionClientId:'c',actionClientSecret:'s',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async(url,options)=>String(url).includes('oauth2')?Response.json({access_token:'token',expires_in:3600}):handler(new URL(url),options)});}

test('contact lookup searches a nested person folder and resolves from message evidence',async()=>{
  const visited=[];const graph=graphWith(url=>{visited.push(decodeURIComponent(url.pathname));
    if(url.pathname.endsWith('/mailFolders/inbox'))return Response.json({id:'inbox-id',displayName:'Inbox',childFolderCount:1});
    if(url.pathname.includes('/mailFolders/inbox-id/childFolders'))return Response.json({value:[{id:'project-id',displayName:'PROJECTS',childFolderCount:1}]});
    if(url.pathname.includes('/mailFolders/project-id/childFolders'))return Response.json({value:[{id:'jack-id',displayName:'JACK',childFolderCount:0}]});
    if(url.pathname.includes('/mailFolders/jack-id/messages'))return Response.json({value:[{subject:'Engagement Letter',from:{emailAddress:{name:'Jack Rawdon',address:'jack.rawdon@example.com'}},toRecipients:[{emailAddress:{name:'Owner',address:'owner@example.com'}}],receivedDateTime:'2026-09-04T03:53:53Z'}]});
    if(url.pathname.endsWith('/messages'))return Response.json({value:[]});throw new Error(`Unexpected ${url}`);
  });
  const result=await graph.resolveVoiceContact({query:'Jack Rodden',context:'Engagement Letter'});
  assert.equal(result.status,'resolved');assert.equal(result.contacts[0].address,'jack.rawdon@example.com');assert.ok(result.contacts[0].folders.some(f=>f.includes('PROJECTS/JACK')));assert.ok(visited.some(p=>p.includes('project-id/childFolders')));
});

test('contact lookup reports ambiguity when the same spoken name maps to two addresses',async()=>{
  const graph=graphWith(url=>{
    if(url.pathname.endsWith('/mailFolders/inbox'))return Response.json({id:'inbox',displayName:'Inbox',childFolderCount:0});
    if(url.pathname.endsWith('/messages'))return Response.json({value:[
      {from:{emailAddress:{name:'Alex Smith',address:'alex.one@example.com'}},receivedDateTime:'2026-09-04T00:00:00Z'},
      {from:{emailAddress:{name:'Alex Smith',address:'alex.two@example.com'}},receivedDateTime:'2026-09-03T00:00:00Z'},
    ]});throw new Error(`Unexpected ${url}`);
  });
  const result=await graph.resolveVoiceContact({query:'Alex Smith'});assert.equal(result.status,'ambiguous');assert.equal(result.contacts.length,2);
});

test('contact lookup never invents an address when evidence is absent',async()=>{
  const graph=graphWith(url=>url.pathname.endsWith('/mailFolders/inbox')?Response.json({id:'inbox',displayName:'Inbox',childFolderCount:0}):Response.json({value:[]}));
  const result=await graph.resolveVoiceContact({query:'Unknown Person'});assert.equal(result.status,'not_found');assert.deepEqual(result.contacts,[]);
});

test('voice exposes automatic Outlook contact lookup',async()=>{
  assert.ok(voiceTools().some(tool=>tool.name==='find_contact'));
  const graph={resolveVoiceContact:async args=>{assert.equal(args.query,'Jack Rodden');return {status:'resolved',contacts:[{name:'Jack Rawdon',address:'jack@example.com'}],foldersSearched:3,messagesScanned:20};}};
  const result=await runVoiceTool('find_contact',{query:'Jack Rodden'},{graph});assert.equal(result.status,'resolved');assert.equal(result.contacts[0].address,'jack@example.com');
});

test('voice requires nested-folder lookup before claiming an attendee is missing',()=>{
  const instructions=realtimeInstructions();
  const contactTool=voiceTools().find(tool=>tool.name==='find_contact');
  assert.match(instructions,/MUST call find_contact before asking for an address/);
  assert.match(instructions,/Never say that you lack access to email subfolders/);
  assert.match(contactTool.description,/every nested person\/project folder/);
});




test('contact already read on this call resolves without scanning folders',async()=>{
 const context={knownContacts:new Map(),readMessages:new Set(),graph:{getVoiceMessage:async()=>({from:{emailAddress:{name:'Christine Normando',address:'christine@example.com'}}}),resolveVoiceContact:async()=>assert.fail('unnecessary mailbox scan')}};
 await runVoiceTool('read_email',{message_id:'selected'},context);
 const found=await runVoiceTool('find_contact',{query:'Christine Normando'},context);
 assert.equal(found.status,'resolved');assert.equal(found.contacts[0].address,'christine@example.com');assert.equal(found.messagesScanned,0);
});

test('cached contact ambiguity is preserved and unknown names use full lookup',async()=>{
 let searches=0;const context={knownContacts:new Map([['a',{name:'Alex Smith',address:'one@example.com'}],['b',{name:'Alex Smith',address:'two@example.com'}]]),graph:{resolveVoiceContact:async()=>{searches++;return {status:'not_found',contacts:[]};}}};
 assert.equal((await runVoiceTool('find_contact',{query:'Alex Smith'},context)).status,'ambiguous');assert.equal(searches,0);
 await runVoiceTool('find_contact',{query:'Unknown Name'},context);assert.equal(searches,1);
});

test('voice acknowledges work first and uses selected reply without contact lookup',()=>{
 const text=realtimeInstructions();assert.match(text,/immediately say one short acknowledgment/);assert.match(text,/Replies to a selected email do not require contact lookup/);assert.match(text,/only if it has not already been read/);
});
