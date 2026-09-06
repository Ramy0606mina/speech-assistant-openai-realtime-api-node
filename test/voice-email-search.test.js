import test from 'node:test';
import assert from 'node:assert/strict';
import {MicrosoftGraphClient} from '../src/microsoft-graph.js';
import {runVoiceTool,voiceTools} from '../src/voice-gateway.js';

function graphWith(handler){return new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'s',actionTenantId:'t',actionClientId:'c',actionClientSecret:'s',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async(url,options)=>String(url).includes('oauth2')?Response.json({access_token:'token',expires_in:3600}):handler(new URL(url),options)});}

test('search pages beyond ten messages and finds an older sender',async()=>{
  let pages=0;
  const graph=graphWith(url=>{
    pages++;
    if(pages===1){const next=new URL(url);next.pathname="/v1.0/users/owner@example.com/mailFolders('inbox')/messages";next.searchParams.set('$skiptoken','next');return Response.json({value:Array.from({length:50},(_,i)=>({id:`recent-${i}`,subject:'Other',from:{emailAddress:{name:'Other'}}})),'@odata.nextLink':next.href});}
    return Response.json({value:[{id:'eve',subject:'Project update',from:{emailAddress:{name:'Eve Trotman',address:'eve@example.com'}},receivedDateTime:'2026-08-20T12:00:00Z',bodyPreview:'Update attached.'}]});
  });
  const result=await graph.searchVoiceMessages({query:'Eve Trotman'});
  assert.equal(pages,2);assert.equal(result.scanned,51);assert.equal(result.complete,true);assert.equal(result.messages[0].id,'eve');
});

test('search reports incomplete rather than claiming absence at scan limit',async()=>{
  let pages=0;
  const graph=graphWith(url=>{pages++;url.searchParams.set('$skiptoken',String(pages));return Response.json({value:Array.from({length:50},(_,i)=>({id:`${pages}-${i}`,subject:'Other'})),'@odata.nextLink':url.href});});
  const result=await graph.searchVoiceMessages({query:'missing',maxScan:50});
  assert.equal(result.messages.length,0);assert.equal(result.scanned,50);assert.equal(result.complete,false);assert.equal(pages,1);
});

test('voice exposes deep email search and forwards folder and dates',async()=>{
  assert.ok(voiceTools().some(tool=>tool.name==='search_email'));
  const graph={searchVoiceMessages:async args=>{assert.deepEqual(args,{mailbox:'principal',folder:'deleteditems',query:'Eve',startIso:'2026-08-01T00:00:00Z',endIso:undefined});return {messages:[],scanned:200,complete:true};}};
  const result=await runVoiceTool('search_email',{folder:'deleteditems',query:'Eve',start_iso:'2026-08-01T00:00:00Z'},{graph});
  assert.deepEqual(result,{success:true,messages:[],scanned:200,complete:true});
});

test('search rejects invalid folders, ranges, and foreign paging links',async()=>{
  const graph=graphWith(()=>Response.json({value:[],'@odata.nextLink':'https://example.com/collect'}));
  await assert.rejects(()=>graph.searchVoiceMessages({folder:'junk',query:'Eve'}),/Unsupported/);
  await assert.rejects(()=>graph.searchVoiceMessages({query:'Eve',startIso:'bad'}),/date range/);
  await assert.rejects(()=>graph.searchVoiceMessages({query:'Eve'}),/paging was invalid/);
});

test('search rejects a Microsoft paging link for another mailbox',async()=>{
  const graph=graphWith(()=>Response.json({value:[],'@odata.nextLink':"https://graph.microsoft.com/v1.0/users/other@example.com/mailFolders('inbox')/messages?$skiptoken=x"}));
  await assert.rejects(()=>graph.searchVoiceMessages({query:'Eve'}),/paging was invalid/);
});

test('search suggests a close sender name after voice transcription changes the spelling',async()=>{
  const graph=graphWith(()=>Response.json({value:[
    {id:'jack',subject:'Engagement Letter and Fee Structure',from:{emailAddress:{name:'Jack Rawdon',address:'jack@example.com'}},receivedDateTime:'2026-09-03T12:00:00Z'},
    {id:'other',subject:'Routine update',from:{emailAddress:{name:'Jane Smith',address:'jane@example.com'}},receivedDateTime:'2026-09-04T12:00:00Z'},
  ]}));
  const result=await graph.searchVoiceMessages({query:'Jack Roden'});
  assert.equal(result.approximateMatch,true);
  assert.equal(result.suggestedSender,'Jack Rawdon');
  assert.equal(result.messages[0].id,'jack');
});

test('voice reports an approximate sender so London asks for confirmation',async()=>{
  const graph={searchVoiceMessages:async()=>({messages:[{id:'jack',subject:'Update',from:{emailAddress:{name:'Jack Rawdon'}}}],scanned:100,complete:true,approximateMatch:true,suggestedSender:'Jack Rawdon'})};
  const result=await runVoiceTool('search_email',{query:'Jack Roden'},{graph});
  assert.equal(result.approximateMatch,true);
  assert.equal(result.suggestedSender,'Jack Rawdon');
});

test('default sender search reaches a close voice transcription beyond one thousand messages',async()=>{
  let page=0;
  const graph=graphWith(url=>{
    page++;
    if(page<=20){
      const next=new URL(url);next.pathname="/v1.0/users/owner@example.com/mailFolders('inbox')/messages";next.searchParams.set('$skiptoken',String(page));
      return Response.json({value:Array.from({length:50},(_,i)=>({id:`${page}-${i}`,subject:'Other',from:{emailAddress:{name:'Someone Else'}}})),'@odata.nextLink':next.href});
    }
    return Response.json({value:[{id:'yves',subject:'Syncev follow-up',from:{emailAddress:{name:'Yves Trauttmann'}},receivedDateTime:'2026-09-04T13:01:48Z'}]});
  });
  const result=await graph.searchVoiceMessages({query:'Eve Trautman'});
  assert.equal(result.approximateMatch,true);
  assert.equal(result.suggestedSender,'Yves Trauttmann');
  assert.equal(result.messages[0].id,'yves');
  assert.equal(result.scanned,1001);
});

