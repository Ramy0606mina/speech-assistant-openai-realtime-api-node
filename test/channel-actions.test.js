import test from 'node:test';
import assert from 'node:assert/strict';
import {runVoiceTool,voiceTools} from '../src/voice-gateway.js';
import {answerOwnerSms} from '../src/sms-conversation.js';
function fixture(){
 const records=new Set(),writes=[];
 const context={callKey:'call-owner-1',graph:{principalMailbox:'owner@example.com',createPersonalReminder:async payload=>{writes.push(payload);return {id:'r1',created:true,reminderOn:true};}},dropbox:{createDeliveryRecord:async key=>{if(records.has(key))return false;records.add(key);return true;}}};
 const args={title:'Review invoice',start_iso:'2035-07-11T09:00:00-04:00'};
 return {context,args,writes};
}
test('phone personal reminder executes without a confirmation argument',async()=>{
 const f=fixture();assert.ok(voiceTools().some(t=>t.name==='create_personal_reminder'));
 const result=await runVoiceTool('create_personal_reminder',f.args,f.context);
 assert.equal(result.success,true);assert.equal(result.invitationsSubmitted,false);assert.equal(f.writes.length,1);
 assert.equal('attendees' in f.writes[0],false);
});
test('reminder replay is blocked durably across fresh call contexts',async()=>{
 const f=fixture();await runVoiceTool('create_personal_reminder',f.args,f.context);
 await assert.rejects(runVoiceTool('create_personal_reminder',f.args,{...f.context}),/already attempted/);assert.equal(f.writes.length,1);
});
test('phone reminder missing time or source never writes',async()=>{
 const f=fixture();await assert.rejects(runVoiceTool('create_personal_reminder',{title:'Review'},f.context),/exact time/);
 await assert.rejects(runVoiceTool('create_personal_reminder',f.args,{...f.context,callKey:''}),/Authenticated/);assert.equal(f.writes.length,0);
});
test('unverified phone reminder cannot return success',async()=>{
 const f=fixture();f.context.graph.createPersonalReminder=async()=>({id:'r1',created:true,reminderOn:false});
 await assert.rejects(runVoiceTool('create_personal_reminder',f.args,f.context),/did not verify/);
});

test('phone creates an Action Register task with durable source and no invitations',async()=>{
 const f=fixture(),tasks=[];f.context.graph.createFollowUp=async args=>{tasks.push(args);return {id:'task1',title:args.title,date:args.date,reminder:'None'};};
 const args={title:'Review quote',date:'2035-07-11',reminder:false};
 const result=await runVoiceTool('create_action_register',args,f.context);
 assert.equal(result.success,true);assert.equal(tasks[0].source,'London owner phone');assert.equal(tasks[0].reminder,false);
 await assert.rejects(runVoiceTool('create_action_register',args,f.context),/already attempted/);assert.equal(tasks.length,1);
});
test('invalid task date fails before provider access',async()=>{
 const f=fixture();f.context.graph.createFollowUp=async()=>assert.fail('invalid task');
 await assert.rejects(runVoiceTool('create_action_register',{title:'Review',date:'2035-02-30',reminder:false},f.context),/valid due date/);
});
test('SMS task tool is available only for owner creation requests and returns verified receipt',async()=>{
 const f=fixture();let writes=0;f.context.graph.createFollowUp=async args=>{writes++;assert.equal(args.source,'London owner SMS');return {id:'task1',title:args.title,date:args.date,reminder:'None'};};
 const openai={respond:async r=>{assert.ok(r.tools.some(t=>t.name==='create_action_register'));return {raw:{output:[{type:'function_call',call_id:'one',name:'create_action_register',arguments:JSON.stringify({title:'Review quote',date:'2035-07-11',reminder:false})}]}};}};
 const reply=await answerOwnerSms({...f.context,requestKey:'sms-1',body:'Create a task to review quote July 11 2035 without reminder',openai});
 assert.match(reply,/Created in London Action Register/);assert.equal(writes,1);
 await answerOwnerSms({...f.context,requestKey:'sms-2',body:'Read the Action Register',openai:{respond:async r=>{assert.ok(!r.tools.some(t=>t.name==='create_action_register'));return {text:'No tasks found.'};}}});
});

test('SMS status update reads selected task and preserves due date',async()=>{
 const f=fixture();let writes=0,round=0;
 f.context.graph.listFollowUps=async()=>[{id:'task1',title:'Review quote',status:'ACTIVE',nextFollowUp:'2035-07-11'}];
 f.context.graph.updateFollowUp=async args=>{writes++;assert.equal(args.nextFollowUp,'2035-07-11');return {id:args.id,title:'Review quote',status:args.status};};
 const openai={respond:async()=>({raw:{output:[{type:'function_call',call_id:'step'+round,name:round++===0?'read_action_register':'update_action_register',arguments:round===1?'{}':JSON.stringify({action_id:'task1',status:'COMPLETED',date:'2035-07-11'})}]}})};
 assert.match(await answerOwnerSms({...f.context,requestKey:'sms-update',body:'Mark the Review quote task completed',openai}),/COMPLETED/);assert.equal(writes,1);
});
test('SMS cannot update an unread task or change its due date',async()=>{
 const f=fixture();f.context.channel='sms';f.context.actionRecords=new Map();f.context.graph.updateFollowUp=async()=>assert.fail('must not update');
 await assert.rejects(runVoiceTool('update_action_register',{action_id:'task1',status:'COMPLETED',date:'2035-07-11'},f.context),/Read and select/);
 f.context.actionRecords.set('task1',{id:'task1',nextFollowUp:'2035-07-12'});
 await assert.rejects(runVoiceTool('update_action_register',{action_id:'task1',status:'COMPLETED',date:'2035-07-11'},f.context),/preserve/);
});

test('phone document analysis requires exact live file selection and supplies real content',async()=>{
 const f=fixture();f.context.channelState={files:new Map(),reads:new Map(),bytes:0};
 f.context.dropbox.search=async()=>[{'.tag':'file',path_display:'/LONDON - ACCESS/quote.txt',name:'quote.txt',size:20}];
 f.context.dropbox.readFile=async(path,max)=>{assert.equal(max,40*1024*1024);return {path,size:20,part:{type:'input_text',text:'Quoted total: $120.'}};};
 f.context.openai={respond:async r=>{assert.equal(r.tools,undefined);assert.match(r.input[0].content[1].text,/120/);return {text:'The quote total is $120 (quote.txt).'};}};
 await assert.rejects(runVoiceTool('read_dropbox_document',{path:'/LONDON - ACCESS/quote.txt',question:'Total?'},f.context),/Search or list/);
 await runVoiceTool('search_dropbox',{query:'quote'},f.context);
 const read=await runVoiceTool('read_dropbox_document',{path:'/LONDON - ACCESS/quote.txt',question:'Total?'},f.context);
 assert.equal(read.read,true);assert.match(read.analysis,/120/);assert.equal(f.context.channelState.bytes,20);
});
test('document limit and download failure cannot be reported as completed analysis',async()=>{
 const f=fixture();f.context.channelState={files:new Map([['/root/a.pdf',{}]]),reads:new Map(),bytes:40*1024*1024};
 f.context.openai={respond:async()=>assert.fail('no document')};f.context.dropbox.readFile=async()=>{throw Error('download failed');};
 await assert.rejects(runVoiceTool('read_dropbox_document',{path:'/root/a.pdf',question:'Analyze'},f.context),/limit/);
 f.context.channelState.bytes=0;
 await assert.rejects(runVoiceTool('read_dropbox_document',{path:'/root/a.pdf',question:'Analyze'},f.context),/download failed/);
});
test('SMS document request searches then analyzes the retrieved source before answering',async()=>{
 const f=fixture();let round=0,read=false;
 f.context.dropbox.search=async()=>[{'.tag':'file',path_display:'/root/quote.txt',name:'quote.txt'}];
 f.context.dropbox.readFile=async path=>{read=true;return {path,size:12,part:{type:'input_text',text:'Price 120'}};};
 const openai={respond:async r=>{
  if(!r.tools){assert.ok(read);return {text:'Price is 120 from quote.txt.'};}
  if(round++===0)return {raw:{output:[{type:'function_call',call_id:'find',name:'search_dropbox',arguments:'{"query":"quote"}'}]}};
  if(round===2)return {raw:{output:[{type:'function_call',call_id:'read',name:'read_dropbox_document',arguments:'{"path":"/root/quote.txt","question":"What is the price?"}'}]}};
  assert.ok(read);return {text:'The quote price is 120.'};
 }};
 assert.match(await answerOwnerSms({...f.context,openai,requestKey:'sms-doc',body:'Read the Dropbox quote and tell me the price'}),/120/);
});

test('phone/SMS rename verifies file identity and retains its folder and extension',async()=>{
 const f=fixture(),path='/root/old.pdf';let writes=0;
 f.context.channelState={files:new Map([[path,{id:'file1',path}]]),reads:new Map(),bytes:0};
 f.context.dropbox.getFileMetadata=async()=>({id:'file1','.tag':'file'});
 f.context.dropbox.renameFile=async(p,name)=>{writes++;assert.equal(p,path);assert.equal(name,'new.pdf');return {id:'file1',path:'/root/new.pdf',name};};
 const result=await runVoiceTool('rename_dropbox_file',{path,new_name:'new.pdf'},f.context);
 assert.equal(result.success,true);assert.equal(writes,1);assert.ok(f.context.channelState.files.has('/root/new.pdf'));
 await assert.rejects(runVoiceTool('rename_dropbox_file',{path,new_name:'new.pdf'},f.context),/Select an exact/);
});
test('rename rejects path traversal, extension changes, replaced files and ledger files',async()=>{
 const f=fixture(),path='/root/old.pdf';f.context.channelState={files:new Map([[path,{id:'file1'}]]),reads:new Map(),bytes:0};
 f.context.dropbox.renameFile=async()=>assert.fail('no rename');f.context.dropbox.getFileMetadata=async()=>({id:'replacement','.tag':'file'});
 for(const new_name of ['../other.pdf','new.docx'])await assert.rejects(runVoiceTool('rename_dropbox_file',{path,new_name},f.context),/same folder/);
 await assert.rejects(runVoiceTool('rename_dropbox_file',{path,new_name:'new.pdf'},f.context),/changed/);
 f.context.channelState.files.set('/root/.london-delivery/a.json',{id:'ledger'});
 await assert.rejects(runVoiceTool('rename_dropbox_file',{path:'/root/.london-delivery/a.json',new_name:'b.json'},f.context),/user document/);
});
test('uncertain rename is not retried after rebuilding request context',async()=>{
 const f=fixture(),path='/root/old.pdf';let writes=0;
 f.context.channelState={files:new Map([[path,{id:'file1'}]]),reads:new Map(),bytes:0};
 f.context.dropbox.getFileMetadata=async()=>({id:'file1','.tag':'file'});
 f.context.dropbox.renameFile=async()=>{writes++;throw Error('timeout');};
 await assert.rejects(runVoiceTool('rename_dropbox_file',{path,new_name:'new.pdf'},f.context),/timeout/);
 await assert.rejects(runVoiceTool('rename_dropbox_file',{path,new_name:'new.pdf'},f.context),/already attempted/);assert.equal(writes,1);
});

test('phone saves actual PDF/Word/Excel report paths using existing renderer and durable claim',async()=>{
 const f=fixture();let saves=0;f.context.dropbox.saveReports=true;
 f.context.dropbox.saveReport=async args=>{saves++;assert.equal(args.includeDocx,true);assert.equal(args.includeXlsx,true);return {id:'report1',path:'/root/a.pdf',docxPath:'/root/a.docx',xlsxPath:'/root/a.xlsx'};};
 const args={subject:'Quote comparison',text:'Verified source comparison',formats:['pdf','docx','xlsx']};
 const result=await runVoiceTool('save_analysis_report',args,f.context);
 assert.equal(result.paths.length,3);assert.match(result.receipt,/PDF, Word, Excel/);
 await assert.rejects(runVoiceTool('save_analysis_report',args,f.context),/already attempted/);assert.equal(saves,1);
});
test('report saving disabled, missing output and partial failure never claim success',async()=>{
 const f=fixture(),args={subject:'Review',text:'Content',formats:['docx']};
 await assert.rejects(runVoiceTool('save_analysis_report',args,f.context),/not configured/);
 f.context.dropbox.saveReports=true;f.context.dropbox.saveReport=async()=>({id:'pdf',path:'/root/a.pdf'});
 await assert.rejects(runVoiceTool('save_analysis_report',args,f.context),/not verified/);
 await assert.rejects(runVoiceTool('save_analysis_report',args,f.context),/already attempted/);
});
test('SMS saves requested report and returns verified concise receipt',async()=>{
 const f=fixture();f.context.dropbox.saveReports=true;let saves=0;
 f.context.dropbox.saveReport=async()=>{saves++;return {id:'r1',path:'/root/a.pdf',xlsxPath:'/root/a.xlsx'};};
 const openai={respond:async r=>{assert.ok(r.tools.some(t=>t.name==='save_analysis_report'));return {raw:{output:[{type:'function_call',call_id:'save',name:'save_analysis_report',arguments:JSON.stringify({subject:'Quote review',text:'Owner supplied figures',formats:['xlsx']})}]}};}};
 assert.match(await answerOwnerSms({...f.context,openai,requestKey:'sms-report',body:'Save an Excel report of the quote review'}),/Saved Excel report/);assert.equal(saves,1);
});

test('SMS action clarification keeps owner intent while a new unrelated question drops write tools',async()=>{
 const f=fixture();f.context.graph.createFollowUp=async args=>({id:'task1',title:args.title,date:args.date,reminder:'None'});
 const body='Create a task to review the quote without a reminder';
 const question=await answerOwnerSms({...f.context,requestKey:'ask',body,openai:{respond:async()=>({text:'What due date should I use?'})}});
 assert.match(question,/^Action needs details:/);
 const history=[{role:'user',content:body},{role:'assistant',content:question}];
 const openai={respond:async r=>{assert.ok(r.tools.some(t=>t.name==='create_action_register'));return {raw:{output:[{type:'function_call',call_id:'save',name:'create_action_register',arguments:JSON.stringify({title:'Review quote',date:'2035-07-11',reminder:false})}]}};}};
 assert.match(await answerOwnerSms({...f.context,requestKey:'answer',body:'July 11 2035',history,openai}),/Created in London/);
 await answerOwnerSms({...f.context,requestKey:'unrelated',body:'What is in my inbox?',history,openai:{respond:async r=>{assert.ok(!r.tools.some(t=>t.name==='create_action_register'));return {text:'No messages found.'};}}});
});
test('SMS fabricated action success and provider failures cannot be reported as completed',async()=>{
 const f=fixture(),body='Create a task to review quote July 11 2035 without a reminder';
 const reply=await answerOwnerSms({...f.context,requestKey:'fake',body,openai:{respond:async()=>({text:'Created the task.'})}});
 assert.match(reply,/Nothing was changed/);assert.doesNotMatch(reply,/Created the task/);
 f.context.graph.createFollowUp=async()=>{throw Error('timeout');};
 const openai={respond:async()=>({raw:{output:[{type:'function_call',call_id:'save',name:'create_action_register',arguments:JSON.stringify({title:'Review',date:'2035-07-11',reminder:false})}]}})};
 assert.match(await answerOwnerSms({...f.context,requestKey:'failed',body,openai}),/not verified/);
});
