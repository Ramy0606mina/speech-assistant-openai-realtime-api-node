import test from 'node:test';
import assert from 'node:assert/strict';
import mammoth from 'mammoth';
import { parseReport, renderReportHtml, reportSubject } from '../src/report-format.js';
import { renderReportDocx } from '../src/report-docx.js';
import { renderReportPdf } from '../src/report-pdf.js';
import { DropboxClient } from '../src/dropbox-client.js';
import { LondonCore } from '../src/london-core.js';
import { StateStore } from '../src/state-store.js';
import { MicrosoftGraphClient } from '../src/microsoft-graph.js';
import { OpenAIClient } from '../src/openai-client.js';

const text = '# Contractor comparison\n\n| Item | Contractor A | Contractor B |\n|---|---:|---:|\n| Price | **$100.00** | $200.00 |\n| Scope | Included | Not specified |';

test('tables render as safe HTML instead of visible Markdown, including escaped email markup',()=>{
  const escaped=text.replace(/([#*|])/g,'\\$1').replace('Included','Included&#x20;');
  const html=renderReportHtml(escaped);
  assert.match(html,/<table /);assert.match(html,/<th /);assert.match(html,/>\$100.00<\/td>/);
  assert.doesNotMatch(html,/\\|\*\*|&#x20;|\|---/);
  const unsafe=renderReportHtml(text.replace('Included','<img src=x onerror=alert(1)>'));
  assert.doesNotMatch(unsafe,/<img/);assert.match(unsafe,/&lt;img/);
  assert.equal(parseReport(text)[1].rows.length,2);
});
test('bad row widths fail rather than shift amounts under another contractor',()=>{
  assert.throws(()=>parseReport(text+'\n| Broken | one | two | extra |'),/inconsistent columns/);
});
test('reply subject does not accumulate completion prefixes',()=>{
  assert.equal(reportSubject('RE: LONDON — Task Complete | RE: LONDON — Task Complete | Compare bids'),'Compare bids');
});
test('Word is a readable document with native table cells and all source values',async()=>{
  const buffer=renderReportDocx({subject:'Compare bids',text});
  assert.equal(buffer.subarray(0,2).toString(),'PK');
  const result=await mammoth.convertToHtml({buffer});
  assert.match(result.value,/<table>/);assert.match(result.value,/Contractor A/);assert.match(result.value,/\$100.00/);assert.match(result.value,/Not specified/);
});
test('PDF renders long comparisons over multiple pages without truncating rows',async()=>{
  const long=text+'\n'+Array.from({length:60},(_,i)=>`| Item ${i} | Included with a lengthy scope description that wraps naturally | Not specified |`).join('\n');
  const pdf=await renderReportPdf({subject:'Compare bids',text:long});
  assert.equal(pdf.subarray(0,5).toString(),'%PDF-');
  assert.ok((pdf.toString().match(/\/Type \/Page\b/g)||[]).length>1);
});
test('comparison saves both genuine PDF and Word before exposing confirmed attachments',async()=>{
  const uploads=[];
  const dbx=new DropboxClient({accessToken:'test',saveReports:true,fetchImpl:async(url,opts)=>{
    if(String(url).endsWith('create_folder_v2'))return Response.json({id:'folder'});
    const path=JSON.parse(opts.headers['Dropbox-API-Arg']).path;uploads.push({path,bytes:opts.body});
    return Response.json({id:'saved',path_display:path});
  }});
  const saved=await dbx.saveReport({taskKey:'test',subject:'Compare bids',text});
  assert.equal(uploads.length,2);assert.ok(saved.docxPath.endsWith('.docx'));assert.equal(saved.attachments.length,2);
  assert.equal(uploads[0].bytes.subarray(0,5).toString(),'%PDF-');assert.equal(uploads[1].bytes.subarray(0,2).toString(),'PK');
});
test('owner comparison is delivered as HTML with the generated attachments and no third-party mail',async()=>{
  let sent;
  const attachments=[{name:'comparison.pdf',contentType:'application/pdf',contentBytes:'JVBERg=='}];
  const core=new LondonCore({state:new StateStore(),graph:{principalMailbox:'owner@example.com',readMailbox:'london@example.com',getLondonMessage:async()=>({from:{emailAddress:{address:'owner@example.com'}},subject:'Compare bids'}),sendMail:async mail=>{sent=mail;}},openai:{analyzeDelegatedEmail:async()=>({text})},dropbox:{saveReports:true,saveReport:async()=>({path:'/London Work/report.pdf',docxPath:'/London Work/report.docx',attachments})}});
  await core.processMessage({id:'format-test'});
  assert.equal(sent.contentType,'HTML');assert.match(sent.body,/<table /);assert.deepEqual(sent.attachments,attachments);assert.equal(sent.to,'owner@example.com');
});
test('unconfirmed Word upload prevents completion delivery',async()=>{
  const dbx=new DropboxClient({accessToken:'test',saveReports:true,fetchImpl:async(url,opts)=>{
    if(String(url).endsWith('create_folder_v2'))return Response.json({id:'folder'});
    const path=JSON.parse(opts.headers['Dropbox-API-Arg']).path;
    return Response.json(path.endsWith('.docx')?{}:{id:'pdf',path_display:path});
  }});
  await assert.rejects(dbx.saveReport({taskKey:'failed',subject:'Compare bids',text}),/Word report/);
});
test('Microsoft receives native file attachments and continues to reject external recipients',async()=>{
  const calls=[];
  const graph=new MicrosoftGraphClient({readTenantId:'tenant',readClientId:'client',readClientSecret:'test',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async(url,options)=>{
    if(String(url).includes('login.microsoftonline.com'))return Response.json({access_token:'test',expires_in:3600});
    calls.push(JSON.parse(options.body));return new Response(null,{status:202});
  }});
  const attachment={name:'report.pdf',contentType:'application/pdf',contentBytes:Buffer.from('%PDF-test').toString('base64')};
  await graph.sendMail({to:'owner@example.com',subject:'Comparison',body:renderReportHtml(text),contentType:'HTML',attachments:[attachment]});
  assert.equal(calls[0].message.attachments[0]['@odata.type'],'#microsoft.graph.fileAttachment');
  assert.equal(calls[0].message.attachments[0].contentBytes,attachment.contentBytes);
  assert.equal(calls[0].message.body.contentType,'HTML');
  await assert.rejects(graph.sendMail({to:'other@example.com',attachments:[attachment]}),/restricted/);
  await assert.rejects(graph.sendMail({to:'owner@example.com',attachments:[{...attachment,contentBytes:'x'.repeat(3*1024*1024+1)}]}),/limit/);
  assert.equal(calls.length,1);
});
test('malformed model tables get one correction attempt before delivery',async()=>{
  let calls=0;
  const client=new OpenAIClient({apiKey:'test',fetchImpl:async()=>Response.json({output_text:++calls===1?text+'\n| bad | extra | cells | here |':text})});
  const result=await client.analyzeDelegatedEmail({subject:'Compare'},[]);
  assert.equal(calls,2);assert.equal(result.text,text);
  const broken=new OpenAIClient({apiKey:'test',fetchImpl:async()=>Response.json({output_text:'| missing | separator |'})});
  await assert.rejects(broken.analyzeDelegatedEmail({subject:'Compare'},[]),/separator/);
});
